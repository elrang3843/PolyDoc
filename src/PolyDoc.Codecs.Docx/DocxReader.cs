using DocumentFormat.OpenXml.Packaging;
using PolyDoc.Core;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace PolyDoc.Codecs.Docx;

/// <summary>
/// DOCX (OOXML WordprocessingML) → PolyDocument 리더.
///
/// Phase C 1차 사이클 매핑:
///   - 단락 / 인라인 런 / 제목(Heading1~6) / 정렬 / 기본 리스트
///   - 굵게·기울임·밑줄·취소선·위첨자·아래첨자·폰트·크기·색상
/// 표·이미지·필드·각주 등은 후속 사이클에서 추가한다.
/// </summary>
public sealed class DocxReader : IDocumentReader
{
    public string FormatId => "docx";

    public PolyDocument Read(Stream input)
    {
        ArgumentNullException.ThrowIfNull(input);

        using var package = WordprocessingDocument.Open(input, isEditable: false);
        var mainPart = package.MainDocumentPart
            ?? throw new InvalidDataException("DOCX package has no main document part.");
        var body = mainPart.Document?.Body
            ?? throw new InvalidDataException("DOCX main document has no body.");

        var document = new PolyDocument();
        var section = new Section();
        document.Sections.Add(section);

        foreach (var element in body.ChildElements)
        {
            switch (element)
            {
                case W.Paragraph wpara:
                    section.Blocks.Add(ReadParagraph(wpara));
                    break;
                // W.Table 은 후속 사이클.
            }
        }

        ReadCoreProperties(package, document.Metadata);
        return document;
    }

    private static Paragraph ReadParagraph(W.Paragraph wp)
    {
        var paragraph = new Paragraph();
        ApplyParagraphProperties(paragraph, wp.ParagraphProperties);

        foreach (var run in wp.Elements<W.Run>())
        {
            var text = string.Concat(run.Elements<W.Text>().Select(t => t.Text));
            if (text.Length == 0)
            {
                continue;
            }
            paragraph.AddText(text, ReadRunStyle(run.RunProperties));
        }

        if (paragraph.Runs.Count == 0)
        {
            paragraph.AddText(string.Empty);
        }
        return paragraph;
    }

    private static void ApplyParagraphProperties(Paragraph paragraph, W.ParagraphProperties? pPr)
    {
        if (pPr is null)
        {
            return;
        }

        // pStyle = "Heading1" .. "Heading6" → OutlineLevel
        var styleId = pPr.ParagraphStyleId?.Val?.Value;
        if (!string.IsNullOrEmpty(styleId)
            && styleId.StartsWith("Heading", StringComparison.OrdinalIgnoreCase)
            && int.TryParse(styleId.AsSpan("Heading".Length), out var level)
            && level is >= 1 and <= 6)
        {
            paragraph.Style.Outline = (OutlineLevel)level;
        }

        // <w:outlineLvl w:val="0"/> 가 별도로 지정된 경우 (0-기반).
        if (paragraph.Style.Outline == OutlineLevel.Body
            && pPr.OutlineLevel?.Val?.Value is int rawLevel
            && rawLevel + 1 is var oneBased and >= 1 and <= 6)
        {
            paragraph.Style.Outline = (OutlineLevel)oneBased;
        }

        // 정렬
        if (pPr.Justification?.Val?.Value is { } jc)
        {
            // OpenXml 3.x 의 JustificationValues 는 IEnumValue 구조체라 switch 패턴에 못 쓴다 → equality.
            if (jc.Equals(W.JustificationValues.Center))
            {
                paragraph.Style.Alignment = Alignment.Center;
            }
            else if (jc.Equals(W.JustificationValues.Right))
            {
                paragraph.Style.Alignment = Alignment.Right;
            }
            else if (jc.Equals(W.JustificationValues.Both))
            {
                paragraph.Style.Alignment = Alignment.Justify;
            }
            else if (jc.Equals(W.JustificationValues.Distribute))
            {
                paragraph.Style.Alignment = Alignment.Distributed;
            }
            else
            {
                paragraph.Style.Alignment = Alignment.Left;
            }
        }

        // 기본 리스트 표식 — Phase C 1차 사이클에선 numId 별 종류 구분 없이 bullet 으로 격하.
        if (pPr.NumberingProperties is not null)
        {
            paragraph.Style.ListMarker = new ListMarker { Kind = ListKind.Bullet };
        }
    }

    private static RunStyle ReadRunStyle(W.RunProperties? rPr)
    {
        var style = new RunStyle();
        if (rPr is null)
        {
            return style;
        }

        if (rPr.Bold is not null)
        {
            style.Bold = rPr.Bold.Val is null || rPr.Bold.Val.Value;
        }
        if (rPr.Italic is not null)
        {
            style.Italic = rPr.Italic.Val is null || rPr.Italic.Val.Value;
        }
        if (rPr.Underline is not null && !(rPr.Underline.Val?.Value.Equals(W.UnderlineValues.None) ?? false))
        {
            style.Underline = true;
        }
        if (rPr.Strike is not null)
        {
            style.Strikethrough = rPr.Strike.Val is null || rPr.Strike.Val.Value;
        }

        if (rPr.FontSize?.Val?.Value is { } sizeStr
            && double.TryParse(sizeStr, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var halfPoints))
        {
            style.FontSizePt = halfPoints / 2.0;
        }

        if (rPr.RunFonts?.Ascii?.Value is { Length: > 0 } fontAscii)
        {
            style.FontFamily = fontAscii;
        }
        else if (rPr.RunFonts?.EastAsia?.Value is { Length: > 0 } fontEa)
        {
            style.FontFamily = fontEa;
        }

        if (rPr.Color?.Val?.Value is { Length: 6 } colorHex)
        {
            try
            {
                style.Foreground = Color.FromHex(colorHex);
            }
            catch (FormatException)
            {
                // 잘못된 색상 표기는 무시.
            }
        }

        if (rPr.VerticalTextAlignment?.Val?.Value is { } vert)
        {
            if (vert.Equals(W.VerticalPositionValues.Superscript))
            {
                style.Superscript = true;
            }
            else if (vert.Equals(W.VerticalPositionValues.Subscript))
            {
                style.Subscript = true;
            }
        }

        return style;
    }

    private static void ReadCoreProperties(WordprocessingDocument package, DocumentMetadata metadata)
    {
        var props = package.PackageProperties;
        if (!string.IsNullOrEmpty(props.Title))
        {
            metadata.Title = props.Title;
        }
        if (!string.IsNullOrEmpty(props.Creator))
        {
            metadata.Author = props.Creator;
        }
        if (props.Created is { } created)
        {
            metadata.Created = new DateTimeOffset(created.ToUniversalTime(), TimeSpan.Zero);
        }
        if (props.Modified is { } modified)
        {
            metadata.Modified = new DateTimeOffset(modified.ToUniversalTime(), TimeSpan.Zero);
        }
    }
}

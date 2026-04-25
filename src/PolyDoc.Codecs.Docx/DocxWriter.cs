using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using PolyDoc.Core;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace PolyDoc.Codecs.Docx;

/// <summary>
/// PolyDocument → DOCX (OOXML WordprocessingML) 라이터.
///
/// Phase C 1차 사이클 매핑 (DocxReader 와 대칭):
///   - 단락 / 인라인 런 / 제목(Heading1~6) / 정렬 / 기본 리스트
///   - 굵게·기울임·밑줄·취소선·위첨자·아래첨자·폰트·크기·색상
/// </summary>
public sealed class DocxWriter : IDocumentWriter
{
    public string FormatId => "docx";

    public void Write(PolyDocument document, Stream output)
    {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(output);

        using var package = WordprocessingDocument.Create(output, WordprocessingDocumentType.Document);
        var mainPart = package.AddMainDocumentPart();
        mainPart.Document = new W.Document();
        var body = new W.Body();
        mainPart.Document.AppendChild(body);

        // 표준 Heading 스타일 정의 (그렇지 않으면 Word 가 기본 스타일을 적용한다).
        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles = BuildStyles();

        foreach (var paragraph in document.EnumerateParagraphs())
        {
            body.AppendChild(BuildParagraph(paragraph));
        }

        // 섹션 속성 (페이지 설정).
        var firstSection = document.Sections.FirstOrDefault();
        body.AppendChild(BuildSectionProperties(firstSection?.Page ?? new PageSettings()));

        WriteCoreProperties(package, document.Metadata);
    }

    private static W.Paragraph BuildParagraph(Paragraph p)
    {
        var wpara = new W.Paragraph();
        var pPr = new W.ParagraphProperties();

        if (p.Style.Outline is var lvl and > OutlineLevel.Body)
        {
            pPr.ParagraphStyleId = new W.ParagraphStyleId { Val = $"Heading{(int)lvl}" };
        }

        var alignment = ToJustification(p.Style.Alignment);
        if (alignment is not null)
        {
            pPr.Justification = new W.Justification { Val = alignment };
        }

        if (p.Style.ListMarker is not null)
        {
            // Phase C 1차 사이클: numId/abstractNumId 정의 없이 표시만 — Word 가 기본 글머리 적용.
            pPr.NumberingProperties = new W.NumberingProperties(
                new W.NumberingLevelReference { Val = p.Style.ListMarker.Level },
                new W.NumberingId { Val = 1 });
        }

        if (pPr.HasChildren)
        {
            wpara.AppendChild(pPr);
        }

        foreach (var run in p.Runs)
        {
            wpara.AppendChild(BuildRun(run));
        }
        return wpara;
    }

    private static W.Run BuildRun(Run run)
    {
        var wrun = new W.Run();
        var rPr = new W.RunProperties();

        if (run.Style.Bold)
        {
            rPr.Bold = new W.Bold();
        }
        if (run.Style.Italic)
        {
            rPr.Italic = new W.Italic();
        }
        if (run.Style.Underline)
        {
            rPr.Underline = new W.Underline { Val = W.UnderlineValues.Single };
        }
        if (run.Style.Strikethrough)
        {
            rPr.Strike = new W.Strike();
        }
        if (run.Style.Superscript)
        {
            rPr.VerticalTextAlignment = new W.VerticalTextAlignment { Val = W.VerticalPositionValues.Superscript };
        }
        if (run.Style.Subscript)
        {
            rPr.VerticalTextAlignment = new W.VerticalTextAlignment { Val = W.VerticalPositionValues.Subscript };
        }

        if (Math.Abs(run.Style.FontSizePt - 11) > 0.001)
        {
            // DOCX 의 sz 값은 half-point 단위.
            var halfPoints = ((int)Math.Round(run.Style.FontSizePt * 2)).ToString(CultureInfo.InvariantCulture);
            rPr.FontSize = new W.FontSize { Val = halfPoints };
        }

        if (!string.IsNullOrEmpty(run.Style.FontFamily))
        {
            rPr.RunFonts = new W.RunFonts
            {
                Ascii = run.Style.FontFamily,
                HighAnsi = run.Style.FontFamily,
                EastAsia = run.Style.FontFamily,
            };
        }

        if (run.Style.Foreground is { } fg)
        {
            rPr.Color = new W.Color { Val = $"{fg.R:X2}{fg.G:X2}{fg.B:X2}" };
        }

        if (rPr.HasChildren)
        {
            wrun.AppendChild(rPr);
        }

        // <w:t xml:space="preserve"> — 양 끝 공백 보존.
        wrun.AppendChild(new W.Text(run.Text) { Space = SpaceProcessingModeValues.Preserve });
        return wrun;
    }

    private static EnumValue<W.JustificationValues>? ToJustification(Alignment alignment) => alignment switch
    {
        Alignment.Center => W.JustificationValues.Center,
        Alignment.Right => W.JustificationValues.Right,
        Alignment.Justify => W.JustificationValues.Both,
        Alignment.Distributed => W.JustificationValues.Distribute,
        Alignment.Left => null,        // 기본값 → 명시 안 함
        _ => null,
    };

    private static W.SectionProperties BuildSectionProperties(PageSettings page)
    {
        // DOCX 는 트위프(twentieth of a point, 1/1440 inch). mm → twips.
        static uint MmToTwips(double mm) => (uint)Math.Round(mm * 56.6929);

        var props = new W.SectionProperties();
        props.AppendChild(new W.PageSize
        {
            Width = MmToTwips(page.WidthMm),
            Height = MmToTwips(page.HeightMm),
            Orient = page.Orientation == PageOrientation.Landscape
                ? W.PageOrientationValues.Landscape
                : W.PageOrientationValues.Portrait,
        });
        props.AppendChild(new W.PageMargin
        {
            Top = (int)MmToTwips(page.MarginTopMm),
            Right = MmToTwips(page.MarginRightMm),
            Bottom = (int)MmToTwips(page.MarginBottomMm),
            Left = MmToTwips(page.MarginLeftMm),
            Header = 720,
            Footer = 720,
            Gutter = 0,
        });
        return props;
    }

    private static W.Styles BuildStyles()
    {
        var styles = new W.Styles();

        // 기본 단락 / 기본 폰트.
        styles.AppendChild(new W.DocDefaults(
            new W.RunPropertiesDefault(new W.RunPropertiesBaseStyle(
                new W.FontSize { Val = "22" })),     // 11pt
            new W.ParagraphPropertiesDefault(new W.ParagraphPropertiesBaseStyle())));

        // Heading 1 ~ 6 — 최소 정의 (Word 의 내장 스타일과 같은 ID 를 쓰면 사용자 환경에서 정상 표시).
        for (int i = 1; i <= 6; i++)
        {
            var headingStyle = new W.Style
            {
                Type = W.StyleValues.Paragraph,
                StyleId = $"Heading{i}",
            };
            headingStyle.AppendChild(new W.StyleName { Val = $"heading {i}" });
            headingStyle.AppendChild(new W.StyleParagraphProperties(
                new W.OutlineLevel { Val = i - 1 }));
            headingStyle.AppendChild(new W.StyleRunProperties(
                new W.Bold(),
                new W.FontSize { Val = ((int)Math.Round((double)((20 - (i - 1) * 2) * 2))).ToString(CultureInfo.InvariantCulture) }));
            styles.AppendChild(headingStyle);
        }

        return styles;
    }

    private static void WriteCoreProperties(WordprocessingDocument package, DocumentMetadata metadata)
    {
        var props = package.PackageProperties;
        if (!string.IsNullOrEmpty(metadata.Title))
        {
            props.Title = metadata.Title;
        }
        if (!string.IsNullOrEmpty(metadata.Author))
        {
            props.Creator = metadata.Author;
        }
        props.Created = metadata.Created.UtcDateTime;
        props.Modified = metadata.Modified.UtcDateTime;
    }
}

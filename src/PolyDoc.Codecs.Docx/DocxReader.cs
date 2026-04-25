using DocumentFormat.OpenXml.Packaging;
using PolyDoc.Core;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using W = DocumentFormat.OpenXml.Wordprocessing;
using WP = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace PolyDoc.Codecs.Docx;

/// <summary>
/// DOCX (OOXML WordprocessingML) → PolyDocument 리더.
///
/// 매핑 범위:
///   - 단락 / 인라인 런 / 제목(Heading1~6) / 정렬 / 기본 리스트
///   - 굵게·기울임·밑줄·취소선·위첨자·아래첨자·폰트·크기·색상
///   - 표 (w:tbl → Table, 셀 병합 포함)
///   - 인라인 이미지 (w:drawing → ImageBlock, ImagePart 바이너리 추출)
///   - 미인식 블록 (도형·SDT 등) → OpaqueBlock 으로 원본 XML 보존
///
/// 각주·필드·고급 표 속성은 후속 사이클에서 추가한다.
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

        var ctx = new ReadContext(mainPart);

        foreach (var element in body.ChildElements)
        {
            switch (element)
            {
                case W.Paragraph wpara:
                    AppendParagraphAndExtractedDrawings(section.Blocks, wpara, ctx);
                    break;
                case W.Table wtable:
                    section.Blocks.Add(ReadTable(wtable, ctx));
                    break;
                case W.SectionProperties:
                    // 섹션 속성은 페이지 설정으로 — 다음 사이클에서 처리.
                    break;
                default:
                    // 미인식 블록은 보존 섬으로 들어온다.
                    section.Blocks.Add(new OpaqueBlock
                    {
                        Format = "docx",
                        Kind = element.LocalName,
                        Xml = element.OuterXml,
                        DisplayLabel = $"[보존된 {element.LocalName}]",
                    });
                    break;
            }
        }

        ReadCoreProperties(package, document.Metadata);
        return document;
    }

    private sealed class ReadContext
    {
        public ReadContext(MainDocumentPart mainPart)
        {
            MainPart = mainPart;
        }

        public MainDocumentPart MainPart { get; }
    }

    private static void AppendParagraphAndExtractedDrawings(IList<Block> target, W.Paragraph wp, ReadContext ctx)
    {
        var paragraph = new Paragraph();
        ApplyParagraphProperties(paragraph, wp.ParagraphProperties);

        var images = new List<ImageBlock>();

        foreach (var run in wp.Elements<W.Run>())
        {
            // 1) 인라인 이미지(w:drawing) 가 있으면 추출해 별도 ImageBlock 으로 등록한다.
            //    해당 Run 의 텍스트는 비어 있을 가능성이 크지만, 텍스트와 그림이 섞여 있어도 텍스트는 단락에,
            //    그림은 단락 뒤 ImageBlock 으로 들어가는 단순한 1차 정책.
            foreach (var drawing in run.Elements<W.Drawing>())
            {
                if (TryExtractImage(drawing, ctx, out var imageBlock))
                {
                    images.Add(imageBlock);
                }
                else
                {
                    target.Add(new OpaqueBlock
                    {
                        Format = "docx",
                        Kind = "drawing",
                        Xml = drawing.OuterXml,
                        DisplayLabel = "[보존된 도형]",
                    });
                }
            }

            // 2) 텍스트 — 빈 텍스트도 RunProperties 의 의미를 살리고 싶지만 현 사이클에서는 무시.
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
        target.Add(paragraph);
        foreach (var img in images)
        {
            target.Add(img);
        }
    }

    private static bool TryExtractImage(W.Drawing drawing, ReadContext ctx, out ImageBlock imageBlock)
    {
        imageBlock = null!;
        var blip = drawing.Descendants<A.Blip>().FirstOrDefault();
        if (blip?.Embed?.Value is not { Length: > 0 } relId)
        {
            return false;
        }

        if (ctx.MainPart.GetPartById(relId) is not ImagePart imagePart)
        {
            return false;
        }

        byte[] bytes;
        using (var stream = imagePart.GetStream())
        using (var ms = new MemoryStream())
        {
            stream.CopyTo(ms);
            bytes = ms.ToArray();
        }

        // 사이즈 — wp:extent (EMU 단위, 1 EMU = 1/914400 inch)
        double widthMm = 0;
        double heightMm = 0;
        var extent = drawing.Descendants<WP.Extent>().FirstOrDefault();
        if (extent is not null)
        {
            widthMm = EmuToMm(extent.Cx?.Value ?? 0);
            heightMm = EmuToMm(extent.Cy?.Value ?? 0);
        }

        // alt text
        var docProps = drawing.Descendants<WP.DocProperties>().FirstOrDefault();
        var description = docProps?.Description?.Value ?? docProps?.Name?.Value;

        imageBlock = new ImageBlock
        {
            MediaType = imagePart.ContentType,
            Data = bytes,
            WidthMm = widthMm,
            HeightMm = heightMm,
            Description = description,
        };
        return true;
    }

    private static double EmuToMm(long emu) => emu / 914400.0 * 25.4;

    private static Table ReadTable(W.Table wtable, ReadContext ctx)
    {
        var table = new Table();

        // 컬럼 너비 (twips → mm). 1 twip = 1/1440 inch.
        var grid = wtable.Elements<W.TableGrid>().FirstOrDefault();
        if (grid is not null)
        {
            foreach (var col in grid.Elements<W.GridColumn>())
            {
                var widthTwips = col.Width?.Value;
                table.Columns.Add(new TableColumn
                {
                    WidthMm = ParseTwipsToMm(widthTwips),
                });
            }
        }

        foreach (var row in wtable.Elements<W.TableRow>())
        {
            var tableRow = new TableRow();
            var rowHeight = row.Elements<W.TableRowProperties>()
                .SelectMany(rp => rp.Elements<W.TableRowHeight>())
                .FirstOrDefault();
            if (rowHeight?.Val?.Value is uint h)
            {
                tableRow.HeightMm = h / 56.6929;        // twips → mm
            }

            foreach (var cell in row.Elements<W.TableCell>())
            {
                var tcPr = cell.Elements<W.TableCellProperties>().FirstOrDefault();
                var span = tcPr?.GridSpan?.Val?.Value ?? 1;
                var vMerge = tcPr?.VerticalMerge;
                var widthTwips = tcPr?.TableCellWidth?.Width?.Value;

                // vMerge 의 continue 셀(merged 로 사라진 자리)은 sparse 표현을 따라 추가 안 함.
                if (vMerge is not null && (vMerge.Val?.Value is null
                    || vMerge.Val.Value.Equals(W.MergedCellValues.Continue)))
                {
                    continue;
                }

                var tableCell = new TableCell
                {
                    ColumnSpan = span,
                };
                if (vMerge?.Val?.Value is { } mergeVal && mergeVal.Equals(W.MergedCellValues.Restart))
                {
                    // 시작 셀 — 실제 RowSpan 은 후속 행을 보고 확정해야 하지만 1차 사이클은 1 로 두고 다음 사이클에서 정밀화.
                    tableCell.RowSpan = 1;
                }
                if (widthTwips is { } w && int.TryParse(w, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out var twips))
                {
                    tableCell.WidthMm = twips / 56.6929;
                }

                foreach (var inner in cell.ChildElements)
                {
                    switch (inner)
                    {
                        case W.Paragraph paraInCell:
                            AppendParagraphAndExtractedDrawings(tableCell.Blocks, paraInCell, ctx);
                            break;
                        case W.Table nested:
                            tableCell.Blocks.Add(ReadTable(nested, ctx));
                            break;
                    }
                }

                if (tableCell.Blocks.Count == 0)
                {
                    tableCell.Blocks.Add(Paragraph.Of(string.Empty));
                }

                tableRow.Cells.Add(tableCell);
            }

            table.Rows.Add(tableRow);
        }

        return table;
    }

    private static double ParseTwipsToMm(string? twipsRaw)
    {
        if (string.IsNullOrEmpty(twipsRaw))
        {
            return 0;
        }
        return double.TryParse(twipsRaw, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var twips)
            ? twips / 56.6929
            : 0;
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

        if (paragraph.Style.Outline == OutlineLevel.Body
            && pPr.OutlineLevel?.Val?.Value is int rawLevel
            && rawLevel + 1 is var oneBased and >= 1 and <= 6)
        {
            paragraph.Style.Outline = (OutlineLevel)oneBased;
        }

        if (pPr.Justification?.Val?.Value is { } jc)
        {
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

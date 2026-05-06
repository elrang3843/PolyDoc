using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using PolyDonky.Core;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using W = DocumentFormat.OpenXml.Wordprocessing;
using WP = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace PolyDonky.Codecs.Docx;

/// <summary>
/// DOCX (OOXML WordprocessingML) → PolyDonkyument 리더.
///
/// 매핑 범위:
///   - 단락 / 인라인 런 / 제목(Heading1~6) / 정렬 / 기본 리스트
///   - 굵게·기울임·밑줄·취소선·위첨자·아래첨자·폰트·크기·색상
///   - 표 (w:tbl → Table, 셀 병합 포함)
///   - 인라인 이미지 (w:drawing → ImageBlock, ImagePart 바이너리 추출)
///   - DrawingML 도형 (wps:wsp → ShapeObject: rect/ellipse/triangle/polygon/polyline/line 등)
///   - 미인식 블록 (SDT 등) → OpaqueBlock 으로 원본 XML 보존
///
/// 각주·필드·고급 표 속성은 후속 사이클에서 추가한다.
/// </summary>
public sealed class DocxReader : IDocumentReader
{
    public string FormatId => "docx";

    public PolyDonkyument Read(Stream input)
    {
        ArgumentNullException.ThrowIfNull(input);

        using var package = WordprocessingDocument.Open(input, isEditable: false);
        var mainPart = package.MainDocumentPart
            ?? throw new InvalidDataException("DOCX package has no main document part.");
        var body = mainPart.Document?.Body
            ?? throw new InvalidDataException("DOCX main document has no body.");

        var document = new PolyDonkyument();
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
        ReadFootnotesAndEndnotes(mainPart, document, ctx);
        return document;
    }

    private static void ReadFootnotesAndEndnotes(MainDocumentPart mainPart, PolyDonkyument document, ReadContext ctx)
    {
        if (mainPart.FootnotesPart is { } fnPart && fnPart.Footnotes is not null)
        {
            foreach (var fn in fnPart.Footnotes.Elements<W.Footnote>())
            {
                var fnType = fn.Type?.Value;
                if (fnType == W.FootnoteEndnoteValues.Separator
                    || fnType == W.FootnoteEndnoteValues.ContinuationSeparator)
                    continue;
                var id = fn.Id?.ToString();
                if (string.IsNullOrEmpty(id) || id == "0" || id == "-1") continue;
                var entry = new FootnoteEntry { Id = id };
                foreach (var para in fn.Elements<W.Paragraph>())
                    AppendParagraphAndExtractedDrawings(entry.Blocks, para, ctx);
                if (entry.Blocks.Count == 0)
                    entry.Blocks.Add(new Paragraph());
                document.Footnotes.Add(entry);
            }
        }

        if (mainPart.EndnotesPart is { } enPart && enPart.Endnotes is not null)
        {
            foreach (var en in enPart.Endnotes.Elements<W.Endnote>())
            {
                var enType = en.Type?.Value;
                if (enType == W.FootnoteEndnoteValues.Separator
                    || enType == W.FootnoteEndnoteValues.ContinuationSeparator)
                    continue;
                var id = en.Id?.ToString();
                if (string.IsNullOrEmpty(id) || id == "0" || id == "-1") continue;
                var entry = new FootnoteEntry { Id = id };
                foreach (var para in en.Elements<W.Paragraph>())
                    AppendParagraphAndExtractedDrawings(entry.Blocks, para, ctx);
                if (entry.Blocks.Count == 0)
                    entry.Blocks.Add(new Paragraph());
                document.Endnotes.Add(entry);
            }
        }
    }

    private sealed class ReadContext
    {
        public ReadContext(MainDocumentPart mainPart)
        {
            MainPart = mainPart;
        }

        public MainDocumentPart MainPart { get; }
    }

    // DrawingML / WordprocessingShape 네임스페이스
    private static readonly XNamespace XnsA   = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace XnsWps = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
    private static readonly XNamespace XnsWp  = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    private static readonly XNamespace XnsWml = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static void AppendParagraphAndExtractedDrawings(IList<Block> target, W.Paragraph wp, ReadContext ctx)
    {
        var paragraph = new Paragraph();
        ApplyParagraphProperties(paragraph, wp.ParagraphProperties);

        // 그림·도형은 단락 뒤에 별도 블록으로 추가 (1차 정책).
        var drawingBlocks = new List<Block>();

        foreach (var run in wp.Elements<W.Run>())
        {
            // 각주·미주 참조 run 은 별도 처리 (텍스트·그림보다 우선 확인).
            if (run.Descendants<W.FootnoteReference>().FirstOrDefault() is { } fnRef
                && fnRef.Id?.Value is { } fnId)
            {
                paragraph.Runs.Add(new Run { FootnoteId = fnId.ToString(System.Globalization.CultureInfo.InvariantCulture) });
                continue;
            }
            if (run.Descendants<W.EndnoteReference>().FirstOrDefault() is { } enRef
                && enRef.Id?.Value is { } enId)
            {
                paragraph.Runs.Add(new Run { EndnoteId = enId.ToString(System.Globalization.CultureInfo.InvariantCulture) });
                continue;
            }

            foreach (var drawing in run.Elements<W.Drawing>())
            {
                if (TryExtractImage(drawing, ctx, out var imageBlock))
                {
                    drawingBlocks.Add(imageBlock);
                }
                else if (TryExtractShape(drawing, out var shapeBlock))
                {
                    drawingBlocks.Add(shapeBlock);
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
        foreach (var block in drawingBlocks)
        {
            target.Add(block);
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

    private static double EmuToMm(long emu) => UnitConverter.EmuToMm(emu);

    private static Table ReadTable(W.Table wtable, ReadContext ctx)
    {
        var table = new Table();

        // 표 속성 (테두리·배경)
        var tblPr = wtable.Elements<W.TableProperties>().FirstOrDefault();
        ReadTableProperties(table, tblPr);

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
                ReadCellProperties(tableCell, tcPr);

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

    // ── 표·셀 속성 (테두리·배경색) ────────────────────────────────────────────

    private static void ReadTableProperties(Table table, W.TableProperties? tblPr)
    {
        if (tblPr is null) return;

        // 표 외곽선: tblBorders > w:top 을 대표값으로 읽는다.
        var borders = tblPr.TableBorders;
        if (borders is not null)
        {
            var top = borders.TopBorder;
            if (top?.Val?.Value is { } tv && !tv.Equals(W.BorderValues.None)
                && top.Size?.Value is { } sz && sz > 0)
            {
                table.BorderThicknessPt = sz / 8.0;
            }
            if (top?.Color?.Value is { Length: 6 } tc && tc != "auto")
                table.BorderColor = "#" + tc.ToUpperInvariant();
        }

        // 표 배경색: tblPr > w:shd/@fill
        var shd = tblPr.Shading;
        if (shd?.Fill?.Value is { Length: 6 } fill && fill != "auto")
            table.BackgroundColor = "#" + fill.ToUpperInvariant();
    }

    private static void ReadCellProperties(TableCell cell, W.TableCellProperties? tcPr)
    {
        if (tcPr is null) return;

        var borders = tcPr.TableCellBorders;
        if (borders is not null)
        {
            cell.BorderTop    = ReadBorderSide(borders.TopBorder);
            cell.BorderBottom = ReadBorderSide(borders.BottomBorder);
            cell.BorderLeft   = ReadBorderSide(borders.LeftBorder);
            cell.BorderRight  = ReadBorderSide(borders.RightBorder);

            // 공통값(BorderThicknessPt/BorderColor)은 상단 면에서 대표값으로 유지 (하위 호환).
            if (cell.BorderTop is { } t)
            {
                if (t.ThicknessPt > 0) cell.BorderThicknessPt = t.ThicknessPt;
                if (t.Color is not null) cell.BorderColor = t.Color;
            }
        }

        // 셀 배경색: tcPr > w:shd/@fill
        var shd = tcPr.Shading;
        if (shd?.Fill?.Value is { Length: 6 } fill && fill != "auto")
            cell.BackgroundColor = "#" + fill.ToUpperInvariant();
    }

    private static CellBorderSide? ReadBorderSide(W.BorderType? border)
    {
        if (border is null) return null;
        if (border.Val?.Value is { } v && v.Equals(W.BorderValues.None)) return null;
        var pt = border.Size?.Value is { } sz && sz > 0 ? sz / 8.0 : 0;
        var color = border.Color?.Value is { Length: 6 } c && c != "auto" ? "#" + c.ToUpperInvariant() : null;
        if (pt <= 0 && color is null) return null;
        return new CellBorderSide(pt, color);
    }

    private static double ParseTwipsToMm(string? twipsRaw)
        => UnitConverter.ParseTwipsToMm(twipsRaw);

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

        if (pPr.PageBreakBefore is { } pbBefore
            && (pbBefore.Val is null || pbBefore.Val.Value))
        {
            paragraph.Style.ForcePageBreakBefore = true;
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

    // ── 도형 파싱 ─────────────────────────────────────────────────────────────

    private static bool TryExtractShape(W.Drawing drawing, out ShapeObject shape)
    {
        shape = null!;
        XElement xml;
        try { xml = XElement.Parse(drawing.OuterXml); }
        catch (System.Xml.XmlException) { return false; }

        // wps:wsp 가 있어야 DrawingML 도형
        var wsp = xml.Descendants(XnsWps + "wsp").FirstOrDefault();
        if (wsp is null) return false;

        var spPr = wsp.Element(XnsWps + "spPr");
        if (spPr is null) return false;

        shape = new ShapeObject();

        // ── 도형 종류 (geometry) ───────────────────────────────────────────────
        var prstGeom = spPr.Descendants(XnsA + "prstGeom").FirstOrDefault();
        var custGeom = spPr.Descendants(XnsA + "custGeom").FirstOrDefault();

        if (prstGeom is not null)
        {
            var prst = prstGeom.Attribute("prst")?.Value ?? "rect";
            shape.Kind = PrstToShapeKind(prst);
            if (shape.Kind is ShapeKind.RegularPolygon or ShapeKind.Star or ShapeKind.Triangle)
                shape.SideCount = PrstToSideCount(prst);
        }
        else if (custGeom is not null)
        {
            ParseCustomGeometry(custGeom, shape);
        }

        // ── 크기·회전 (xfrm) ─────────────────────────────────────────────────
        var xfrm = spPr.Descendants(XnsA + "xfrm").FirstOrDefault();
        if (xfrm is not null)
        {
            if (xfrm.Attribute("rot") is { } rotAttr
                && long.TryParse(rotAttr.Value, out var rotVal))
            {
                shape.RotationAngleDeg = rotVal / 60000.0;
            }

            var ext = xfrm.Element(XnsA + "ext");
            if (ext is not null)
            {
                if (long.TryParse(ext.Attribute("cx")?.Value, out var cxVal) && cxVal > 0)
                    shape.WidthMm  = EmuToMm(cxVal);
                if (long.TryParse(ext.Attribute("cy")?.Value, out var cyVal) && cyVal > 0)
                    shape.HeightMm = EmuToMm(cyVal);
            }
        }

        // ── 채우기 ────────────────────────────────────────────────────────────
        if (spPr.Descendants(XnsA + "noFill").Any())
        {
            shape.FillColor = null;
        }
        else
        {
            var solidFill = spPr.Descendants(XnsA + "solidFill").FirstOrDefault();
            if (solidFill is not null)
            {
                var srgbClr = solidFill.Element(XnsA + "srgbClr");
                if (srgbClr?.Attribute("val") is { } cv && cv.Value.Length == 6)
                    shape.FillColor = "#" + cv.Value.ToUpperInvariant();

                var alpha = srgbClr?.Element(XnsA + "alpha")
                    ?? solidFill.Descendants(XnsA + "alpha").FirstOrDefault();
                if (alpha?.Attribute("val") is { } av
                    && int.TryParse(av.Value, out var alphaVal))
                {
                    shape.FillOpacity = alphaVal / 100000.0;
                }
            }
        }

        // ── 선 (ln) ───────────────────────────────────────────────────────────
        var ln = spPr.Descendants(XnsA + "ln").FirstOrDefault();
        if (ln is not null)
        {
            if (ln.Descendants(XnsA + "noFill").Any())
            {
                shape.StrokeThicknessPt = 0;
            }
            else
            {
                if (ln.Attribute("w") is { } wAttr
                    && long.TryParse(wAttr.Value, out var wVal))
                {
                    shape.StrokeThicknessPt = wVal / 12700.0;
                }

                var lnSolid = ln.Descendants(XnsA + "solidFill").FirstOrDefault();
                if (lnSolid?.Element(XnsA + "srgbClr") is { } lnClr
                    && lnClr.Attribute("val") is { } lv && lv.Value.Length == 6)
                {
                    shape.StrokeColor = "#" + lv.Value.ToUpperInvariant();
                }

                var prstDash = ln.Descendants(XnsA + "prstDash").FirstOrDefault();
                if (prstDash?.Attribute("val") is { } dv)
                {
                    shape.StrokeDash = dv.Value switch
                    {
                        "dash" or "lgDash" or "lgDashDot" or "lgDashDotDot" or "sysDash" => StrokeDash.Dashed,
                        "dot" or "sysDot" => StrokeDash.Dotted,
                        "dashDot" or "sysDashDot" or "sysDashDotDot" => StrokeDash.DashDot,
                        _ => StrokeDash.Solid,
                    };
                }

                if (ln.Element(XnsA + "headEnd") is { } headEnd)
                    shape.StartArrow = ArrowTypeFromAttr(headEnd.Attribute("type")?.Value);
                if (ln.Element(XnsA + "tailEnd") is { } tailEnd)
                    shape.EndArrow = ArrowTypeFromAttr(tailEnd.Attribute("type")?.Value);
            }
        }

        // ── 위치·wrap (anchor / inline) ───────────────────────────────────────
        var anchor = xml.Descendants(XnsWp + "anchor").FirstOrDefault();
        var inline  = xml.Descendants(XnsWp + "inline").FirstOrDefault();

        if (anchor is not null)
        {
            ParseAnchorPosition(anchor, shape);
        }
        else if (inline is not null)
        {
            shape.WrapMode = ImageWrapMode.Inline;
            if (shape.WidthMm <= 0 || shape.HeightMm <= 0)
            {
                var ext = inline.Element(XnsWp + "extent");
                if (ext is not null)
                {
                    if (long.TryParse(ext.Attribute("cx")?.Value, out var ecx) && ecx > 0)
                        shape.WidthMm  = EmuToMm(ecx);
                    if (long.TryParse(ext.Attribute("cy")?.Value, out var ecy) && ecy > 0)
                        shape.HeightMm = EmuToMm(ecy);
                }
            }
        }

        // ── 레이블 (txbx) ─────────────────────────────────────────────────────
        var txbx = wsp.Descendants(XnsWps + "txbx").FirstOrDefault();
        if (txbx is not null)
        {
            var labelText = string.Concat(txbx.Descendants(XnsWml + "t").Select(t => t.Value));
            if (!string.IsNullOrEmpty(labelText))
                shape.LabelText = labelText;
        }

        return true;
    }

    private static void ParseAnchorPosition(XElement anchor, ShapeObject shape)
    {
        // behindDoc 속성
        var behindDoc = anchor.Attribute("behindDoc")?.Value;
        bool behind = behindDoc is "1" or "true";

        if (anchor.Descendants(XnsWp + "wrapNone").Any())
        {
            shape.WrapMode = behind ? ImageWrapMode.BehindText : ImageWrapMode.InFrontOfText;
        }
        else if (anchor.Descendants(XnsWp + "wrapSquare").Any()
              || anchor.Descendants(XnsWp + "wrapTopAndBottom").Any())
        {
            shape.WrapMode = ImageWrapMode.WrapLeft;
        }
        else
        {
            shape.WrapMode = ImageWrapMode.InFrontOfText;
        }

        if (anchor.Element(XnsWp + "positionH") is { } posH
            && posH.Element(XnsWp + "posOffset") is { } phOff
            && long.TryParse(phOff.Value, out var hOff))
        {
            shape.OverlayXMm = EmuToMm(hOff);
        }

        if (anchor.Element(XnsWp + "positionV") is { } posV
            && posV.Element(XnsWp + "posOffset") is { } pvOff
            && long.TryParse(pvOff.Value, out var vOff))
        {
            shape.OverlayYMm = EmuToMm(vOff);
        }

        var ext = anchor.Element(XnsWp + "extent");
        if (ext is not null)
        {
            if (long.TryParse(ext.Attribute("cx")?.Value, out var ecx) && ecx > 0)
                shape.WidthMm  = EmuToMm(ecx);
            if (long.TryParse(ext.Attribute("cy")?.Value, out var ecy) && ecy > 0)
                shape.HeightMm = EmuToMm(ecy);
        }
    }

    private static void ParseCustomGeometry(XElement custGeom, ShapeObject shape)
    {
        var path = custGeom.Descendants(XnsA + "path").FirstOrDefault();
        if (path is null) return;

        long pathW = 1, pathH = 1;
        if (long.TryParse(path.Attribute("w")?.Value, out var pw) && pw > 0) pathW = pw;
        if (long.TryParse(path.Attribute("h")?.Value, out var ph) && ph > 0) pathH = ph;

        // 라이터가 path w/h 를 cx/cy(EMU)로 설정했으므로 좌표도 EMU 단위.
        // shape.WidthMm 은 이 시점에서 아직 기본값일 수 있으므로 pathW 기반으로 mm 환산.
        double bboxWMm = EmuToMm(pathW);
        double bboxHMm = EmuToMm(pathH);

        bool hasClose = path.Descendants(XnsA + "close").Any();
        bool hasCurve = path.Descendants(XnsA + "cubicBezTo").Any()
                     || path.Descendants(XnsA + "quadBezTo").Any();
        shape.Kind = (hasClose, hasCurve) switch
        {
            (true,  true)  => ShapeKind.ClosedSpline,
            (false, true)  => ShapeKind.Spline,
            (true,  false) => ShapeKind.Polygon,
            (false, false) => ShapeKind.Polyline,
        };

        foreach (var cmd in path.Elements())
        {
            // moveTo/lnTo 는 단일 <a:pt>, cubicBezTo 는 3개(c1, c2, end), quadBezTo 는 2개(c, end).
            // 곡선의 endpoint 만 ShapePoint 로 채택해 점 시퀀스를 복원.
            string ln = cmd.Name.LocalName;
            XElement? endPt = ln switch
            {
                "moveTo" or "lnTo" => cmd.Element(XnsA + "pt"),
                "cubicBezTo"       => cmd.Elements(XnsA + "pt").ElementAtOrDefault(2),
                "quadBezTo"        => cmd.Elements(XnsA + "pt").ElementAtOrDefault(1),
                _                  => null,
            };
            if (endPt is null) continue;
            if (long.TryParse(endPt.Attribute("x")?.Value, out var px)
                && long.TryParse(endPt.Attribute("y")?.Value, out var py))
            {
                shape.Points.Add(new ShapePoint
                {
                    X = (double)px / pathW * bboxWMm,
                    Y = (double)py / pathH * bboxHMm,
                });
            }
        }

        // 닫힌 스플라인: writer 가 시작점으로 돌아오는 마지막 segment 까지 출력하므로
        // 마지막 endpoint 가 첫 점과 동일 → 중복 제거.
        if (shape.Kind == ShapeKind.ClosedSpline && shape.Points.Count >= 2)
        {
            var first = shape.Points[0];
            var last  = shape.Points[^1];
            if (Math.Abs(first.X - last.X) < 0.01 && Math.Abs(first.Y - last.Y) < 0.01)
                shape.Points.RemoveAt(shape.Points.Count - 1);
        }

        if (shape.Points.Count == 2 && shape.Kind == ShapeKind.Polyline)
            shape.Kind = ShapeKind.Line;
    }

    private static ShapeKind PrstToShapeKind(string prst) => prst switch
    {
        "rect"                => ShapeKind.Rectangle,
        "roundRect"           => ShapeKind.RoundedRect,
        "ellipse" or "circle" => ShapeKind.Ellipse,
        "line" or "straightConnector1" or "bentConnector2" or "bentConnector3" => ShapeKind.Line,
        "triangle" or "rtTriangle" => ShapeKind.Triangle,
        "pentagon"  => ShapeKind.RegularPolygon,
        "hexagon"   => ShapeKind.RegularPolygon,
        "heptagon"  => ShapeKind.RegularPolygon,
        "octagon"   => ShapeKind.RegularPolygon,
        "decagon"   => ShapeKind.RegularPolygon,
        "dodecagon" => ShapeKind.RegularPolygon,
        "star4" or "star5" or "star6" or "star7" or "star8" or "star10"
            or "star12" or "star16" or "star24" or "star32" => ShapeKind.Star,
        _ => ShapeKind.Rectangle,
    };

    private static int PrstToSideCount(string prst) => prst switch
    {
        "triangle" or "rtTriangle" => 3,
        "pentagon"  => 5,
        "hexagon"   => 6,
        "heptagon"  => 7,
        "octagon"   => 8,
        "decagon"   => 10,
        "dodecagon" => 12,
        "star4"  => 4,
        "star5"  => 5,
        "star6"  => 6,
        "star7"  => 7,
        "star8"  => 8,
        "star10" => 10,
        "star12" => 12,
        "star16" => 16,
        "star24" => 24,
        "star32" => 32,
        _ => 5,
    };

    private static ShapeArrow ArrowTypeFromAttr(string? value) => value switch
    {
        "arrow" or "stealth" => ShapeArrow.Open,
        "triangle"           => ShapeArrow.Filled,
        "diamond"            => ShapeArrow.Diamond,
        "oval"               => ShapeArrow.Circle,
        _                    => ShapeArrow.None,
    };

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

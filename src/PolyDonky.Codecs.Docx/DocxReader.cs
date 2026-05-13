using System.Globalization;
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
                case W.SectionProperties sectPr:
                    ReadSectionProperties(section.Page, sectPr, ctx);
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
        var drawingBlocks = new List<Block>();

        // 복합 필드(fldChar begin/separate/end) 상태 머신.
        bool inField        = false;   // begin ~ separate 사이에서 instrText 수집 중
        bool inFieldDisplay = false;   // separate ~ end 사이 (표시용 텍스트 → 건너뜀)
        var  fieldInstr     = new System.Text.StringBuilder();

        void FlushPendingField()
        {
            if (!inField && !inFieldDisplay) return;
            var ft = ParseFieldInstr(fieldInstr.ToString());
            if (ft is not null)
                paragraph.Runs.Add(new Run { Field = ft.Value });
            inField = inFieldDisplay = false;
            fieldInstr.Clear();
        }

        void ProcessRun(W.Run run, string? hyperlinkUrl = null)
        {
            // 각주·미주 참조 — 최우선 처리.
            if (run.Descendants<W.FootnoteReference>().FirstOrDefault() is { } fnRef
                && fnRef.Id?.Value is { } fnId)
            {
                FlushPendingField();
                paragraph.Runs.Add(new Run { FootnoteId = fnId.ToString(CultureInfo.InvariantCulture) });
                return;
            }
            if (run.Descendants<W.EndnoteReference>().FirstOrDefault() is { } enRef
                && enRef.Id?.Value is { } enId)
            {
                FlushPendingField();
                paragraph.Runs.Add(new Run { EndnoteId = enId.ToString(CultureInfo.InvariantCulture) });
                return;
            }

            // fldChar — 상태 전이.
            if (run.Descendants<W.FieldChar>().FirstOrDefault() is { } fc)
            {
                var fct = fc.FieldCharType?.Value;
                if (fct == W.FieldCharValues.Begin)
                {
                    inField = true;
                    inFieldDisplay = false;
                    fieldInstr.Clear();
                }
                else if (fct == W.FieldCharValues.Separate)
                {
                    inFieldDisplay = true;
                    // 필드 타입을 미리 판별해 두고, End 에서 최종 emit.
                }
                else if (fct == W.FieldCharValues.End)
                {
                    FlushPendingField();
                }
                return;
            }

            // instrText — 필드 명령어 수집.
            if (run.Descendants<W.FieldCode>().FirstOrDefault() is { } instrText)
            {
                fieldInstr.Append(instrText.Text);
                return;
            }

            // Separate ~ End 사이 표시용 텍스트는 건너뜀.
            if (inFieldDisplay) return;

            // 그림·도형.
            foreach (var drawing in run.Elements<W.Drawing>())
            {
                if (TryExtractImage(drawing, ctx, out var imageBlock))
                    drawingBlocks.Add(imageBlock);
                else if (TryExtractShape(drawing, out var shapeBlock))
                    drawingBlocks.Add(shapeBlock);
                else
                    target.Add(new OpaqueBlock
                    {
                        Format = "docx", Kind = "drawing",
                        Xml = drawing.OuterXml, DisplayLabel = "[보존된 도형]",
                    });
            }

            // 텍스트.
            var text = string.Concat(run.Elements<W.Text>().Select(t => t.Text));
            if (text.Length == 0) return;

            var style = ReadRunStyle(run.RunProperties);
            if (hyperlinkUrl is not null)
                paragraph.Runs.Add(new Run { Text = text, Style = style, Url = hyperlinkUrl });
            else
                paragraph.AddText(text, style);
        }

        foreach (var element in wp.ChildElements)
        {
            switch (element)
            {
                case W.Hyperlink hyperlink:
                {
                    var url = ResolveHyperlinkUrl(hyperlink, ctx.MainPart);
                    foreach (var hRun in hyperlink.Elements<W.Run>())
                        ProcessRun(hRun, url);
                    break;
                }
                case W.SimpleField fldSimple:
                {
                    var ft = ParseFieldInstr(fldSimple.Instruction?.Value);
                    if (ft is not null)
                        paragraph.Runs.Add(new Run { Field = ft.Value });
                    break;
                }
                case W.Run run:
                    ProcessRun(run);
                    break;
            }
        }
        FlushPendingField();

        if (paragraph.Runs.Count == 0)
            paragraph.AddText(string.Empty);
        target.Add(paragraph);
        foreach (var block in drawingBlocks)
            target.Add(block);
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

        // ── RowSpan 사전 계산: vMerge Restart 셀의 실제 병합 행 수를 미리 산출 ──
        // vMergeSpan[(rowIdx, startColIdx)] = 실제 RowSpan 값
        var rawRows = wtable.Elements<W.TableRow>().ToList();
        var vMergeSpan = new System.Collections.Generic.Dictionary<(int r, int c), int>();
        {
            // 1단계: (rowIdx, colIdx) → isContinue 맵 구성
            var isContinue = new System.Collections.Generic.HashSet<(int, int)>();
            var isRestart  = new System.Collections.Generic.HashSet<(int, int)>();
            for (int ri = 0; ri < rawRows.Count; ri++)
            {
                int colPos = 0;
                foreach (var c in rawRows[ri].Elements<W.TableCell>())
                {
                    var p = c.Elements<W.TableCellProperties>().FirstOrDefault();
                    int gs = (int)(p?.GridSpan?.Val?.Value ?? 1);
                    var vm = p?.VerticalMerge;
                    if (vm is not null)
                    {
                        if (vm.Val?.Value is null || vm.Val.Value.Equals(W.MergedCellValues.Continue))
                            isContinue.Add((ri, colPos));
                        else if (vm.Val.Value.Equals(W.MergedCellValues.Restart))
                            isRestart.Add((ri, colPos));
                    }
                    colPos += gs;
                }
            }
            // 2단계: Restart 셀마다 아래로 Continue 행 수를 센다
            foreach (var key in isRestart)
            {
                int span = 1;
                while (isContinue.Contains((key.Item1 + span, key.Item2))) span++;
                vMergeSpan[key] = span;
            }
        }

        int rowIdx = 0;
        foreach (var row in rawRows)
        {
            var tableRow = new TableRow();
            var trPr = row.Elements<W.TableRowProperties>().FirstOrDefault();
            var rowHeight = trPr?.Elements<W.TableRowHeight>().FirstOrDefault();
            if (rowHeight?.Val?.Value is uint h)
                tableRow.HeightMm = h / 56.6929;        // twips → mm
            // w:tblHeader → 머리글 행
            if (trPr?.Elements<W.TableHeader>().Any() == true)
                tableRow.IsHeader = true;

            int colIdx = 0;
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
                    colIdx += (int)span;
                    continue;
                }

                var tableCell = new TableCell
                {
                    ColumnSpan = (int)span,
                };
                if (vMerge?.Val?.Value is { } mergeVal && mergeVal.Equals(W.MergedCellValues.Restart))
                {
                    tableCell.RowSpan = vMergeSpan.TryGetValue((rowIdx, colIdx), out var rs) ? rs : 1;
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
                colIdx += (int)span;
            }

            table.Rows.Add(tableRow);
            rowIdx++;
        }

        return table;
    }

    // ── 표·셀 속성 (테두리·배경색) ────────────────────────────────────────────

    private static void ReadTableProperties(Table table, W.TableProperties? tblPr)
    {
        if (tblPr is null) return;

        // 표 외곽선 — 면별로 읽고, 공통값은 top 대표값으로 유지
        var borders = tblPr.TableBorders;
        if (borders is not null)
        {
            table.BorderTop    = ReadBorderSide(borders.TopBorder);
            table.BorderBottom = ReadBorderSide(borders.BottomBorder);
            table.BorderLeft   = ReadBorderSide(borders.LeftBorder);
            table.BorderRight  = ReadBorderSide(borders.RightBorder);
            table.InnerBorderHorizontal = ReadBorderSide(borders.InsideHorizontalBorder);
            table.InnerBorderVertical   = ReadBorderSide(borders.InsideVerticalBorder);

            // 공통값: top 면이 있으면 대표값, 없으면 다른 면 중 첫번째 값
            var rep = table.BorderTop ?? table.BorderBottom ?? table.BorderLeft ?? table.BorderRight;
            if (rep is { } r)
            {
                if (r.ThicknessPt > 0) table.BorderThicknessPt = r.ThicknessPt;
                if (r.Color is not null) table.BorderColor = r.Color;
            }
        }

        // 표 수평 정렬: tblPr > w:jc
        if (tblPr.TableJustification?.Val?.Value is { } jc)
        {
            if      (jc == W.TableRowAlignmentValues.Center) table.HAlign = TableHAlign.Center;
            else if (jc == W.TableRowAlignmentValues.Right)  table.HAlign = TableHAlign.Right;
        }

        // 표 바깥 여백: tblPr > w:tblInd (왼쪽 들여쓰기만 OOXML 에서 지원)
        if (tblPr.TableIndentation?.Width?.Value is { } indTwips)
            table.OuterMarginLeftMm = indTwips / 56.6929;

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

        // 셀 패딩: tcPr > w:tcMar (twips → mm)
        var mar = tcPr.TableCellMargin;
        if (mar is not null)
        {
            if (mar.TopMargin?.Width?.Value is { } mt && int.TryParse(mt, out var mtv) && mtv > 0)
                cell.PaddingTopMm    = mtv / 56.6929;
            if (mar.BottomMargin?.Width?.Value is { } mb && int.TryParse(mb, out var mbv) && mbv > 0)
                cell.PaddingBottomMm = mbv / 56.6929;
            if (mar.StartMargin?.Width?.Value is { } ml && int.TryParse(ml, out var mlv) && mlv > 0)
                cell.PaddingLeftMm   = mlv / 56.6929;
            else if (mar.LeftMargin?.Width?.Value is { } ll && int.TryParse(ll, out var llv) && llv > 0)
                cell.PaddingLeftMm   = llv / 56.6929;
            if (mar.EndMargin?.Width?.Value is { } mr && int.TryParse(mr, out var mrv) && mrv > 0)
                cell.PaddingRightMm  = mrv / 56.6929;
            else if (mar.RightMargin?.Width?.Value is { } rr && int.TryParse(rr, out var rrv) && rrv > 0)
                cell.PaddingRightMm  = rrv / 56.6929;
        }

        // 셀 세로 정렬: tcPr > w:vAlign
        if (tcPr.TableCellVerticalAlignment?.Val?.Value is { } va)
        {
            if      (va == W.TableVerticalAlignmentValues.Center) cell.VerticalAlign = CellVerticalAlign.Middle;
            else if (va == W.TableVerticalAlignmentValues.Bottom) cell.VerticalAlign = CellVerticalAlign.Bottom;
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
            string ln = cmd.Name.LocalName;

            if (ln == "cubicBezTo")
            {
                // 세 점: c1(나가는 제어점), c2(들어오는 제어점), end.
                // c1 은 이전 앵커(현재 shape.Points 마지막)의 OutCtrl,
                // c2 는 새 앵커(end)의 InCtrl 에 저장해 라운드트립 시 곡선 형태를 보존한다.
                var ptElems = cmd.Elements(XnsA + "pt").ToList();
                if (ptElems.Count < 3) continue;

                static double ToMm(XElement pt, long pathDim, double bboxMm)
                    => long.TryParse(pt.Attribute("x")?.Value, out var v)
                       ? (double)v / pathDim * bboxMm : 0;
                static double ToMmY(XElement pt, long pathDim, double bboxMm)
                    => long.TryParse(pt.Attribute("y")?.Value, out var v)
                       ? (double)v / pathDim * bboxMm : 0;

                double c1x = ToMm (ptElems[0], pathW, bboxWMm), c1y = ToMmY(ptElems[0], pathH, bboxHMm);
                double c2x = ToMm (ptElems[1], pathW, bboxWMm), c2y = ToMmY(ptElems[1], pathH, bboxHMm);
                double ex  = ToMm (ptElems[2], pathW, bboxWMm), ey  = ToMmY(ptElems[2], pathH, bboxHMm);

                // c1 → 직전 앵커의 OutCtrl
                if (shape.Points.Count > 0)
                {
                    var prev = shape.Points[^1];
                    prev.OutCtrlX = c1x;
                    prev.OutCtrlY = c1y;
                }

                // end 앵커 + InCtrl(c2)
                shape.Points.Add(new ShapePoint { X = ex, Y = ey, InCtrlX = c2x, InCtrlY = c2y });
                continue;
            }

            XElement? endPt = ln switch
            {
                "moveTo" or "lnTo" => cmd.Element(XnsA + "pt"),
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

    // ── 섹션 속성 (페이지 설정 + 머리말/꼬리말) ──────────────────────────────────

    private static void ReadSectionProperties(PageSettings page, W.SectionProperties sectPr, ReadContext ctx)
    {
        if (sectPr.GetFirstChild<W.PageSize>() is { } pgSz)
        {
            if (pgSz.Width?.Value  is uint w) page.WidthMm  = UnitConverter.TwipsToMm(w);
            if (pgSz.Height?.Value is uint h) page.HeightMm = UnitConverter.TwipsToMm(h);
            page.Orientation = pgSz.Orient?.Value == W.PageOrientationValues.Landscape
                ? PageOrientation.Landscape : PageOrientation.Portrait;
        }
        if (sectPr.GetFirstChild<W.PageMargin>() is { } pgMar)
        {
            if (pgMar.Top?.Value    is int  top)   page.MarginTopMm    = UnitConverter.TwipsToMm(top);
            if (pgMar.Bottom?.Value is int  bot)   page.MarginBottomMm = UnitConverter.TwipsToMm(bot);
            if (pgMar.Left?.Value   is uint left)  page.MarginLeftMm   = UnitConverter.TwipsToMm(left);
            if (pgMar.Right?.Value  is uint right) page.MarginRightMm  = UnitConverter.TwipsToMm(right);
            if (pgMar.Header?.Value is uint hdr)   page.MarginHeaderMm = UnitConverter.TwipsToMm(hdr);
            if (pgMar.Footer?.Value is uint ftr)   page.MarginFooterMm = UnitConverter.TwipsToMm(ftr);
        }
        ReadHeaderFooterFromSectPr(page, sectPr, ctx);
    }

    private static void ReadHeaderFooterFromSectPr(PageSettings page, W.SectionProperties sectPr, ReadContext ctx)
    {
        // 기본(default) 헤더/푸터만 읽는다. Even/Odd 는 후속 사이클에서 추가.
        foreach (var hRef in sectPr.Elements<W.HeaderReference>())
        {
            if (hRef.Type?.Value == W.HeaderFooterValues.Even) continue;
            if (hRef.Id?.Value is not { Length: > 0 } relId) continue;
            try
            {
                if (ctx.MainPart.GetPartById(relId) is not HeaderPart hp) continue;
                page.Header = ReadHeaderFooterContent(hp.Header?.Elements<W.Paragraph>(), ctx);
            }
            catch { /* 깨진 관계 무시 */ }
            break;
        }
        foreach (var fRef in sectPr.Elements<W.FooterReference>())
        {
            if (fRef.Type?.Value == W.HeaderFooterValues.Even) continue;
            if (fRef.Id?.Value is not { Length: > 0 } relId) continue;
            try
            {
                if (ctx.MainPart.GetPartById(relId) is not FooterPart fp) continue;
                page.Footer = ReadHeaderFooterContent(fp.Footer?.Elements<W.Paragraph>(), ctx);
            }
            catch { /* 깨진 관계 무시 */ }
            break;
        }
    }

    private static HeaderFooterContent ReadHeaderFooterContent(
        IEnumerable<W.Paragraph>? paragraphs, ReadContext ctx)
    {
        var content = new HeaderFooterContent();
        if (paragraphs is null) return content;

        foreach (var wp in paragraphs)
        {
            var slot = ParagraphAlignmentToSlot(wp.ParagraphProperties?.Justification?.Val?.Value, content);
            var blocks = new List<Block>();
            AppendParagraphAndExtractedDrawings(blocks, wp, ctx);
            foreach (var block in blocks)
            {
                if (block is Paragraph p)
                    slot.Paragraphs.Add(p);
            }
        }
        return content;
    }

    private static HeaderFooterSlot ParagraphAlignmentToSlot(
        W.JustificationValues? jc, HeaderFooterContent content)
    {
        if (jc == W.JustificationValues.Center) return content.Center;
        if (jc == W.JustificationValues.Right)  return content.Right;
        return content.Left;
    }

    // ── 하이퍼링크 ─────────────────────────────────────────────────────────────

    private static string? ResolveHyperlinkUrl(W.Hyperlink hyperlink, MainDocumentPart mainPart)
    {
        if (hyperlink.Id?.Value is { Length: > 0 } relId)
        {
            try
            {
                var rel = mainPart.HyperlinkRelationships.FirstOrDefault(r => r.Id == relId);
                if (rel is not null)
                    return rel.Uri.IsAbsoluteUri ? rel.Uri.AbsoluteUri : rel.Uri.ToString();
            }
            catch { /* 잘못된 URI 무시 */ }
        }
        if (hyperlink.Anchor?.Value is { Length: > 0 } anchor)
            return "#" + anchor;
        return null;
    }

    // ── 필드 코드 파싱 ─────────────────────────────────────────────────────────

    private static FieldType? ParseFieldInstr(string? instr)
    {
        if (string.IsNullOrWhiteSpace(instr)) return null;
        var upper = instr.Trim().ToUpperInvariant();
        if (MatchField(upper, "PAGE"))     return FieldType.Page;
        if (MatchField(upper, "NUMPAGES")) return FieldType.NumPages;
        if (MatchField(upper, "DATE"))     return FieldType.Date;
        if (MatchField(upper, "TIME"))     return FieldType.Time;
        if (MatchField(upper, "AUTHOR"))   return FieldType.Author;
        if (MatchField(upper, "TITLE"))    return FieldType.Title;
        return null;
    }

    private static bool MatchField(string upper, string name)
        => upper.StartsWith(name, StringComparison.Ordinal)
        && (upper.Length == name.Length || !char.IsLetterOrDigit(upper[name.Length]));

    // ── 코어 속성 ──────────────────────────────────────────────────────────────

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

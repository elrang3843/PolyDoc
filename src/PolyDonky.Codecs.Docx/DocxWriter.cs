using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using PolyDonky.Core;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using W = DocumentFormat.OpenXml.Wordprocessing;
using WP = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace PolyDonky.Codecs.Docx;

/// <summary>
/// PolyDonkyument → DOCX (OOXML WordprocessingML) 라이터.
///
/// 매핑 범위:
///   - 단락 / 인라인 런 / 제목(Heading1~6) / 정렬 / 기본 리스트
///   - 굵게·기울임·밑줄·취소선·위첨자·아래첨자·폰트·크기·색상
///   - 표 (Table → w:tbl, 셀 너비·병합·중첩 표 포함)
///   - 이미지 (ImageBlock → w:drawing + ImagePart 등록)
///   - OpaqueBlock — 원본 OuterXml 을 그대로 다시 출력 (보존 섬)
/// </summary>
public sealed class DocxWriter : IDocumentWriter
{
    public string FormatId => "docx";

    public void Write(PolyDonkyument document, Stream output)
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

        var ctx = new WriteContext(mainPart);

        foreach (var section in document.Sections)
        {
            foreach (var block in section.Blocks)
            {
                AppendBlock(body, block, ctx);
            }
        }

        // 섹션 속성 (페이지 설정).
        var firstSection = document.Sections.FirstOrDefault();
        body.AppendChild(BuildSectionProperties(firstSection?.Page ?? new PageSettings()));

        WriteFootnotesAndEndnotes(mainPart, document, ctx);
        WriteCoreProperties(package, document.Metadata);
    }

    private static void WriteFootnotesAndEndnotes(MainDocumentPart mainPart, PolyDonkyument document, WriteContext ctx)
    {
        if (document.Footnotes.Count > 0)
        {
            var fnPart = mainPart.AddNewPart<FootnotesPart>();
            var footnotes = new W.Footnotes();
            // OOXML 필수 separator 각주.
            footnotes.AppendChild(new W.Footnote { Type = W.FootnoteEndnoteValues.Separator, Id = -1 });
            footnotes.AppendChild(new W.Footnote { Type = W.FootnoteEndnoteValues.ContinuationSeparator, Id = 0 });
            foreach (var entry in document.Footnotes)
            {
                if (!int.TryParse(entry.Id, System.Globalization.NumberStyles.Integer,
                    System.Globalization.CultureInfo.InvariantCulture, out var fnId))
                    continue;
                var fn = new W.Footnote { Id = fnId };
                foreach (var block in entry.Blocks)
                    fn.AppendChild(block is Paragraph p ? BuildParagraph(p) : BuildParagraph(Paragraph.Of(string.Empty)));
                if (fn.ChildElements.Count == 0) fn.AppendChild(BuildParagraph(Paragraph.Of(string.Empty)));
                footnotes.AppendChild(fn);
            }
            fnPart.Footnotes = footnotes;
        }

        if (document.Endnotes.Count > 0)
        {
            var enPart = mainPart.AddNewPart<EndnotesPart>();
            var endnotes = new W.Endnotes();
            endnotes.AppendChild(new W.Endnote { Type = W.FootnoteEndnoteValues.Separator, Id = -1 });
            endnotes.AppendChild(new W.Endnote { Type = W.FootnoteEndnoteValues.ContinuationSeparator, Id = 0 });
            foreach (var entry in document.Endnotes)
            {
                if (!int.TryParse(entry.Id, System.Globalization.NumberStyles.Integer,
                    System.Globalization.CultureInfo.InvariantCulture, out var enId))
                    continue;
                var en = new W.Endnote { Id = enId };
                foreach (var block in entry.Blocks)
                    en.AppendChild(block is Paragraph p ? BuildParagraph(p) : BuildParagraph(Paragraph.Of(string.Empty)));
                if (en.ChildElements.Count == 0) en.AppendChild(BuildParagraph(Paragraph.Of(string.Empty)));
                endnotes.AppendChild(en);
            }
            enPart.Endnotes = endnotes;
        }
    }

    private sealed class WriteContext
    {
        private uint _nextDrawingId = 1;

        public WriteContext(MainDocumentPart mainPart)
        {
            MainPart = mainPart;
        }

        public MainDocumentPart MainPart { get; }

        public uint NextDrawingId() => _nextDrawingId++;
    }

    private static void AppendBlock(OpenXmlCompositeElement target, Block block, WriteContext ctx)
    {
        switch (block)
        {
            case Paragraph p:
                target.AppendChild(BuildParagraph(p));
                break;
            case Table t:
                target.AppendChild(BuildTable(t, ctx));
                break;
            case ImageBlock image:
                target.AppendChild(BuildImageParagraph(image, ctx));
                break;
            case ShapeObject shape:
                target.AppendChild(BuildShapeParagraph(shape, ctx));
                break;
            case OpaqueBlock opaque when !string.IsNullOrEmpty(opaque.Xml):
                AppendOpaqueXml(target, opaque);
                break;
            case OpaqueBlock opaque:
                target.AppendChild(BuildParagraph(Paragraph.Of(opaque.DisplayLabel)));
                break;
        }
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

        if (p.Style.ForcePageBreakBefore)
        {
            pPr.PageBreakBefore = new W.PageBreakBefore();
        }

        if (pPr.HasChildren)
        {
            wpara.AppendChild(pPr);
        }

        foreach (var run in p.Runs)
        {
            if (run.FootnoteId is { Length: > 0 } fnId
                && int.TryParse(fnId, System.Globalization.NumberStyles.Integer,
                    System.Globalization.CultureInfo.InvariantCulture, out var fnInt))
            {
                var fnRpr = new W.RunProperties();
                fnRpr.AppendChild(new W.VerticalTextAlignment { Val = W.VerticalPositionValues.Superscript });
                var fnRun = new W.Run();
                fnRun.AppendChild(fnRpr);
                fnRun.AppendChild(new W.FootnoteReference { Id = fnInt });
                wpara.AppendChild(fnRun);
                continue;
            }
            if (run.EndnoteId is { Length: > 0 } enId
                && int.TryParse(enId, System.Globalization.NumberStyles.Integer,
                    System.Globalization.CultureInfo.InvariantCulture, out var enInt))
            {
                var enRpr = new W.RunProperties();
                enRpr.AppendChild(new W.VerticalTextAlignment { Val = W.VerticalPositionValues.Superscript });
                var enRun = new W.Run();
                enRun.AppendChild(enRpr);
                enRun.AppendChild(new W.EndnoteReference { Id = enInt });
                wpara.AppendChild(enRun);
                continue;
            }
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

    private static W.Table BuildTable(Table table, WriteContext ctx)
    {
        var wtable = new W.Table();

        var tblPr = new W.TableProperties(BuildTableBorders(table));
        if (!string.IsNullOrEmpty(table.BackgroundColor))
        {
            tblPr.AppendChild(new W.Shading
            {
                Val   = W.ShadingPatternValues.Clear,
                Color = "auto",
                Fill  = table.BackgroundColor!.TrimStart('#').ToUpperInvariant(),
            });
        }
        wtable.AppendChild(tblPr);

        if (table.Columns.Count > 0)
        {
            var grid = new W.TableGrid();
            foreach (var col in table.Columns)
            {
                grid.AppendChild(new W.GridColumn
                {
                    Width = MmToTwipsString(col.WidthMm),
                });
            }
            wtable.AppendChild(grid);
        }

        foreach (var row in table.Rows)
        {
            var wrow = new W.TableRow();
            if (row.HeightMm > 0)
            {
                wrow.AppendChild(new W.TableRowProperties(
                    new W.TableRowHeight { Val = (uint)Math.Round(row.HeightMm * 56.6929) }));
            }
            foreach (var cell in row.Cells)
            {
                var wcell = new W.TableCell();
                var tcPr = new W.TableCellProperties();

                if (cell.ColumnSpan > 1)
                {
                    tcPr.AppendChild(new W.GridSpan { Val = cell.ColumnSpan });
                }
                if (cell.WidthMm > 0)
                {
                    tcPr.AppendChild(new W.TableCellWidth
                    {
                        Width = MmToTwipsString(cell.WidthMm),
                        Type = W.TableWidthUnitValues.Dxa,
                    });
                }

                // 셀 테두리 — 면별 지정이 하나라도 있으면 출력, 없으면 공통값 사용
                if (cell.BorderTop is not null || cell.BorderBottom is not null ||
                    cell.BorderLeft is not null || cell.BorderRight is not null ||
                    cell.BorderThicknessPt > 0 || !string.IsNullOrEmpty(cell.BorderColor))
                {
                    tcPr.AppendChild(BuildCellBorders(cell));
                }

                // 셀 배경색
                if (!string.IsNullOrEmpty(cell.BackgroundColor))
                {
                    tcPr.AppendChild(new W.Shading
                    {
                        Val   = W.ShadingPatternValues.Clear,
                        Color = "auto",
                        Fill  = cell.BackgroundColor!.TrimStart('#').ToUpperInvariant(),
                    });
                }

                if (tcPr.HasChildren)
                {
                    wcell.AppendChild(tcPr);
                }

                if (cell.Blocks.Count == 0)
                {
                    wcell.AppendChild(BuildParagraph(Paragraph.Of(string.Empty)));
                }
                else
                {
                    foreach (var inner in cell.Blocks)
                    {
                        AppendBlock(wcell, inner, ctx);
                    }
                    // DOCX 는 셀의 마지막 자식이 항상 paragraph 여야 한다.
                    if (wcell.LastChild is not W.Paragraph)
                    {
                        wcell.AppendChild(BuildParagraph(Paragraph.Of(string.Empty)));
                    }
                }

                wrow.AppendChild(wcell);
            }
            wtable.AppendChild(wrow);
        }
        return wtable;
    }

    private static string MmToTwipsString(double mm)
        => UnitConverter.MmToTwipsString(mm);

    // ── 표·셀 테두리 빌더 ─────────────────────────────────────────────────────

    private static W.TableBorders BuildTableBorders(Table table)
    {
        // Word 가 tblBorders 없으면 테두리를 표시하지 않는 경우가 있어 항상 출력한다.
        uint sz    = table.BorderThicknessPt > 0 ? (uint)Math.Round(table.BorderThicknessPt * 8) : 4U;
        string clr = !string.IsNullOrEmpty(table.BorderColor)
            ? table.BorderColor!.TrimStart('#').ToUpperInvariant() : "auto";

        return new W.TableBorders(
            new W.TopBorder              { Val = W.BorderValues.Single, Size = sz, Color = clr },
            new W.LeftBorder             { Val = W.BorderValues.Single, Size = sz, Color = clr },
            new W.BottomBorder           { Val = W.BorderValues.Single, Size = sz, Color = clr },
            new W.RightBorder            { Val = W.BorderValues.Single, Size = sz, Color = clr },
            new W.InsideHorizontalBorder { Val = W.BorderValues.Single, Size = sz, Color = clr },
            new W.InsideVerticalBorder   { Val = W.BorderValues.Single, Size = sz, Color = clr });
    }

    private static W.TableCellBorders BuildCellBorders(TableCell cell)
    {
        uint   defSz  = cell.BorderThicknessPt > 0 ? (uint)Math.Round(cell.BorderThicknessPt * 8) : 4U;
        string defClr = !string.IsNullOrEmpty(cell.BorderColor)
            ? cell.BorderColor!.TrimStart('#').ToUpperInvariant() : "auto";

        return new W.TableCellBorders(
            BuildOneBorder<W.TopBorder>   (cell.BorderTop,    defSz, defClr),
            BuildOneBorder<W.LeftBorder>  (cell.BorderLeft,   defSz, defClr),
            BuildOneBorder<W.BottomBorder>(cell.BorderBottom, defSz, defClr),
            BuildOneBorder<W.RightBorder> (cell.BorderRight,  defSz, defClr));
    }

    private static T BuildOneBorder<T>(CellBorderSide? side, uint defSz, string defClr)
        where T : W.BorderType, new()
    {
        uint   sz  = side.HasValue && side.Value.ThicknessPt > 0
            ? (uint)Math.Round(side.Value.ThicknessPt * 8) : defSz;
        string clr = side.HasValue && !string.IsNullOrEmpty(side.Value.Color)
            ? side.Value.Color!.TrimStart('#').ToUpperInvariant() : defClr;
        return new T { Val = W.BorderValues.Single, Size = sz, Color = clr };
    }

    private static W.Paragraph BuildImageParagraph(ImageBlock image, WriteContext ctx)
    {
        var paragraph = new W.Paragraph();
        var run = new W.Run();
        run.AppendChild(BuildImageDrawing(image, ctx));
        paragraph.AppendChild(run);
        return paragraph;
    }

    private static W.Drawing BuildImageDrawing(ImageBlock image, WriteContext ctx)
    {
        var imagePart = ctx.MainPart.AddImagePart(MediaTypeToPartType(image.MediaType));
        using (var stream = imagePart.GetStream())
        {
            stream.Write(image.Data, 0, image.Data.Length);
        }
        var relId = ctx.MainPart.GetIdOfPart(imagePart);

        long cx = MmToEmu(image.WidthMm > 0 ? image.WidthMm : 80);     // 기본 80mm
        long cy = MmToEmu(image.HeightMm > 0 ? image.HeightMm : 60);

        var drawingId = ctx.NextDrawingId();
        var name = $"Picture {drawingId}";

        return new W.Drawing(
            new WP.Inline(
                new WP.Extent { Cx = cx, Cy = cy },
                new WP.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                new WP.DocProperties
                {
                    Id = drawingId,
                    Name = name,
                    Description = image.Description ?? string.Empty,
                },
                new WP.NonVisualGraphicFrameDrawingProperties(
                    new A.GraphicFrameLocks { NoChangeAspect = true }),
                new A.Graphic(
                    new A.GraphicData(
                        new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                                new PIC.NonVisualDrawingProperties { Id = 0U, Name = name },
                                new PIC.NonVisualPictureDrawingProperties()),
                            new PIC.BlipFill(
                                new A.Blip { Embed = relId },
                                new A.Stretch(new A.FillRectangle())),
                            new PIC.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset { X = 0L, Y = 0L },
                                    new A.Extents { Cx = cx, Cy = cy }),
                                new A.PresetGeometry(new A.AdjustValueList())
                                {
                                    Preset = A.ShapeTypeValues.Rectangle,
                                })
                        )
                    )
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
            )
            {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U,
            });
    }

    private static long MmToEmu(double mm) => UnitConverter.MmToEmu(mm);

    // ── 도형 출력 ─────────────────────────────────────────────────────────────

    private static W.Paragraph BuildShapeParagraph(ShapeObject shape, WriteContext ctx)
    {
        var paragraph = new W.Paragraph();
        var run = new W.Run();
        run.AppendChild(BuildShapeDrawing(shape, ctx));
        paragraph.AppendChild(run);
        return paragraph;
    }

    private static W.Drawing BuildShapeDrawing(ShapeObject shape, WriteContext ctx)
    {
        long cx = MmToEmu(shape.WidthMm  > 0 ? shape.WidthMm  : 40);
        long cy = MmToEmu(shape.HeightMm > 0 ? shape.HeightMm : 30);
        var drawingId = ctx.NextDrawingId();
        var name = $"Shape {drawingId}";

        // a:graphicData の中身を XML 文字列で注入 (wps: はエクステンション名前空間)
        var graphicData = new A.GraphicData
        {
            Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
        };
        graphicData.InnerXml = BuildWspXml(shape, cx, cy);

        var graphic = new A.Graphic(graphicData);

        if (shape.WrapMode is ImageWrapMode.InFrontOfText or ImageWrapMode.BehindText)
        {
            long posX = MmToEmu(shape.OverlayXMm);
            long posY = MmToEmu(shape.OverlayYMm);
            bool behind = shape.WrapMode == ImageWrapMode.BehindText;

            var posH = new WP.HorizontalPosition { RelativeFrom = WP.HorizontalRelativePositionValues.Page };
            posH.AppendChild(new WP.PositionOffset(posX.ToString(CultureInfo.InvariantCulture)));

            var posV = new WP.VerticalPosition { RelativeFrom = WP.VerticalRelativePositionValues.Page };
            posV.AppendChild(new WP.PositionOffset(posY.ToString(CultureInfo.InvariantCulture)));

            var anchor = new WP.Anchor
            {
                DistanceFromTop    = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft   = 0U,
                DistanceFromRight  = 0U,
                SimplePos          = false,
                RelativeHeight     = 0U,
                BehindDoc          = behind,
                Locked             = false,
                LayoutInCell       = true,
                AllowOverlap       = true,
            };
            anchor.AppendChild(new WP.SimplePosition { X = 0L, Y = 0L });
            anchor.AppendChild(posH);
            anchor.AppendChild(posV);
            anchor.AppendChild(new WP.Extent { Cx = cx, Cy = cy });
            anchor.AppendChild(new WP.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L });
            anchor.AppendChild(new WP.WrapNone());
            anchor.AppendChild(new WP.DocProperties { Id = drawingId, Name = name });
            anchor.AppendChild(new WP.NonVisualGraphicFrameDrawingProperties());
            anchor.AppendChild(graphic);

            return new W.Drawing(anchor);
        }
        else
        {
            // Inline (WrapMode.Inline 및 WrapLeft/WrapRight 1차 처리)
            return new W.Drawing(
                new WP.Inline(
                    new WP.Extent { Cx = cx, Cy = cy },
                    new WP.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                    new WP.DocProperties { Id = drawingId, Name = name },
                    new WP.NonVisualGraphicFrameDrawingProperties(),
                    graphic
                )
                {
                    DistanceFromTop    = 0U,
                    DistanceFromBottom = 0U,
                    DistanceFromLeft   = 0U,
                    DistanceFromRight  = 0U,
                });
        }
    }

    /// <summary>wps:wsp 요소의 OuterXml 을 반환한다.</summary>
    private static string BuildWspXml(ShapeObject shape, long cx, long cy)
    {
        long rotEmu  = (long)Math.Round(shape.RotationAngleDeg * 60000);
        var geomXml  = BuildGeometryXml(shape, cx, cy);
        var fillXml  = BuildFillXml(shape);
        var lineXml  = BuildOutlineXml(shape);
        var txbxXml  = BuildLabelTxbxXml(shape);

        return
            "<wps:wsp" +
            " xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\"" +
            " xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"" +
            " xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
            "<wps:spPr>" +
            $"<a:xfrm rot=\"{rotEmu}\"><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm>" +
            geomXml +
            fillXml +
            lineXml +
            "</wps:spPr>" +
            txbxXml +
            "<wps:bodyPr/>" +
            "</wps:wsp>";
    }

    private static string BuildGeometryXml(ShapeObject shape, long cx, long cy)
    {
        var preset = GetPresetGeometry(shape);
        if (preset is not null)
            return $"<a:prstGeom prst=\"{preset}\"><a:avLst/></a:prstGeom>";

        return BuildCustomGeometryXml(shape, cx, cy);
    }

    private static string? GetPresetGeometry(ShapeObject shape) => shape.Kind switch
    {
        ShapeKind.Rectangle   => "rect",
        ShapeKind.RoundedRect => "roundRect",
        ShapeKind.Ellipse     => "ellipse",
        ShapeKind.Triangle    => "triangle",
        ShapeKind.Line when shape.Points.Count < 2 => "line",
        ShapeKind.RegularPolygon => shape.SideCount switch
        {
            3  => "triangle",
            4  => "rect",
            5  => "pentagon",
            6  => "hexagon",
            7  => "heptagon",
            8  => "octagon",
            10 => "decagon",
            12 => "dodecagon",
            _  => null,
        },
        ShapeKind.Star => shape.SideCount switch
        {
            4  => "star4",
            5  => "star5",
            6  => "star6",
            7  => "star7",
            8  => "star8",
            10 => "star10",
            12 => "star12",
            16 => "star16",
            24 => "star24",
            32 => "star32",
            _  => null,
        },
        _ => null,
    };

    private static string BuildCustomGeometryXml(ShapeObject shape, long cx, long cy)
    {
        var pts = shape.Points;
        if (pts.Count < 2)
            return "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>";

        double wMm = shape.WidthMm  > 0 ? shape.WidthMm  : 40;
        double hMm = shape.HeightMm > 0 ? shape.HeightMm : 30;
        bool closed   = shape.Kind is ShapeKind.Polygon or ShapeKind.ClosedSpline;
        bool isSpline = shape.Kind is ShapeKind.Spline  or ShapeKind.ClosedSpline;

        var sb = new System.Text.StringBuilder();
        sb.Append("<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>");
        sb.Append("<a:rect l=\"0\" t=\"0\" r=\"r\" b=\"b\"/>");
        sb.Append($"<a:pathLst><a:path w=\"{cx}\" h=\"{cy}\">");

        // 모든 점을 EMU 좌표로 미리 환산.
        int n = pts.Count;
        var xs = new long[n];
        var ys = new long[n];
        for (int i = 0; i < n; i++)
        {
            xs[i] = (long)Math.Round(pts[i].X / wMm * cx);
            ys[i] = (long)Math.Round(pts[i].Y / hMm * cy);
        }

        sb.Append($"<a:moveTo><a:pt x=\"{xs[0]}\" y=\"{ys[0]}\"/></a:moveTo>");

        if (isSpline)
        {
            // Catmull-Rom → cubic Bezier (tension=0.5) 변환.
            // 닫힌 스플라인은 시작점으로 돌아오는 마지막 segment 도 그려 매끄러운 폐곡선.
            int segments = closed ? n : n - 1;
            for (int i = 0; i < segments; i++)
            {
                int i1 = i;
                int i2 = (i + 1) % n;
                int i0 = closed ? (i - 1 + n) % n : Math.Max(0, i - 1);
                int i3 = closed ? (i + 2) % n     : Math.Min(n - 1, i + 2);
                long c1x = xs[i1] + (xs[i2] - xs[i0]) / 6;
                long c1y = ys[i1] + (ys[i2] - ys[i0]) / 6;
                long c2x = xs[i2] - (xs[i3] - xs[i1]) / 6;
                long c2y = ys[i2] - (ys[i3] - ys[i1]) / 6;
                sb.Append("<a:cubicBezTo>");
                sb.Append($"<a:pt x=\"{c1x}\" y=\"{c1y}\"/>");
                sb.Append($"<a:pt x=\"{c2x}\" y=\"{c2y}\"/>");
                sb.Append($"<a:pt x=\"{xs[i2]}\" y=\"{ys[i2]}\"/>");
                sb.Append("</a:cubicBezTo>");
            }
        }
        else
        {
            for (int i = 1; i < n; i++)
            {
                sb.Append($"<a:lnTo><a:pt x=\"{xs[i]}\" y=\"{ys[i]}\"/></a:lnTo>");
            }
        }

        if (closed) sb.Append("<a:close/>");
        sb.Append("</a:path></a:pathLst></a:custGeom>");
        return sb.ToString();
    }

    private static string BuildFillXml(ShapeObject shape)
    {
        if (string.IsNullOrEmpty(shape.FillColor))
            return "<a:noFill/>";

        var hex = shape.FillColor!.TrimStart('#').ToUpperInvariant();
        if (Math.Abs(shape.FillOpacity - 1.0) < 0.001)
            return $"<a:solidFill><a:srgbClr val=\"{hex}\"/></a:solidFill>";

        var alphaVal = (int)Math.Round(shape.FillOpacity * 100000);
        return $"<a:solidFill><a:srgbClr val=\"{hex}\"><a:alpha val=\"{alphaVal}\"/></a:srgbClr></a:solidFill>";
    }

    private static string BuildOutlineXml(ShapeObject shape)
    {
        if (shape.StrokeThicknessPt <= 0)
            return "<a:ln><a:noFill/></a:ln>";

        long strokeEmu = (long)Math.Round(shape.StrokeThicknessPt * 12700);
        var hex = (shape.StrokeColor ?? "#000000").TrimStart('#').ToUpperInvariant();

        var dashXml = shape.StrokeDash switch
        {
            StrokeDash.Dashed  => "<a:prstDash val=\"dash\"/>",
            StrokeDash.Dotted  => "<a:prstDash val=\"dot\"/>",
            StrokeDash.DashDot => "<a:prstDash val=\"dashDot\"/>",
            _                  => string.Empty,
        };

        var headXml = shape.StartArrow != ShapeArrow.None
            ? $"<a:headEnd type=\"{ArrowToType(shape.StartArrow)}\"/>" : string.Empty;
        var tailXml = shape.EndArrow != ShapeArrow.None
            ? $"<a:tailEnd type=\"{ArrowToType(shape.EndArrow)}\"/>" : string.Empty;

        return $"<a:ln w=\"{strokeEmu}\">" +
               $"<a:solidFill><a:srgbClr val=\"{hex}\"/></a:solidFill>" +
               dashXml + headXml + tailXml +
               "</a:ln>";
    }

    private static string BuildLabelTxbxXml(ShapeObject shape)
    {
        if (string.IsNullOrEmpty(shape.LabelText)) return string.Empty;
        var escaped = shape.LabelText!
            .Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");
        return
            "<wps:txbx>" +
            "<w:txbxContent><w:p><w:r><w:t>" +
            escaped +
            "</w:t></w:r></w:p></w:txbxContent>" +
            "</wps:txbx>";
    }

    private static string ArrowToType(ShapeArrow arrow) => arrow switch
    {
        ShapeArrow.Open    => "arrow",
        ShapeArrow.Filled  => "triangle",
        ShapeArrow.Diamond => "diamond",
        ShapeArrow.Circle  => "oval",
        _                  => "none",
    };

    /// <summary>
    /// OpaqueBlock 의 OuterXml 을 임시 Body 컨테이너로 파싱한 뒤 그 자식 요소들을 target 에 옮긴다.
    /// (OpenXmlUnknownElement 의 OuterXml 직접 주입은 namespace 컨텍스트 부재로 직렬화 단계에서 실패.)
    /// </summary>
    private static void AppendOpaqueXml(OpenXmlCompositeElement target, OpaqueBlock opaque)
    {
        try
        {
            var temp = new W.Body();
            temp.InnerXml = opaque.Xml!;
            foreach (var child in temp.ChildElements.ToList())
            {
                child.Remove();
                target.AppendChild(child);
            }
        }
        catch
        {
            // 잘못된 XML 은 빈 단락으로 격하 (사용자 데이터 손실 방지).
            target.AppendChild(BuildParagraph(Paragraph.Of(opaque.DisplayLabel)));
        }
    }

    private static PartTypeInfo MediaTypeToPartType(string mediaType) => mediaType switch
    {
        "image/png" => ImagePartType.Png,
        "image/jpeg" => ImagePartType.Jpeg,
        "image/gif" => ImagePartType.Gif,
        "image/bmp" => ImagePartType.Bmp,
        "image/tiff" => ImagePartType.Tiff,
        "image/x-emf" => ImagePartType.Emf,
        "image/x-wmf" => ImagePartType.Wmf,
        _ => ImagePartType.Png,
    };

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
        var props = new W.SectionProperties();
        props.AppendChild(new W.PageSize
        {
            Width = UnitConverter.MmToTwipsUInt(page.WidthMm),
            Height = UnitConverter.MmToTwipsUInt(page.HeightMm),
            Orient = page.Orientation == PageOrientation.Landscape
                ? W.PageOrientationValues.Landscape
                : W.PageOrientationValues.Portrait,
        });
        props.AppendChild(new W.PageMargin
        {
            Top = UnitConverter.MmToTwipsInt(page.MarginTopMm),
            Right = UnitConverter.MmToTwipsUInt(page.MarginRightMm),
            Bottom = UnitConverter.MmToTwipsInt(page.MarginBottomMm),
            Left = UnitConverter.MmToTwipsUInt(page.MarginLeftMm),
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

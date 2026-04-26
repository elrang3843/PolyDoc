using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using PolyDoc.Core;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using W = DocumentFormat.OpenXml.Wordprocessing;
using WP = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace PolyDoc.Codecs.Docx;

/// <summary>
/// PolyDocument → DOCX (OOXML WordprocessingML) 라이터.
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

        WriteCoreProperties(package, document.Metadata);
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

    private static W.Table BuildTable(Table table, WriteContext ctx)
    {
        var wtable = new W.Table();

        // 표 기본 속성 — 테두리 정도. (Word 가 보더 없이 보여주는 일이 있어 기본 single 적용.)
        var tblPr = new W.TableProperties(
            new W.TableBorders(
                new W.TopBorder { Val = W.BorderValues.Single, Size = 4U },
                new W.LeftBorder { Val = W.BorderValues.Single, Size = 4U },
                new W.BottomBorder { Val = W.BorderValues.Single, Size = 4U },
                new W.RightBorder { Val = W.BorderValues.Single, Size = 4U },
                new W.InsideHorizontalBorder { Val = W.BorderValues.Single, Size = 4U },
                new W.InsideVerticalBorder { Val = W.BorderValues.Single, Size = 4U }));
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
        => ((int)Math.Round(mm * 56.6929)).ToString(CultureInfo.InvariantCulture);

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

    private static long MmToEmu(double mm) => (long)Math.Round(mm / 25.4 * 914400.0);

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

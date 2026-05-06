using PolyDonky.Codecs.Docx;
using PolyDonky.Core;

namespace PolyDonky.Codecs.Docx.Tests;

public class DocxRoundTripTests
{
    [Fact]
    public void RoundTrip_PreservesParagraphsAndHeadings()
    {
        var doc = new PolyDonkyument();
        doc.Metadata.Title = "DOCX 라운드트립";
        doc.Metadata.Author = "Noh JinMoon";
        var section = new Section();
        doc.Sections.Add(section);

        var heading1 = new Paragraph { Style = { Outline = OutlineLevel.H1 } };
        heading1.AddText("제목 1");
        section.Blocks.Add(heading1);

        var heading2 = new Paragraph { Style = { Outline = OutlineLevel.H2 } };
        heading2.AddText("부제목");
        section.Blocks.Add(heading2);

        var body = new Paragraph();
        body.AddText("본문 ");
        body.AddText("강조", new RunStyle { Bold = true });
        body.AddText(" 와 ");
        body.AddText("기울임", new RunStyle { Italic = true });
        body.AddText(".");
        section.Blocks.Add(body);

        var roundTripped = WriteThenRead(doc);

        Assert.Equal("DOCX 라운드트립", roundTripped.Metadata.Title);
        Assert.Equal("Noh JinMoon", roundTripped.Metadata.Author);

        var paragraphs = roundTripped.EnumerateParagraphs().ToList();
        Assert.Equal(3, paragraphs.Count);
        Assert.Equal(OutlineLevel.H1, paragraphs[0].Style.Outline);
        Assert.Equal("제목 1", paragraphs[0].GetPlainText());
        Assert.Equal(OutlineLevel.H2, paragraphs[1].Style.Outline);
        Assert.Equal("부제목", paragraphs[1].GetPlainText());
        Assert.Equal(OutlineLevel.Body, paragraphs[2].Style.Outline);
        Assert.Contains(paragraphs[2].Runs, r => r.Style.Bold && r.Text == "강조");
        Assert.Contains(paragraphs[2].Runs, r => r.Style.Italic && r.Text == "기울임");
    }

    [Fact]
    public void RoundTrip_PreservesAlignment()
    {
        var doc = new PolyDonkyument();
        var section = new Section();
        doc.Sections.Add(section);

        foreach (var alignment in new[] { Alignment.Left, Alignment.Center, Alignment.Right, Alignment.Justify })
        {
            var p = new Paragraph { Style = { Alignment = alignment } };
            p.AddText($"align={alignment}");
            section.Blocks.Add(p);
        }

        var roundTripped = WriteThenRead(doc);
        var paragraphs = roundTripped.EnumerateParagraphs().ToList();

        Assert.Equal(Alignment.Left, paragraphs[0].Style.Alignment);
        Assert.Equal(Alignment.Center, paragraphs[1].Style.Alignment);
        Assert.Equal(Alignment.Right, paragraphs[2].Style.Alignment);
        Assert.Equal(Alignment.Justify, paragraphs[3].Style.Alignment);
    }

    [Fact]
    public void RoundTrip_PreservesUnderlineStrikeAndScripts()
    {
        var doc = new PolyDonkyument();
        var section = new Section();
        doc.Sections.Add(section);
        var p = new Paragraph();
        p.AddText("u", new RunStyle { Underline = true });
        p.AddText("s", new RunStyle { Strikethrough = true });
        p.AddText("super", new RunStyle { Superscript = true });
        p.AddText("sub", new RunStyle { Subscript = true });
        section.Blocks.Add(p);

        var roundTripped = WriteThenRead(doc);
        var runs = roundTripped.EnumerateParagraphs().Single().Runs;

        Assert.True(runs.Single(r => r.Text == "u").Style.Underline);
        Assert.True(runs.Single(r => r.Text == "s").Style.Strikethrough);
        Assert.True(runs.Single(r => r.Text == "super").Style.Superscript);
        Assert.True(runs.Single(r => r.Text == "sub").Style.Subscript);
    }

    [Fact]
    public void RoundTrip_PreservesFontFamilyAndSize()
    {
        var doc = new PolyDonkyument();
        var section = new Section();
        doc.Sections.Add(section);
        var p = new Paragraph();
        p.AddText("monospace", new RunStyle { FontFamily = "Consolas", FontSizePt = 14 });
        section.Blocks.Add(p);

        var roundTripped = WriteThenRead(doc);
        var run = roundTripped.EnumerateParagraphs().Single().Runs.Single();

        Assert.Equal("Consolas", run.Style.FontFamily);
        Assert.Equal(14, run.Style.FontSizePt, precision: 1);
    }

    [Fact]
    public void RoundTrip_PreservesForegroundColor()
    {
        var doc = new PolyDonkyument();
        var section = new Section();
        doc.Sections.Add(section);
        var p = new Paragraph();
        p.AddText("red", new RunStyle { Foreground = Color.FromHex("#FF3300") });
        section.Blocks.Add(p);

        var roundTripped = WriteThenRead(doc);
        var run = roundTripped.EnumerateParagraphs().Single().Runs.Single();

        Assert.NotNull(run.Style.Foreground);
        Assert.Equal(0xFF, run.Style.Foreground!.Value.R);
        Assert.Equal(0x33, run.Style.Foreground!.Value.G);
        Assert.Equal(0x00, run.Style.Foreground!.Value.B);
    }

    [Fact]
    public void Read_ThrowsOnNonDocxStream()
    {
        var bytes = System.Text.Encoding.UTF8.GetBytes("not a docx");
        using var ms = new MemoryStream(bytes);
        Assert.ThrowsAny<Exception>(() => new DocxReader().Read(ms));
    }

    [Fact]
    public void RoundTrip_PreservesTableStructure()
    {
        var table = new Table();
        table.Columns.Add(new TableColumn { WidthMm = 40 });
        table.Columns.Add(new TableColumn { WidthMm = 60 });

        var headerRow = new TableRow();
        headerRow.Cells.Add(new TableCell { Blocks = { Paragraph.Of("이름") } });
        headerRow.Cells.Add(new TableCell { Blocks = { Paragraph.Of("값") } });
        table.Rows.Add(headerRow);

        var dataRow = new TableRow();
        dataRow.Cells.Add(new TableCell { Blocks = { Paragraph.Of("월") } });
        dataRow.Cells.Add(new TableCell { Blocks = { Paragraph.Of("31일") } });
        table.Rows.Add(dataRow);

        var doc = new PolyDonkyument();
        var section = new Section();
        section.Blocks.Add(table);
        doc.Sections.Add(section);

        var roundTripped = WriteThenRead(doc);
        var t = roundTripped.Sections[0].Blocks.OfType<Table>().Single();

        Assert.Equal(2, t.Rows.Count);
        Assert.Equal(2, t.Rows[0].Cells.Count);
        Assert.Equal("이름", ((Paragraph)t.Rows[0].Cells[0].Blocks[0]).GetPlainText());
        Assert.Equal("값", ((Paragraph)t.Rows[0].Cells[1].Blocks[0]).GetPlainText());
        Assert.Equal("월", ((Paragraph)t.Rows[1].Cells[0].Blocks[0]).GetPlainText());
        Assert.Equal("31일", ((Paragraph)t.Rows[1].Cells[1].Blocks[0]).GetPlainText());
    }

    [Fact]
    public void RoundTrip_PreservesImageBytes()
    {
        // 1×1 투명 PNG 의 최소 바이트열.
        byte[] tinyPng = new byte[]
        {
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
            0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
            0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
            0x89, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x44, 0x41,
            0x54, 0x78, 0x9C, 0x63, 0x00, 0x01, 0x00, 0x00,
            0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00,
            0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
            0x42, 0x60, 0x82,
        };

        var doc = new PolyDonkyument();
        var section = new Section();
        doc.Sections.Add(section);
        section.Blocks.Add(new ImageBlock
        {
            MediaType = "image/png",
            Data = tinyPng,
            WidthMm = 50,
            HeightMm = 30,
            Description = "라운드트립 이미지",
        });

        var roundTripped = WriteThenRead(doc);
        var image = roundTripped.Sections[0].Blocks.OfType<ImageBlock>().Single();

        Assert.Equal(tinyPng, image.Data);
        Assert.Equal("image/png", image.MediaType);
        Assert.Equal(50, image.WidthMm, precision: 0);
        Assert.Equal(30, image.HeightMm, precision: 0);
        Assert.Equal("라운드트립 이미지", image.Description);
    }

    [Fact]
    public void RoundTrip_OpaqueBlock_IsPreservedThroughDocx()
    {
        // OpaqueBlock 의 OuterXml 은 Word 의 sdt(content control) 같은 미인식 요소를 보존.
        // OpenXmlUnknownElement 는 leading whitespace 를 element name 의 일부로 해석하므로 single-line 으로 정리.
        const string opaqueXml =
            "<w:sdt xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
            "<w:sdtContent><w:p><w:r><w:t>opaque payload</w:t></w:r></w:p></w:sdtContent>" +
            "</w:sdt>";

        var doc = new PolyDonkyument();
        var section = new Section();
        doc.Sections.Add(section);
        section.Blocks.Add(new OpaqueBlock
        {
            Format = "docx",
            Kind = "sdt",
            Xml = opaqueXml,
            DisplayLabel = "[보존된 sdt]",
        });

        // 라운드트립 후 OpaqueBlock 으로 다시 보존되어야 한다 (DocxReader 가 미인식 블록을 OpaqueBlock 으로 흡수).
        var roundTripped = WriteThenRead(doc);

        var preserved = roundTripped.Sections[0].Blocks.OfType<OpaqueBlock>().FirstOrDefault();
        Assert.NotNull(preserved);
        Assert.Equal("docx", preserved!.Format);
        Assert.Equal("sdt", preserved.Kind);
        Assert.Contains("opaque payload", preserved.Xml ?? string.Empty);
    }

    // ── 표 테두리·배경색 라운드트립 테스트 ──────────────────────────────────────

    [Fact]
    public void RoundTrip_PreservesTableBorderThicknessAndColor()
    {
        var table = new Table
        {
            BorderThicknessPt = 2.0,
            BorderColor       = "#FF0000",
        };
        table.Columns.Add(new TableColumn { WidthMm = 50 });
        var row = new TableRow();
        row.Cells.Add(new TableCell { Blocks = { Paragraph.Of("셀") } });
        table.Rows.Add(row);

        var doc = new PolyDonkyument();
        var section = new Section();
        section.Blocks.Add(table);
        doc.Sections.Add(section);

        var rt = WriteThenRead(doc);
        var t = rt.Sections[0].Blocks.OfType<Table>().Single();

        Assert.Equal(2.0, t.BorderThicknessPt, precision: 1);
        Assert.Equal("#FF0000", t.BorderColor, StringComparer.OrdinalIgnoreCase);
    }

    [Fact]
    public void RoundTrip_PreservesTableBackgroundColor()
    {
        var table = new Table
        {
            BackgroundColor = "#FFEECC",
        };
        table.Columns.Add(new TableColumn { WidthMm = 50 });
        var row = new TableRow();
        row.Cells.Add(new TableCell { Blocks = { Paragraph.Of("셀") } });
        table.Rows.Add(row);

        var doc = new PolyDonkyument();
        var section = new Section();
        section.Blocks.Add(table);
        doc.Sections.Add(section);

        var rt = WriteThenRead(doc);
        var t = rt.Sections[0].Blocks.OfType<Table>().Single();

        Assert.Equal("#FFEECC", t.BackgroundColor, StringComparer.OrdinalIgnoreCase);
    }

    [Fact]
    public void RoundTrip_PreservesCellBorderAndBackground()
    {
        var table = new Table();
        table.Columns.Add(new TableColumn { WidthMm = 40 });
        table.Columns.Add(new TableColumn { WidthMm = 40 });

        var row = new TableRow();
        row.Cells.Add(new TableCell
        {
            Blocks            = { Paragraph.Of("헤더") },
            BackgroundColor   = "#336699",
            BorderThicknessPt = 1.5,
            BorderColor       = "#000080",
        });
        row.Cells.Add(new TableCell
        {
            Blocks          = { Paragraph.Of("일반") },
            BackgroundColor = "#FFFFFF",
        });
        table.Rows.Add(row);

        var doc = new PolyDonkyument();
        var section = new Section();
        section.Blocks.Add(table);
        doc.Sections.Add(section);

        var rt = WriteThenRead(doc);
        var cells = rt.Sections[0].Blocks.OfType<Table>().Single().Rows[0].Cells;

        Assert.Equal("#336699", cells[0].BackgroundColor, StringComparer.OrdinalIgnoreCase);
        Assert.Equal(1.5,       cells[0].BorderThicknessPt, precision: 1);
        Assert.Equal("#000080", cells[0].BorderColor, StringComparer.OrdinalIgnoreCase);
        Assert.Equal("#FFFFFF", cells[1].BackgroundColor, StringComparer.OrdinalIgnoreCase);
    }

    [Fact]
    public void RoundTrip_PreservesPerSideCellBorders()
    {
        var table = new Table();
        table.Columns.Add(new TableColumn { WidthMm = 60 });

        var row = new TableRow();
        row.Cells.Add(new TableCell
        {
            Blocks    = { Paragraph.Of("셀") },
            BorderTop    = new CellBorderSide(3.0, "#FF0000"),
            BorderBottom = new CellBorderSide(1.0, "#00FF00"),
            BorderLeft   = new CellBorderSide(2.0, "#0000FF"),
            BorderRight  = new CellBorderSide(0.5, "#FFFF00"),
        });
        table.Rows.Add(row);

        var doc = new PolyDonkyument();
        var section = new Section();
        section.Blocks.Add(table);
        doc.Sections.Add(section);

        var rt = WriteThenRead(doc);
        var cell = rt.Sections[0].Blocks.OfType<Table>().Single().Rows[0].Cells[0];

        Assert.NotNull(cell.BorderTop);
        Assert.Equal(3.0,     cell.BorderTop!.Value.ThicknessPt, precision: 1);
        Assert.Equal("#FF0000", cell.BorderTop.Value.Color,  StringComparer.OrdinalIgnoreCase);

        Assert.NotNull(cell.BorderBottom);
        Assert.Equal(1.0,     cell.BorderBottom!.Value.ThicknessPt, precision: 1);
        Assert.Equal("#00FF00", cell.BorderBottom.Value.Color, StringComparer.OrdinalIgnoreCase);

        Assert.NotNull(cell.BorderLeft);
        Assert.Equal(2.0,     cell.BorderLeft!.Value.ThicknessPt, precision: 1);
        Assert.Equal("#0000FF", cell.BorderLeft.Value.Color,  StringComparer.OrdinalIgnoreCase);

        Assert.NotNull(cell.BorderRight);
        Assert.Equal(0.5,     cell.BorderRight!.Value.ThicknessPt, precision: 1);
        Assert.Equal("#FFFF00", cell.BorderRight.Value.Color, StringComparer.OrdinalIgnoreCase);
    }

    // ── 도형 라운드트립 테스트 ────────────────────────────────────────────────

    [Fact]
    public void RoundTrip_PreservesRectangleShape()
    {
        var doc = new PolyDonkyument();
        var section = new Section();
        doc.Sections.Add(section);

        section.Blocks.Add(new ShapeObject
        {
            Kind               = ShapeKind.Rectangle,
            WrapMode           = ImageWrapMode.InFrontOfText,
            WidthMm            = 60,
            HeightMm           = 40,
            OverlayXMm         = 20,
            OverlayYMm         = 30,
            FillColor          = "#FF0000",
            StrokeColor        = "#0000FF",
            StrokeThicknessPt  = 2.0,
        });

        var rt = WriteThenRead(doc);
        var shape = rt.Sections[0].Blocks.OfType<ShapeObject>().Single();

        Assert.Equal(ShapeKind.Rectangle,          shape.Kind);
        Assert.Equal(ImageWrapMode.InFrontOfText,  shape.WrapMode);
        Assert.Equal(60,  shape.WidthMm,  precision: 0);
        Assert.Equal(40,  shape.HeightMm, precision: 0);
        Assert.Equal(20,  shape.OverlayXMm, precision: 0);
        Assert.Equal(30,  shape.OverlayYMm, precision: 0);
        Assert.Equal("#FF0000", shape.FillColor,   StringComparer.OrdinalIgnoreCase);
        Assert.Equal("#0000FF", shape.StrokeColor, StringComparer.OrdinalIgnoreCase);
        Assert.Equal(2.0, shape.StrokeThicknessPt, precision: 1);
    }

    [Fact]
    public void RoundTrip_PreservesEllipseShape()
    {
        var doc = new PolyDonkyument();
        var section = new Section();
        doc.Sections.Add(section);

        section.Blocks.Add(new ShapeObject
        {
            Kind      = ShapeKind.Ellipse,
            WrapMode  = ImageWrapMode.Inline,
            WidthMm   = 50,
            HeightMm  = 30,
            FillColor = "#00FF00",
        });

        var rt = WriteThenRead(doc);
        var shape = rt.Sections[0].Blocks.OfType<ShapeObject>().Single();

        Assert.Equal(ShapeKind.Ellipse,       shape.Kind);
        Assert.Equal(ImageWrapMode.Inline,    shape.WrapMode);
        Assert.Equal(50, shape.WidthMm,  precision: 0);
        Assert.Equal(30, shape.HeightMm, precision: 0);
        Assert.Equal("#00FF00", shape.FillColor, StringComparer.OrdinalIgnoreCase);
    }

    [Fact]
    public void RoundTrip_PreservesPolylineShape()
    {
        var doc = new PolyDonkyument();
        var section = new Section();
        doc.Sections.Add(section);

        section.Blocks.Add(new ShapeObject
        {
            Kind     = ShapeKind.Polyline,
            WrapMode = ImageWrapMode.InFrontOfText,
            WidthMm  = 80,
            HeightMm = 40,
            Points   = { new ShapePoint { X = 0, Y = 40 }, new ShapePoint { X = 40, Y = 0 }, new ShapePoint { X = 80, Y = 40 } },
            FillColor = null,
            StrokeColor = "#000000",
            StrokeThicknessPt = 1.5,
        });

        var rt = WriteThenRead(doc);
        var shape = rt.Sections[0].Blocks.OfType<ShapeObject>().Single();

        Assert.Equal(ShapeKind.Polyline, shape.Kind);
        Assert.Equal(3, shape.Points.Count);
        Assert.Equal(80, shape.WidthMm,  precision: 0);
        Assert.Equal(40, shape.HeightMm, precision: 0);
        Assert.Equal(0,  shape.Points[0].X, precision: 0);
        Assert.Equal(40, shape.Points[0].Y, precision: 0);
        Assert.Equal(40, shape.Points[1].X, precision: 0);
        Assert.Equal(0,  shape.Points[1].Y, precision: 0);
        Assert.Equal(80, shape.Points[2].X, precision: 0);
        Assert.Equal(40, shape.Points[2].Y, precision: 0);
    }

    [Fact]
    public void RoundTrip_PreservesSplineShape()
    {
        var doc = new PolyDonkyument();
        var section = new Section();
        doc.Sections.Add(section);

        section.Blocks.Add(new ShapeObject
        {
            Kind     = ShapeKind.Spline,
            WrapMode = ImageWrapMode.InFrontOfText,
            WidthMm  = 80,
            HeightMm = 40,
            Points   = { new ShapePoint { X = 0,  Y = 40 },
                         new ShapePoint { X = 20, Y = 0  },
                         new ShapePoint { X = 60, Y = 40 },
                         new ShapePoint { X = 80, Y = 0  } },
            FillColor         = null,
            StrokeColor       = "#0000FF",
            StrokeThicknessPt = 1.5,
        });

        var rt = WriteThenRead(doc);
        var shape = rt.Sections[0].Blocks.OfType<ShapeObject>().Single();

        Assert.Equal(ShapeKind.Spline, shape.Kind);
        Assert.Equal(4, shape.Points.Count);
        Assert.Equal(0,  shape.Points[0].X, precision: 0);
        Assert.Equal(40, shape.Points[0].Y, precision: 0);
        Assert.Equal(20, shape.Points[1].X, precision: 0);
        Assert.Equal(0,  shape.Points[1].Y, precision: 0);
        Assert.Equal(60, shape.Points[2].X, precision: 0);
        Assert.Equal(40, shape.Points[2].Y, precision: 0);
        Assert.Equal(80, shape.Points[3].X, precision: 0);
        Assert.Equal(0,  shape.Points[3].Y, precision: 0);
    }

    [Fact]
    public void RoundTrip_PreservesClosedSplineShape()
    {
        var doc = new PolyDonkyument();
        var section = new Section();
        doc.Sections.Add(section);

        section.Blocks.Add(new ShapeObject
        {
            Kind     = ShapeKind.ClosedSpline,
            WrapMode = ImageWrapMode.InFrontOfText,
            WidthMm  = 60,
            HeightMm = 60,
            Points   = { new ShapePoint { X = 30, Y = 0  },
                         new ShapePoint { X = 60, Y = 30 },
                         new ShapePoint { X = 30, Y = 60 },
                         new ShapePoint { X = 0,  Y = 30 } },
            FillColor         = "#FFCCCC",
            StrokeColor       = "#FF0000",
            StrokeThicknessPt = 1.0,
        });

        var rt = WriteThenRead(doc);
        var shape = rt.Sections[0].Blocks.OfType<ShapeObject>().Single();

        Assert.Equal(ShapeKind.ClosedSpline, shape.Kind);
        Assert.Equal(4, shape.Points.Count);
        Assert.Equal(30, shape.Points[0].X, precision: 0);
        Assert.Equal(0,  shape.Points[0].Y, precision: 0);
        Assert.Equal(60, shape.Points[1].X, precision: 0);
        Assert.Equal(30, shape.Points[1].Y, precision: 0);
        Assert.Equal(30, shape.Points[2].X, precision: 0);
        Assert.Equal(60, shape.Points[2].Y, precision: 0);
        Assert.Equal(0,  shape.Points[3].X, precision: 0);
        Assert.Equal(30, shape.Points[3].Y, precision: 0);
        Assert.Equal("#FFCCCC", shape.FillColor, StringComparer.OrdinalIgnoreCase);
    }

    [Fact]
    public void RoundTrip_PreservesShapeRotation()
    {
        var doc = new PolyDonkyument();
        var section = new Section();
        doc.Sections.Add(section);

        section.Blocks.Add(new ShapeObject
        {
            Kind             = ShapeKind.Rectangle,
            WrapMode         = ImageWrapMode.InFrontOfText,
            WidthMm          = 50,
            HeightMm         = 30,
            RotationAngleDeg = 45.0,
        });

        var rt = WriteThenRead(doc);
        var shape = rt.Sections[0].Blocks.OfType<ShapeObject>().Single();

        Assert.Equal(45.0, shape.RotationAngleDeg, precision: 1);
    }

    [Fact]
    public void RoundTrip_PreservesShapeStrokeDash()
    {
        var doc = new PolyDonkyument();
        var section = new Section();
        doc.Sections.Add(section);

        section.Blocks.Add(new ShapeObject
        {
            Kind              = ShapeKind.Rectangle,
            WrapMode          = ImageWrapMode.Inline,
            WidthMm           = 40,
            HeightMm          = 20,
            StrokeDash        = StrokeDash.Dashed,
            StrokeThicknessPt = 1.0,
        });

        var rt = WriteThenRead(doc);
        var shape = rt.Sections[0].Blocks.OfType<ShapeObject>().Single();

        Assert.Equal(StrokeDash.Dashed, shape.StrokeDash);
    }

    [Fact]
    public void RoundTrip_PreservesShapeLabel()
    {
        var doc = new PolyDonkyument();
        var section = new Section();
        doc.Sections.Add(section);

        section.Blocks.Add(new ShapeObject
        {
            Kind      = ShapeKind.Ellipse,
            WrapMode  = ImageWrapMode.InFrontOfText,
            WidthMm   = 60,
            HeightMm  = 40,
            LabelText = "도형 레이블",
        });

        var rt = WriteThenRead(doc);
        var shape = rt.Sections[0].Blocks.OfType<ShapeObject>().Single();

        Assert.Equal("도형 레이블", shape.LabelText);
    }

    [Fact]
    public void RoundTrip_PreservesShapeBehindText()
    {
        var doc = new PolyDonkyument();
        var section = new Section();
        doc.Sections.Add(section);

        section.Blocks.Add(new ShapeObject
        {
            Kind       = ShapeKind.Rectangle,
            WrapMode   = ImageWrapMode.BehindText,
            WidthMm    = 50,
            HeightMm   = 30,
            FillColor  = "#CCCCCC",
            OverlayXMm = 10,
            OverlayYMm = 10,
        });

        var rt = WriteThenRead(doc);
        var shape = rt.Sections[0].Blocks.OfType<ShapeObject>().Single();

        Assert.Equal(ImageWrapMode.BehindText, shape.WrapMode);
        Assert.Equal(50, shape.WidthMm,  precision: 0);
        Assert.Equal(30, shape.HeightMm, precision: 0);
    }

    [Fact]
    public void RoundTrip_PreservesForcePageBreakBefore()
    {
        var doc = new PolyDonkyument();
        var section = new Section();
        doc.Sections.Add(section);

        section.Blocks.Add(Paragraph.Of("첫 번째 단락"));
        var p2 = Paragraph.Of("두 번째 단락 — 강제 페이지 나누기");
        p2.Style.ForcePageBreakBefore = true;
        section.Blocks.Add(p2);
        section.Blocks.Add(Paragraph.Of("세 번째 단락"));

        var rt = WriteThenRead(doc);
        var paragraphs = rt.Sections[0].Blocks.OfType<Paragraph>().ToList();

        Assert.Equal(3, paragraphs.Count);
        Assert.False(paragraphs[0].Style.ForcePageBreakBefore);
        Assert.True(paragraphs[1].Style.ForcePageBreakBefore);
        Assert.False(paragraphs[2].Style.ForcePageBreakBefore);
    }

    [Fact]
    public void RoundTrip_PreservesFootnotesAndEndnotes()
    {
        var doc = new PolyDonkyument();
        var section = new Section();
        doc.Sections.Add(section);

        var fn = new FootnoteEntry { Id = "1" };
        fn.Blocks.Add(Paragraph.Of("각주 내용 A"));
        doc.Footnotes.Add(fn);

        var en = new FootnoteEntry { Id = "2" };
        en.Blocks.Add(Paragraph.Of("미주 내용 B"));
        doc.Endnotes.Add(en);

        var p = new Paragraph();
        p.AddText("본문");
        p.Runs.Add(new Run { FootnoteId = "1" });
        p.AddText(" 텍스트");
        p.Runs.Add(new Run { EndnoteId = "2" });
        section.Blocks.Add(p);

        var rt = WriteThenRead(doc);

        Assert.Single(rt.Footnotes);
        Assert.Equal("각주 내용 A", rt.Footnotes[0].Blocks.OfType<Paragraph>().First().GetPlainText());

        Assert.Single(rt.Endnotes);
        Assert.Equal("미주 내용 B", rt.Endnotes[0].Blocks.OfType<Paragraph>().First().GetPlainText());

        var runs = rt.Sections[0].Blocks.OfType<Paragraph>().First().Runs;
        Assert.Contains(runs, r => r.FootnoteId is not null);
        Assert.Contains(runs, r => r.EndnoteId is not null);
    }

    private static PolyDonkyument WriteThenRead(PolyDonkyument document)
    {
        using var ms = new MemoryStream();
        new DocxWriter().Write(document, ms);
        ms.Position = 0;
        return new DocxReader().Read(ms);
    }
}

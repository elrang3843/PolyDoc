using PolyDoc.Codecs.Docx;
using PolyDoc.Core;

namespace PolyDoc.Codecs.Docx.Tests;

public class DocxRoundTripTests
{
    [Fact]
    public void RoundTrip_PreservesParagraphsAndHeadings()
    {
        var doc = new PolyDocument();
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
        var doc = new PolyDocument();
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
        var doc = new PolyDocument();
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
        var doc = new PolyDocument();
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
        var doc = new PolyDocument();
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

        var doc = new PolyDocument();
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

        var doc = new PolyDocument();
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

        var doc = new PolyDocument();
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

    private static PolyDocument WriteThenRead(PolyDocument document)
    {
        using var ms = new MemoryStream();
        new DocxWriter().Write(document, ms);
        ms.Position = 0;
        return new DocxReader().Read(ms);
    }
}

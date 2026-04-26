using System.IO.Compression;
using PolyDoc.Codecs.Hwpx;
using PolyDoc.Core;

namespace PolyDoc.Codecs.Hwpx.Tests;

public class HwpxRoundTripTests
{
    [Fact]
    public void RoundTrip_PreservesParagraphsAndHeadings()
    {
        var doc = new PolyDocument();
        doc.Metadata.Title = "HWPX 라운드트립";
        doc.Metadata.Author = "Noh JinMoon";

        var section = new Section();
        doc.Sections.Add(section);

        var heading = new Paragraph { Style = { Outline = OutlineLevel.H1 } };
        heading.AddText("제목");
        section.Blocks.Add(heading);

        var body = new Paragraph();
        body.AddText("본문 ");
        body.AddText("강조", new RunStyle { Bold = true });
        body.AddText(" 와 ");
        body.AddText("기울임", new RunStyle { Italic = true });
        body.AddText(" 끝.");
        section.Blocks.Add(body);

        var roundTripped = WriteThenRead(doc);

        Assert.Equal("HWPX 라운드트립", roundTripped.Metadata.Title);
        Assert.Equal("Noh JinMoon", roundTripped.Metadata.Author);

        var paragraphs = roundTripped.EnumerateParagraphs().ToList();
        Assert.Equal(2, paragraphs.Count);
        Assert.Equal(OutlineLevel.H1, paragraphs[0].Style.Outline);
        Assert.Equal("제목", paragraphs[0].GetPlainText());
        Assert.Equal(OutlineLevel.Body, paragraphs[1].Style.Outline);
        Assert.Contains(paragraphs[1].Runs, r => r.Style.Bold && r.Text == "강조");
        Assert.Contains(paragraphs[1].Runs, r => r.Style.Italic && r.Text == "기울임");
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
    public void RoundTrip_PreservesUnderlineAndStrikethrough()
    {
        var doc = new PolyDocument();
        var section = new Section();
        doc.Sections.Add(section);

        var p = new Paragraph();
        p.AddText("u", new RunStyle { Underline = true });
        p.AddText("s", new RunStyle { Strikethrough = true });
        section.Blocks.Add(p);

        var run = WriteThenRead(doc).EnumerateParagraphs().Single().Runs;
        Assert.True(run.Single(r => r.Text == "u").Style.Underline);
        Assert.True(run.Single(r => r.Text == "s").Style.Strikethrough);
    }

    [Fact]
    public void Package_HasMimetypeAsFirstStoredEntry()
    {
        var doc = PolyDocument.Empty();
        doc.Sections[0].Blocks.Add(Paragraph.Of("content"));

        var bytes = WriteToBytes(doc);
        using var ms = new MemoryStream(bytes);
        using var archive = new ZipArchive(ms, ZipArchiveMode.Read);

        // HWPX 사양: 첫 entry 는 mimetype 이며 STORED (압축 없음).
        var first = archive.Entries[0];
        Assert.Equal("mimetype", first.FullName);
        Assert.Equal(first.CompressedLength, first.Length);
    }

    [Fact]
    public void Package_ContainsRequiredParts()
    {
        var doc = PolyDocument.Empty();
        var bytes = WriteToBytes(doc);
        using var ms = new MemoryStream(bytes);
        using var archive = new ZipArchive(ms, ZipArchiveMode.Read);

        Assert.NotNull(archive.GetEntry("mimetype"));
        Assert.NotNull(archive.GetEntry("META-INF/container.xml"));
        Assert.NotNull(archive.GetEntry("Contents/content.hpf"));
        Assert.NotNull(archive.GetEntry("Contents/header.xml"));
        Assert.NotNull(archive.GetEntry("Contents/section0.xml"));
        Assert.NotNull(archive.GetEntry("version.xml"));
    }

    [Fact]
    public void Read_ThrowsOnWrongMimetype()
    {
        // 임의 ZIP 에 잘못된 mimetype 을 넣어 우리 reader 가 거부하는지 확인.
        using var ms = new MemoryStream();
        using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            var entry = archive.CreateEntry("mimetype", CompressionLevel.NoCompression);
            using var stream = entry.Open();
            var bytes = System.Text.Encoding.UTF8.GetBytes("application/epub+zip");
            stream.Write(bytes, 0, bytes.Length);
        }
        ms.Position = 0;
        Assert.Throws<InvalidDataException>(() => new HwpxReader().Read(ms));
    }

    [Fact]
    public void RoundTrip_PreservesTableStructure()
    {
        var table = new Table();
        table.Columns.Add(new TableColumn { WidthMm = 30 });
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
        });

        var roundTripped = WriteThenRead(doc);
        var image = roundTripped.Sections[0].Blocks.OfType<ImageBlock>().Single();

        Assert.Equal(tinyPng, image.Data);
        Assert.Equal("image/png", image.MediaType);
        Assert.Equal(50, image.WidthMm, precision: 0);
        Assert.Equal(30, image.HeightMm, precision: 0);
    }

    [Fact]
    public void RoundTrip_DedupesIdenticalImagesIntoSingleBinDataPart()
    {
        var bytes = new byte[] { 1, 2, 3, 4, 5 };
        var doc = new PolyDocument();
        var section = new Section();
        doc.Sections.Add(section);
        section.Blocks.Add(new ImageBlock { MediaType = "image/png", Data = bytes });
        section.Blocks.Add(new ImageBlock { MediaType = "image/png", Data = (byte[])bytes.Clone() });

        var packageBytes = WriteToBytes(doc);

        using var ms = new MemoryStream(packageBytes);
        using var archive = new ZipArchive(ms, ZipArchiveMode.Read);
        var binDataEntries = archive.Entries.Count(e => e.FullName.StartsWith("BinData/", StringComparison.Ordinal));
        Assert.Equal(1, binDataEntries);
    }

    [Fact]
    public void RoundTrip_PreservesFontSize()
    {
        var doc = new PolyDocument();
        var section = new Section();
        doc.Sections.Add(section);

        var p = new Paragraph();
        p.AddText("작은",  new RunStyle { FontSizePt = 9 });
        p.AddText("보통",  new RunStyle { FontSizePt = 11 });
        p.AddText("큰",    new RunStyle { FontSizePt = 18 });
        p.AddText("매우큰", new RunStyle { FontSizePt = 36 });
        section.Blocks.Add(p);

        var runs = WriteThenRead(doc).EnumerateParagraphs().Single().Runs;
        Assert.Equal(9,  runs.Single(r => r.Text == "작은").Style.FontSizePt,  precision: 1);
        Assert.Equal(11, runs.Single(r => r.Text == "보통").Style.FontSizePt,  precision: 1);
        Assert.Equal(18, runs.Single(r => r.Text == "큰").Style.FontSizePt,    precision: 1);
        Assert.Equal(36, runs.Single(r => r.Text == "매우큰").Style.FontSizePt, precision: 1);
    }

    [Fact]
    public void RoundTrip_PreservesForegroundColor()
    {
        var doc = new PolyDocument();
        var section = new Section();
        doc.Sections.Add(section);

        var p = new Paragraph();
        p.AddText("빨강", new RunStyle { Foreground = new Color(255, 0,   0)   });
        p.AddText("초록", new RunStyle { Foreground = new Color(0,   128, 0)   });
        p.AddText("파랑", new RunStyle { Foreground = new Color(0,   0,   255) });
        section.Blocks.Add(p);

        var runs = WriteThenRead(doc).EnumerateParagraphs().Single().Runs;
        Assert.Equal(new Color(255, 0,   0),   runs.Single(r => r.Text == "빨강").Style.Foreground);
        Assert.Equal(new Color(0,   128, 0),   runs.Single(r => r.Text == "초록").Style.Foreground);
        Assert.Equal(new Color(0,   0,   255), runs.Single(r => r.Text == "파랑").Style.Foreground);
    }

    [Fact]
    public void RoundTrip_PreservesFontFamily()
    {
        var doc = new PolyDocument();
        var section = new Section();
        doc.Sections.Add(section);

        var p = new Paragraph();
        p.AddText("고딕",   new RunStyle { FontFamily = "맑은 고딕" });
        p.AddText("명조",   new RunStyle { FontFamily = "바탕" });
        p.AddText("고정폭", new RunStyle { FontFamily = "Consolas" });
        section.Blocks.Add(p);

        var runs = WriteThenRead(doc).EnumerateParagraphs().Single().Runs;
        Assert.Equal("맑은 고딕", runs.Single(r => r.Text == "고딕").Style.FontFamily);
        Assert.Equal("바탕",      runs.Single(r => r.Text == "명조").Style.FontFamily);
        Assert.Equal("Consolas",  runs.Single(r => r.Text == "고정폭").Style.FontFamily);
    }

    [Fact]
    public void RoundTrip_PreservesMultipleStylesCombined()
    {
        var doc = new PolyDocument();
        var section = new Section();
        doc.Sections.Add(section);

        var p = new Paragraph();
        p.AddText("복합", new RunStyle
        {
            FontFamily = "바탕",
            FontSizePt = 14,
            Bold = true,
            Foreground = new Color(200, 50, 50),
        });
        section.Blocks.Add(p);

        var run = WriteThenRead(doc).EnumerateParagraphs().Single().Runs.Single();
        Assert.Equal("바탕",            run.Style.FontFamily);
        Assert.Equal(14,                run.Style.FontSizePt, precision: 1);
        Assert.True(run.Style.Bold);
        Assert.Equal(new Color(200, 50, 50), run.Style.Foreground);
    }

    private static byte[] WriteToBytes(PolyDocument doc)
    {
        using var ms = new MemoryStream();
        new HwpxWriter().Write(doc, ms);
        return ms.ToArray();
    }

    private static PolyDocument WriteThenRead(PolyDocument doc)
    {
        var bytes = WriteToBytes(doc);
        using var ms = new MemoryStream(bytes);
        return new HwpxReader().Read(ms);
    }
}

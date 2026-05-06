using System.IO.Compression;
using PolyDonky.Codecs.Hwpx;
using PolyDonky.Core;

namespace PolyDonky.Codecs.Hwpx.Tests;

public class HwpxRoundTripTests
{
    [Fact]
    public void RoundTrip_PreservesParagraphsAndHeadings()
    {
        var doc = new PolyDonkyument();
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
    public void RoundTrip_PreservesUnderlineAndStrikethrough()
    {
        var doc = new PolyDonkyument();
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
        var doc = PolyDonkyument.Empty();
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
        var doc = PolyDonkyument.Empty();
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
        var doc = new PolyDonkyument();
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
        var doc = new PolyDonkyument();
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
        var doc = new PolyDonkyument();
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
        var doc = new PolyDonkyument();
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
        var doc = new PolyDonkyument();
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

    [Fact]
    public void RoundTrip_PreservesHwpxOpaqueShape()
    {
        // hp:rect 모양의 최소 XML — 실제 한컴 hwpx 도형 구조 시뮬레이션.
        const string rectXml = "<hp:rect xmlns:hp=\"http://www.hancom.co.kr/hwpml/2011/paragraph\" " +
                                "width=\"5000\" height=\"3000\" />";

        var doc = new PolyDonkyument();
        var section = new Section();
        doc.Sections.Add(section);
        section.Blocks.Add(Paragraph.Of("앞"));
        section.Blocks.Add(new OpaqueBlock
        {
            Format = "hwpx",
            Kind = "rect",
            Xml = rectXml,
            DisplayLabel = "[사각형]",
        });
        section.Blocks.Add(Paragraph.Of("뒤"));

        var roundTripped = WriteThenRead(doc);
        var blocks = roundTripped.Sections[0].Blocks;

        // OpaqueBlock 이 보존되어 중간에 살아남아야 한다.
        var opaque = blocks.OfType<OpaqueBlock>().SingleOrDefault();
        Assert.NotNull(opaque);
        Assert.Equal("hwpx", opaque.Format);
        Assert.Equal("rect", opaque.Kind);
        Assert.NotNull(opaque.Xml);
        Assert.Contains("rect", opaque.Xml, StringComparison.Ordinal);
    }

    [Fact]
    public void RoundTrip_NonHwpxOpaqueBecomesPlaceholderParagraph()
    {
        var doc = new PolyDonkyument();
        var section = new Section();
        doc.Sections.Add(section);
        section.Blocks.Add(new OpaqueBlock
        {
            Format = "docx",
            Kind = "sdt",
            Xml = "<w:sdt/>",
            DisplayLabel = "[SDT]",
        });

        // docx 포맷 OpaqueBlock 은 HWPX 에 placeholder 단락으로 변환된다 — 예외 없이 완료돼야 한다.
        var roundTripped = WriteThenRead(doc);
        Assert.NotEmpty(roundTripped.EnumerateParagraphs());
    }

    [Fact]
    public void RoundTrip_PreservesTableBorderColorAndThickness()
    {
        var t = new Table {
            BorderColor = "#FF0000",
            BorderThicknessPt = 2.0,
        };
        t.Columns.Add(new TableColumn { WidthMm = 30 });
        t.Rows.Add(new TableRow { Cells = { new TableCell { Blocks = { Paragraph.Of("A") } } } });
        var doc = new PolyDonkyument();
        var section = new Section();
        section.Blocks.Add(t);
        doc.Sections.Add(section);

        var rt = WriteThenRead(doc);
        var rtTable = rt.Sections[0].Blocks.OfType<Table>().Single();

        Assert.Equal("#FF0000", rtTable.BorderColor);
        // 2pt → 표준 0.7mm 로 snap 후 reader 가 mm → pt 환산 — 2pt 근처값.
        Assert.InRange(rtTable.BorderThicknessPt, 1.9, 2.1);
    }

    [Fact]
    public void RoundTrip_PreservesTableBackgroundColor()
    {
        var t = new Table {
            BorderColor = "#000000",
            BorderThicknessPt = 0.5,
            BackgroundColor = "#EEEEEE",
        };
        t.Columns.Add(new TableColumn { WidthMm = 30 });
        t.Rows.Add(new TableRow { Cells = { new TableCell { Blocks = { Paragraph.Of("A") } } } });
        var doc = new PolyDonkyument();
        var section = new Section();
        section.Blocks.Add(t);
        doc.Sections.Add(section);

        var rt = WriteThenRead(doc);
        var rtTable = rt.Sections[0].Blocks.OfType<Table>().Single();

        Assert.Equal("#EEEEEE", rtTable.BackgroundColor);
    }

    [Fact]
    public void RoundTrip_PreservesTableCellBorderColorAndThickness()
    {
        // 2x2 표 — 좌상 셀의 inner side(bottom/right) 가 cell BorderColor 를 보존.
        // 1x1 표는 모든 면이 외곽 = 표 BorderColor 라서 cell 색이 인코딩에 안 들어감.
        // 2pt → 0.7mm 정확 매핑 → reader 가 다시 ~1.98pt 회수 (snap 오차 < 1%).
        // 다른 값은 표준 너비에 매핑되며 round-trip 오차가 클 수 있음.
        TableCell MakeCell(string color) => new()
        {
            Blocks = { Paragraph.Of("X") },
            BorderColor = color,
            BorderThicknessPt = 2.0,
        };
        var t = new Table { BorderColor = "#000000", BorderThicknessPt = 0.5 };
        t.Columns.Add(new TableColumn { WidthMm = 30 });
        t.Columns.Add(new TableColumn { WidthMm = 30 });
        t.Rows.Add(new TableRow { Cells = { MakeCell("#0000FF"), MakeCell("#0000FF") } });
        t.Rows.Add(new TableRow { Cells = { MakeCell("#0000FF"), MakeCell("#0000FF") } });
        var doc = new PolyDonkyument();
        var section = new Section();
        section.Blocks.Add(t);
        doc.Sections.Add(section);

        var rt = WriteThenRead(doc);
        var rtCell = rt.Sections[0].Blocks.OfType<Table>().Single().Rows[0].Cells[0];

        Assert.Equal("#0000FF", rtCell.BorderColor);
        Assert.InRange(rtCell.BorderThicknessPt, 1.9, 2.1);
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

        var fn = new FootnoteEntry { Id = "fn01" };
        fn.Blocks.Add(Paragraph.Of("각주 내용 A"));
        doc.Footnotes.Add(fn);

        var en = new FootnoteEntry { Id = "en01" };
        en.Blocks.Add(Paragraph.Of("미주 내용 B"));
        doc.Endnotes.Add(en);

        var p = new Paragraph();
        p.AddText("본문");
        p.Runs.Add(new Run { FootnoteId = "fn01" });
        p.AddText(" 텍스트");
        p.Runs.Add(new Run { EndnoteId = "en01" });
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

    private static byte[] WriteToBytes(PolyDonkyument doc)
    {
        using var ms = new MemoryStream();
        new HwpxWriter().Write(doc, ms);
        return ms.ToArray();
    }

    private static PolyDonkyument WriteThenRead(PolyDonkyument doc)
    {
        var bytes = WriteToBytes(doc);
        using var ms = new MemoryStream(bytes);
        return new HwpxReader().Read(ms);
    }
}

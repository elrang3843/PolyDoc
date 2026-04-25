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

    private static PolyDocument WriteThenRead(PolyDocument document)
    {
        using var ms = new MemoryStream();
        new DocxWriter().Write(document, ms);
        ms.Position = 0;
        return new DocxReader().Read(ms);
    }
}

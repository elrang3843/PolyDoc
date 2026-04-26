using PolyDoc.Codecs.Text;
using PolyDoc.Core;

namespace PolyDoc.Codecs.Text.Tests;

public class PlainTextRoundTripTests
{
    [Fact]
    public void Reader_OneLinePerParagraph_LfTerminated()
    {
        var doc = PlainTextReader.FromText("첫 줄\n두 번째\n세 번째");

        var blocks = doc.Sections[0].Blocks;
        Assert.Equal(3, blocks.Count);
        Assert.Equal("첫 줄", ((Paragraph)blocks[0]).GetPlainText());
        Assert.Equal("두 번째", ((Paragraph)blocks[1]).GetPlainText());
        Assert.Equal("세 번째", ((Paragraph)blocks[2]).GetPlainText());
    }

    [Fact]
    public void Reader_StripsCrLf()
    {
        var doc = PlainTextReader.FromText("first\r\nsecond\r\n");

        var paragraphs = doc.EnumerateParagraphs().Select(p => p.GetPlainText()).ToList();
        Assert.Equal(new[] { "first", "second" }, paragraphs);
    }

    [Fact]
    public void Writer_JoinsParagraphsWithLf()
    {
        var doc = new PolyDocument();
        var section = new Section();
        section.Blocks.Add(Paragraph.Of("a"));
        section.Blocks.Add(Paragraph.Of("b"));
        section.Blocks.Add(Paragraph.Of("c"));
        doc.Sections.Add(section);

        Assert.Equal("a\nb\nc", PlainTextWriter.ToText(doc));
    }

    [Fact]
    public void RoundTrip_PreservesAllLines()
    {
        const string original = "첫 번째\n두 번째 — 한글\n세 번째";
        var doc = PlainTextReader.FromText(original);
        Assert.Equal(original, PlainTextWriter.ToText(doc));
    }

    [Fact]
    public void Reader_RejectsNullText()
    {
        Assert.Throws<ArgumentNullException>(() => PlainTextReader.FromText(null!));
    }
}

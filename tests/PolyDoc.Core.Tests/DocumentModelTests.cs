using PolyDoc.Core;

namespace PolyDoc.Core.Tests;

public class DocumentModelTests
{
    [Fact]
    public void Empty_Document_HasOneSectionAndNoBlocks()
    {
        var doc = PolyDocument.Empty();

        Assert.Single(doc.Sections);
        Assert.Empty(doc.Sections[0].Blocks);
    }

    [Fact]
    public void Paragraph_AddText_AppendsRunWithGivenStyle()
    {
        var p = new Paragraph();
        p.AddText("hello", new RunStyle { Bold = true });

        Assert.Single(p.Runs);
        Assert.Equal("hello", p.Runs[0].Text);
        Assert.True(p.Runs[0].Style.Bold);
    }

    [Fact]
    public void Paragraph_GetPlainText_ConcatsAllRuns()
    {
        var p = new Paragraph();
        p.AddText("one ");
        p.AddText("two ");
        p.AddText("three");

        Assert.Equal("one two three", p.GetPlainText());
    }

    [Fact]
    public void EnumerateParagraphs_FlattensAcrossSections()
    {
        var doc = new PolyDocument();
        var s1 = new Section();
        s1.Blocks.Add(Paragraph.Of("a"));
        s1.Blocks.Add(Paragraph.Of("b"));
        var s2 = new Section();
        s2.Blocks.Add(Paragraph.Of("c"));
        doc.Sections.Add(s1);
        doc.Sections.Add(s2);

        var texts = doc.EnumerateParagraphs().Select(p => p.GetPlainText()).ToList();

        Assert.Equal(new[] { "a", "b", "c" }, texts);
    }

    [Theory]
    [InlineData("#FF0000", 255, 0, 0, 255)]
    [InlineData("#00FF00", 0, 255, 0, 255)]
    [InlineData("#1234567F", 0x12, 0x34, 0x56, 0x7F)]
    public void Color_FromHex_ParsesCorrectly(string hex, byte r, byte g, byte b, byte a)
    {
        var color = Color.FromHex(hex);

        Assert.Equal(r, color.R);
        Assert.Equal(g, color.G);
        Assert.Equal(b, color.B);
        Assert.Equal(a, color.A);
    }

    [Fact]
    public void Color_ToHex_OmitsAlphaWhenOpaque()
    {
        Assert.Equal("#FF8800", new Color(0xFF, 0x88, 0x00).ToHex());
        Assert.Equal("#FF880080", new Color(0xFF, 0x88, 0x00, 0x80).ToHex());
    }

    [Fact]
    public void Color_FromHex_RejectsInvalidLength()
    {
        Assert.Throws<FormatException>(() => Color.FromHex("#FFF"));
    }
}

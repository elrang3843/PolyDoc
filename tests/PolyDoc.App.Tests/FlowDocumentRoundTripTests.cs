using System.Linq;
using System.Windows;
using PolyDoc.App.Services;
using PolyDoc.Core;
using Wpf = System.Windows.Documents;
using WpfMedia = System.Windows.Media;

namespace PolyDoc.App.Tests;

public class FlowDocumentRoundTripTests
{
    [Fact]
    public void Build_PreservesFontAndSizeOnRun()
    {
        var doc = SingleParagraph(text: "큰 글씨", style: new RunStyle
        {
            FontFamily = "맑은 고딕",
            FontSizePt = 18,
        });

        var fd = FlowDocumentBuilder.Build(doc);
        var run = FirstWpfRun(fd);

        Assert.Equal("맑은 고딕", run.FontFamily.Source);
        Assert.Equal(FlowDocumentBuilder.PtToDip(18), run.FontSize, precision: 1);
    }

    [Fact]
    public void Build_PreservesBoldItalicUnderline()
    {
        var doc = SingleParagraph(text: "ABC", style: new RunStyle
        {
            Bold = true,
            Italic = true,
            Underline = true,
            Strikethrough = true,
        });

        var run = FirstWpfRun(FlowDocumentBuilder.Build(doc));

        Assert.Equal(FontWeights.Bold, run.FontWeight);
        Assert.Equal(FontStyles.Italic, run.FontStyle);
        Assert.NotNull(run.TextDecorations);
        Assert.Contains(run.TextDecorations!, d => d.Location == TextDecorationLocation.Underline);
        Assert.Contains(run.TextDecorations!, d => d.Location == TextDecorationLocation.Strikethrough);
    }

    [Fact]
    public void Build_PreservesForegroundColor()
    {
        var doc = SingleParagraph(text: "red", style: new RunStyle
        {
            Foreground = Color.FromHex("#FF3300"),
        });

        var run = FirstWpfRun(FlowDocumentBuilder.Build(doc));

        Assert.IsType<WpfMedia.SolidColorBrush>(run.Foreground);
        var brush = (WpfMedia.SolidColorBrush)run.Foreground!;
        Assert.Equal(0xFF, brush.Color.R);
        Assert.Equal(0x33, brush.Color.G);
        Assert.Equal(0x00, brush.Color.B);
    }

    [Fact]
    public void Build_HeadingParagraphHasLargerFontAndBold()
    {
        var p = new Paragraph();
        p.Style.Outline = OutlineLevel.H1;
        p.AddText("제목");
        var doc = WrapInDocument(p);

        var fd = FlowDocumentBuilder.Build(doc);
        var wpfPara = (Wpf.Paragraph)fd.Blocks.First();

        Assert.True(wpfPara.FontSize > FlowDocumentBuilder.PtToDip(12));
        Assert.Equal(FontWeights.SemiBold, wpfPara.FontWeight);
    }

    [Fact]
    public void Build_AlignmentMappedCorrectly()
    {
        var doc = new PolyDocument();
        var section = new Section();
        doc.Sections.Add(section);
        foreach (var alignment in new[] { Alignment.Left, Alignment.Center, Alignment.Right, Alignment.Justify })
        {
            var p = new Paragraph { Style = { Alignment = alignment } };
            p.AddText(alignment.ToString());
            section.Blocks.Add(p);
        }

        var fd = FlowDocumentBuilder.Build(doc);
        var wpfParas = fd.Blocks.OfType<Wpf.Paragraph>().ToList();

        Assert.Equal(TextAlignment.Left, wpfParas[0].TextAlignment);
        Assert.Equal(TextAlignment.Center, wpfParas[1].TextAlignment);
        Assert.Equal(TextAlignment.Right, wpfParas[2].TextAlignment);
        Assert.Equal(TextAlignment.Justify, wpfParas[3].TextAlignment);
    }

    [Fact]
    public void Build_BulletListBecomesWpfListWithDiscMarker()
    {
        var doc = new PolyDocument();
        var section = new Section();
        doc.Sections.Add(section);
        foreach (var text in new[] { "a", "b" })
        {
            var p = new Paragraph();
            p.Style.ListMarker = new ListMarker { Kind = ListKind.Bullet };
            p.AddText(text);
            section.Blocks.Add(p);
        }

        var fd = FlowDocumentBuilder.Build(doc);
        var list = (Wpf.List)fd.Blocks.First();

        Assert.Equal(TextMarkerStyle.Disc, list.MarkerStyle);
        Assert.Equal(2, list.ListItems.Count);
    }

    [Fact]
    public void RoundTrip_PreservesBoldAndColor()
    {
        var doc = SingleParagraph(text: "강조", style: new RunStyle
        {
            Bold = true,
            Foreground = Color.FromHex("#0066CC"),
        });

        var fd = FlowDocumentBuilder.Build(doc);
        var rebuilt = FlowDocumentParser.Parse(fd, originalForMerge: doc);

        var run = rebuilt.EnumerateParagraphs().Single().Runs.Single();
        Assert.True(run.Style.Bold);
        Assert.NotNull(run.Style.Foreground);
        Assert.Equal(0x00, run.Style.Foreground!.Value.R);
        Assert.Equal(0x66, run.Style.Foreground!.Value.G);
        Assert.Equal(0xCC, run.Style.Foreground!.Value.B);
    }

    [Fact]
    public void RoundTrip_PreservesAlignmentAndOutlineLevel()
    {
        var p = new Paragraph
        {
            Style = { Outline = OutlineLevel.H2, Alignment = Alignment.Center },
        };
        p.AddText("부제목");
        var doc = WrapInDocument(p);

        var fd = FlowDocumentBuilder.Build(doc);
        var rebuilt = FlowDocumentParser.Parse(fd, originalForMerge: doc);

        var rebuiltP = rebuilt.EnumerateParagraphs().Single();
        Assert.Equal(OutlineLevel.H2, rebuiltP.Style.Outline);
        Assert.Equal(Alignment.Center, rebuiltP.Style.Alignment);
    }

    [Fact]
    public void RoundTrip_PreservesKoreanTypographyExtrasViaMergeBase()
    {
        // FlowDocument 가 표현 못 하는 한글 조판 속성(WidthPercent, LetterSpacingPx)은
        // originalForMerge 를 통해 비파괴 보존되어야 한다.
        var p = new Paragraph();
        p.AddText("장평 90", new RunStyle { WidthPercent = 90, LetterSpacingPx = 1.5 });
        var doc = WrapInDocument(p);

        var fd = FlowDocumentBuilder.Build(doc);
        var rebuilt = FlowDocumentParser.Parse(fd, originalForMerge: doc);

        var run = rebuilt.EnumerateParagraphs().Single().Runs.Single();
        Assert.Equal(90, run.Style.WidthPercent);
        Assert.Equal(1.5, run.Style.LetterSpacingPx);
    }

    private static PolyDocument SingleParagraph(string text, RunStyle style)
    {
        var p = new Paragraph();
        p.AddText(text, style);
        return WrapInDocument(p);
    }

    private static PolyDocument WrapInDocument(Paragraph p)
    {
        var doc = new PolyDocument();
        var section = new Section();
        section.Blocks.Add(p);
        doc.Sections.Add(section);
        return doc;
    }

    private static Wpf.Run FirstWpfRun(Wpf.FlowDocument fd)
    {
        var para = fd.Blocks.OfType<Wpf.Paragraph>().First();
        return para.Inlines.OfType<Wpf.Run>().First();
    }
}

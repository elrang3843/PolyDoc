using PolyDonky.Core;
using System.Text.Json;

namespace PolyDonky.Core.Tests;

public class DocumentModelTests
{
    [Fact]
    public void Empty_Document_HasOneSectionAndNoBlocks()
    {
        var doc = PolyDonkyument.Empty();

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
        var doc = new PolyDonkyument();
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

    // ── IWPF 통합 (2026-04-29) — TextBoxObject 가 Block 트리 안에 들어왔는지 검증 ──

    [Fact]
    public void TextBoxObject_IsBlock_AndSerializesWithTextboxDiscriminator()
    {
        var tb = new TextBoxObject
        {
            OverlayXMm = 10,
            OverlayYMm = 20,
            WidthMm    = 50,
            HeightMm   = 30,
        };
        tb.SetPlainText("hello");

        // Block 으로 다형 직렬화 가능해야 한다.
        Block asBlock = tb;
        var json = JsonSerializer.Serialize(asBlock, JsonDefaults.Options);
        Assert.Contains("\"$type\": \"textbox\"", json);

        // 라운드트립
        var back = JsonSerializer.Deserialize<Block>(json, JsonDefaults.Options);
        var roundTripped = Assert.IsType<TextBoxObject>(back);
        Assert.Equal(10, roundTripped.OverlayXMm);
        Assert.Equal(20, roundTripped.OverlayYMm);
        Assert.Equal(50, roundTripped.WidthMm);
        Assert.Equal(30, roundTripped.HeightMm);
        Assert.Equal("hello", roundTripped.GetPlainText());
    }

    [Fact]
    public void Section_LegacyFloatingObjects_AreMigratedIntoBlocks()
    {
        // 옛 빌드(글상자가 Section.FloatingObjects 에 저장되던 시절) 의 JSON 을 읽으면
        // Section.Blocks 로 자동 흡수되어야 한다.
        const string legacy = """
        {
          "blocks": [],
          "floatingObjects": [
            { "$type": "textbox", "xMm": 30, "yMm": 40, "widthMm": 80, "heightMm": 60 }
          ]
        }
        """;

        var section = JsonSerializer.Deserialize<Section>(legacy, JsonDefaults.Options);

        Assert.NotNull(section);
        var tb = Assert.IsType<TextBoxObject>(Assert.Single(section!.Blocks));
        Assert.Equal(30, tb.OverlayXMm);
        Assert.Equal(40, tb.OverlayYMm);
        Assert.Equal(80, tb.WidthMm);
        Assert.Equal(60, tb.HeightMm);
    }

    [Fact]
    public void Section_DoesNotEmitFloatingObjectsField_AfterUnification()
    {
        var section = new Section();
        section.Blocks.Add(new TextBoxObject { OverlayXMm = 1, OverlayYMm = 2 });

        var json = JsonSerializer.Serialize(section, JsonDefaults.Options);

        // 통합 후 출력에는 floatingObjects 키가 없어야 함 (글상자도 blocks 안에).
        Assert.DoesNotContain("floatingObjects", json);
        Assert.Contains("\"$type\": \"textbox\"", json);
    }
}

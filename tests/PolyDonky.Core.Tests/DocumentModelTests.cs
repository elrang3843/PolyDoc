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
    // CSS 3/4-char shorthand (#RGB / #RGBA)
    [InlineData("#FFF", 255, 255, 255, 255)]
    [InlineData("#000", 0, 0, 0, 255)]
    [InlineData("#F80", 255, 136, 0, 255)]
    [InlineData("#F80A", 255, 136, 0, 170)]
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
        // 2-char and 5-char hex strings are not valid CSS color formats.
        Assert.Throws<FormatException>(() => Color.FromHex("#FF"));
        Assert.Throws<FormatException>(() => Color.FromHex("#FFFFF"));
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

    // ── 페이지-로컬 anchoring (2026-04-29) ─────────────────────────────────

    [Fact]
    public void AllOverlayBlocks_ImplementIOverlayAnchored()
    {
        // 도형·이미지·표·글상자 모두 일관된 anchoring 인터페이스를 가져야 한다.
        Assert.IsAssignableFrom<IOverlayAnchored>(new ImageBlock());
        Assert.IsAssignableFrom<IOverlayAnchored>(new ShapeObject());
        Assert.IsAssignableFrom<IOverlayAnchored>(new Table());
        Assert.IsAssignableFrom<IOverlayAnchored>(new TextBoxObject());
    }

    [Fact]
    public void AnchorPageIndex_RoundTripsThroughJson()
    {
        var tb = new TextBoxObject
        {
            AnchorPageIndex = 2,
            OverlayXMm      = 15,
            OverlayYMm      = 25,
            WidthMm         = 50,
            HeightMm        = 30,
        };

        var json = JsonSerializer.Serialize<Block>(tb, JsonDefaults.Options);
        var back = JsonSerializer.Deserialize<Block>(json, JsonDefaults.Options);

        var roundTripped = Assert.IsType<TextBoxObject>(back);
        Assert.Equal(2,  roundTripped.AnchorPageIndex);
        Assert.Equal(15, roundTripped.OverlayXMm);
        Assert.Equal(25, roundTripped.OverlayYMm);
    }

    [Fact]
    public void OverlayAnchorMigration_ConvertsContinuousYToPageLocal()
    {
        // A4 세로 297mm 페이지 기준 — 2페이지 (Y=297..594) 에 위치한 도형이
        // 옛 빌드에선 OverlayYMm = 350 (연속 Y) 으로 저장됐다고 가정.
        var doc = new PolyDonkyument();
        var section = new Section();
        section.Page = new PageSettings { SizeKind = PaperSizeKind.A4, WidthMm = 210, HeightMm = 297 };
        section.Blocks.Add(new ShapeObject
        {
            WrapMode   = ImageWrapMode.InFrontOfText,
            OverlayXMm = 50,
            OverlayYMm = 350,   // 연속 Y — 페이지 2 의 53mm 지점
        });
        doc.Sections.Add(section);

        var migrated = OverlayAnchorMigration.MigrateContinuousAnchors(doc);

        Assert.True(migrated);
        var shape = (ShapeObject)section.Blocks[0];
        Assert.Equal(1,  shape.AnchorPageIndex);  // 0-based 두 번째 페이지
        Assert.Equal(50, shape.OverlayXMm);
        Assert.Equal(53, shape.OverlayYMm);       // 350 - 297 = 53
    }

    [Fact]
    public void OverlayAnchorMigration_IsIdempotent()
    {
        // 이미 페이지-로컬인 좌표는 두 번째 호출에서 변형되지 않아야 한다.
        var doc = new PolyDonkyument();
        var section = new Section();
        section.Page = new PageSettings { SizeKind = PaperSizeKind.A4, WidthMm = 210, HeightMm = 297 };
        section.Blocks.Add(new ImageBlock
        {
            WrapMode        = ImageWrapMode.BehindText,
            AnchorPageIndex = 1,
            OverlayXMm      = 30,
            OverlayYMm      = 40,   // < pageHeight, 이미 페이지-로컬
        });
        doc.Sections.Add(section);

        OverlayAnchorMigration.MigrateContinuousAnchors(doc);
        OverlayAnchorMigration.MigrateContinuousAnchors(doc);

        var img = (ImageBlock)section.Blocks[0];
        Assert.Equal(1,  img.AnchorPageIndex);
        Assert.Equal(30, img.OverlayXMm);
        Assert.Equal(40, img.OverlayYMm);
    }

    [Fact]
    public void OverlayAnchorMigration_RecursesIntoTableCells_AndTextBoxContent()
    {
        // 표 셀 안의 도형, 글상자 안의 도형도 모두 마이그레이션 대상.
        var doc = new PolyDonkyument();
        var section = new Section();
        section.Page = new PageSettings { SizeKind = PaperSizeKind.A4, WidthMm = 210, HeightMm = 297 };

        var cell = new TableCell();
        cell.Blocks.Add(new ShapeObject
        {
            WrapMode = ImageWrapMode.InFrontOfText,
            OverlayYMm = 600,   // 페이지 3
        });
        var row = new TableRow(); row.Cells.Add(cell);
        var table = new Table(); table.Rows.Add(row);
        section.Blocks.Add(table);

        var tb = new TextBoxObject();
        tb.Content.Clear();
        tb.Content.Add(new ImageBlock
        {
            WrapMode = ImageWrapMode.BehindText,
            OverlayYMm = 320,   // 페이지 2
        });
        section.Blocks.Add(tb);

        doc.Sections.Add(section);

        OverlayAnchorMigration.MigrateContinuousAnchors(doc);

        var nestedShape = (ShapeObject)cell.Blocks[0];
        Assert.Equal(2, nestedShape.AnchorPageIndex);
        Assert.Equal(6, nestedShape.OverlayYMm);  // 600 - 2*297 = 6

        var nestedImg = (ImageBlock)tb.Content[0];
        Assert.Equal(1,  nestedImg.AnchorPageIndex);
        Assert.Equal(23, nestedImg.OverlayYMm);   // 320 - 297 = 23
    }
}

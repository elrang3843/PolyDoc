using System.IO.Compression;
using System.Xml.Linq;
using PolyDonky.Codecs.Hwpx;
using PolyDonky.Core;

namespace PolyDonky.Codecs.Hwpx.Tests;

/// <summary>
/// Writer 출력 XML 구조 회귀 방지. 라운드트립이 아니라 HWPX 패키지 안의 header.xml /
/// content.hpf / section*.xml 을 직접 파싱해 한컴 호환에 필요한 속성·자식·순서를 점검한다.
/// 각 케이스는 <c>claude/hwpx-writer-g3-D3qKT</c> 세션에서 실파일·OWPML 모델·HwpForge ground
/// truth 로 검증한 동작을 담는다.
/// </summary>
public class HwpxStructuralTests
{
    private static readonly XNamespace Hp  = HwpxNamespaces.Paragraph;
    private static readonly XNamespace Hh  = HwpxNamespaces.Head;
    private static readonly XNamespace Hc  = HwpxNamespaces.Common;
    private static readonly XNamespace Opf = HwpxNamespaces.OpfPackage;

    // ── multi-section secPr ─────────────────────────────────────────────────────

    [Fact]
    public void EverySection_GetsSecPrInjected()
    {
        var doc = new PolyDonkyument();
        for (int i = 0; i < 3; i++)
        {
            var s = new Section();
            s.Blocks.Add(Paragraph.Of($"section{i}"));
            doc.Sections.Add(s);
        }

        var pkg = WritePackage(doc);
        for (int i = 0; i < 3; i++)
        {
            var sec = ReadXml(pkg, $"Contents/section{i}.xml");
            var secPrs = sec.Descendants(Hp + "secPr").ToList();
            Assert.True(secPrs.Count >= 1,
                $"section{i}.xml must contain at least one hp:secPr (got {secPrs.Count})");
        }
    }

    // ── picture (hc:img namespace + outer attrs + manifest) ────────────────────

    [Fact]
    public void Picture_UsesHcImgNamespace_NotHpImg()
    {
        var doc = NewDocWithImage(out _);
        var pkg = WritePackage(doc);
        var sec = ReadXml(pkg, "Contents/section0.xml");
        var pic = sec.Descendants(Hp + "pic").Single();

        Assert.Single(pic.Elements(Hc + "img"));
        Assert.Empty(pic.Elements(Hp + "img"));
    }

    [Fact]
    public void Picture_HasRequiredOuterAttributes()
    {
        var doc = NewDocWithImage(out _);
        var pic = ReadXml(WritePackage(doc), "Contents/section0.xml")
            .Descendants(Hp + "pic").Single();

        foreach (var name in new[] {
            "id", "zOrder", "numberingType", "textWrap", "textFlow",
            "lock", "dropcapstyle", "href", "groupLevel", "instid", "reverse" })
            Assert.NotNull(pic.Attribute(name));
        Assert.Equal("PICTURE", pic.Attribute("numberingType")!.Value);
        // ratio attribute belongs to Rectangle, not Picture (per OWPML CPictureType).
        Assert.Null(pic.Attribute("ratio"));
    }

    [Fact]
    public void Picture_HasImgRectImgClipImgDimEffects()
    {
        var doc = NewDocWithImage(out _);
        var pic = ReadXml(WritePackage(doc), "Contents/section0.xml")
            .Descendants(Hp + "pic").Single();
        Assert.Single(pic.Elements(Hp + "imgRect"));
        Assert.Single(pic.Elements(Hp + "imgClip"));
        Assert.Single(pic.Elements(Hp + "imgDim"));
        Assert.Single(pic.Elements(Hp + "effects"));
        // Picture is not an AbstractDrawingObjectType — must NOT have shadow or fillBrush.
        Assert.Empty(pic.Elements(Hp + "shadow"));
        Assert.Empty(pic.Elements(Hc + "fillBrush"));
    }

    [Fact]
    public void Picture_ImgClip_RightBottomEqualToWidthHeight()
    {
        // imgClip 이 (0,0,0,0) 이면 한컴이 이미지를 안 보이게 클리핑.
        var doc = NewDocWithImage(out _);
        var pic = ReadXml(WritePackage(doc), "Contents/section0.xml")
            .Descendants(Hp + "pic").Single();
        var clip = pic.Element(Hp + "imgClip")!;
        Assert.Equal("0", clip.Attribute("left")!.Value);
        Assert.Equal("0", clip.Attribute("top")!.Value);
        Assert.NotEqual("0", clip.Attribute("right")!.Value);
        Assert.NotEqual("0", clip.Attribute("bottom")!.Value);
    }

    [Fact]
    public void Manifest_BinDataItem_HasIsEmbededAttribute()
    {
        var doc = NewDocWithImage(out _);
        var pkg = WritePackage(doc);
        var hpf = ReadXml(pkg, "Contents/content.hpf");
        var item = hpf.Descendants(Opf + "item")
            .Single(e => e.Attribute("href")!.Value.StartsWith("BinData/"));
        Assert.Equal("1", item.Attribute("isEmbeded")!.Value);
    }

    [Fact]
    public void Header_HasNoBinDataList_EvenWithEmbeddedImage()
    {
        // 실 한컴 파일은 binDataList 없이 manifest 의 opf:item id 만 사용.
        var doc = NewDocWithImage(out _);
        var hdr = ReadXml(WritePackage(doc), "Contents/header.xml");
        Assert.Empty(hdr.Descendants(Hh + "binDataList"));
    }

    // ── shape kind mapping ──────────────────────────────────────────────────────

    [Theory]
    [InlineData(ShapeKind.Line,           "line")]
    [InlineData(ShapeKind.Rectangle,      "rect")]
    [InlineData(ShapeKind.RoundedRect,    "rect")]
    [InlineData(ShapeKind.Ellipse,        "ellipse")]
    [InlineData(ShapeKind.Triangle,       "polygon")]
    [InlineData(ShapeKind.Polygon,        "polygon")]
    [InlineData(ShapeKind.Polyline,       "polygon")]
    [InlineData(ShapeKind.RegularPolygon, "polygon")]
    [InlineData(ShapeKind.Star,           "polygon")]
    [InlineData(ShapeKind.Spline,         "polygon")]
    [InlineData(ShapeKind.ClosedSpline,   "polygon")]
    public void Shape_MapsToExpectedHwpxElement(ShapeKind kind, string expectedLocalName)
    {
        var doc = DocWithShape(new ShapeObject { Kind = kind, WidthMm = 30, HeightMm = 20 });
        var sec = ReadXml(WritePackage(doc), "Contents/section0.xml");
        Assert.Single(sec.Descendants(Hp + expectedLocalName));
    }

    [Fact]
    public void Shape_AllUseTextWrapInFrontOfText_NotTopAndBottom()
    {
        // 본문 흐름에 영향 주지 않으려면 IN_FRONT_OF_TEXT.
        var doc = DocWithShape(new ShapeObject { Kind = ShapeKind.Rectangle, WidthMm = 30, HeightMm = 20 });
        var rect = ReadXml(WritePackage(doc), "Contents/section0.xml")
            .Descendants(Hp + "rect").Single();
        Assert.Equal("IN_FRONT_OF_TEXT", rect.Attribute("textWrap")!.Value);
    }

    [Fact]
    public void Shape_PosIsAnchored_NotInline()
    {
        // 모든 도형은 anchored overlay (treatAsChar=0) — inline 이면 페이지 흐름 깨짐.
        var doc = DocWithShape(new ShapeObject { Kind = ShapeKind.Rectangle, WidthMm = 30, HeightMm = 20 });
        var pos = ReadXml(WritePackage(doc), "Contents/section0.xml")
            .Descendants(Hp + "rect").Single()
            .Element(Hp + "pos")!;
        Assert.Equal("0",     pos.Attribute("treatAsChar")!.Value);
        Assert.Equal("PAPER", pos.Attribute("vertRelTo")!.Value);
        Assert.Equal("PAPER", pos.Attribute("horzRelTo")!.Value);
    }

    [Fact]
    public void Rectangle_HasFourCornerPoints()
    {
        var doc = DocWithShape(new ShapeObject { Kind = ShapeKind.Rectangle, WidthMm = 30, HeightMm = 20 });
        var rect = ReadXml(WritePackage(doc), "Contents/section0.xml")
            .Descendants(Hp + "rect").Single();
        Assert.Single(rect.Elements(Hc + "pt0"));
        Assert.Single(rect.Elements(Hc + "pt1"));
        Assert.Single(rect.Elements(Hc + "pt2"));
        Assert.Single(rect.Elements(Hc + "pt3"));
    }

    [Fact]
    public void Ellipse_UsesCenterAndAxes_NotCorners()
    {
        // OWPML CEllipseType 은 center/ax1/ax2/start1/start2/end1/end2 를 사용.
        var doc = DocWithShape(new ShapeObject { Kind = ShapeKind.Ellipse, WidthMm = 30, HeightMm = 20 });
        var elp = ReadXml(WritePackage(doc), "Contents/section0.xml")
            .Descendants(Hp + "ellipse").Single();
        Assert.Single(elp.Elements(Hc + "center"));
        Assert.Single(elp.Elements(Hc + "ax1"));
        Assert.Single(elp.Elements(Hc + "ax2"));
        Assert.Empty(elp.Elements(Hc + "pt0"));
    }

    [Fact]
    public void RoundedRect_RatioAttribute_NonZero()
    {
        var doc = DocWithShape(new ShapeObject {
            Kind = ShapeKind.RoundedRect, WidthMm = 30, HeightMm = 20, CornerRadiusMm = 5 });
        var rect = ReadXml(WritePackage(doc), "Contents/section0.xml")
            .Descendants(Hp + "rect").Single();
        var ratio = int.Parse(rect.Attribute("ratio")!.Value);
        Assert.True(ratio > 0, $"RoundedRect ratio must be > 0, got {ratio}");
    }

    [Fact]
    public void Line_HasNoFillBrush_EvenWithFillColor()
    {
        // 선 도형은 시각적으로 채울 수 없으므로 fillBrush 자체를 출력하지 않음.
        // FillColor 가 설정되어 있어도 Line 은 fillBrush 생략 (한컴 ground truth 패턴).
        var doc = DocWithShape(new ShapeObject {
            Kind = ShapeKind.Line, FillColor = "#FF0000", WidthMm = 30, HeightMm = 5 });
        var line = ReadXml(WritePackage(doc), "Contents/section0.xml")
            .Descendants(Hp + "line").Single();
        Assert.Empty(line.Elements(Hc + "fillBrush"));
    }

    [Fact]
    public void OpenPolygon_DoesNotAddClosingPoint()
    {
        // Polyline = 열린 도형 — 첫 점 반복 없이 끝.
        var open = new ShapeObject { Kind = ShapeKind.Polyline, WidthMm = 30, HeightMm = 20 };
        open.Points.Add(new ShapePoint { X = 0,  Y = 0  });
        open.Points.Add(new ShapePoint { X = 30, Y = 20 });
        open.Points.Add(new ShapePoint { X = 0,  Y = 20 });
        var doc = DocWithShape(open);
        var poly = ReadXml(WritePackage(doc), "Contents/section0.xml")
            .Descendants(Hp + "polygon").Single();
        var pts = poly.Elements(Hc + "pt").ToList();
        // 3 input points, no closing repeat.
        Assert.Equal(3, pts.Count);
    }

    [Fact]
    public void ClosedPolygon_AddsClosingPoint()
    {
        // Polygon = 닫힌 도형 — 마지막에 첫 점 반복.
        var closed = new ShapeObject { Kind = ShapeKind.Polygon, WidthMm = 30, HeightMm = 20 };
        closed.Points.Add(new ShapePoint { X = 0,  Y = 0  });
        closed.Points.Add(new ShapePoint { X = 30, Y = 20 });
        closed.Points.Add(new ShapePoint { X = 0,  Y = 20 });
        var doc = DocWithShape(closed);
        var poly = ReadXml(WritePackage(doc), "Contents/section0.xml")
            .Descendants(Hp + "polygon").Single();
        var pts = poly.Elements(Hc + "pt").ToList();
        // 3 input points + 1 closing repeat = 4.
        Assert.Equal(4, pts.Count);
    }

    [Fact]
    public void Shape_HostingParagraph_MovedAfterFirstParagraph()
    {
        // AnchorPageIndex==0 도형은 첫 단락 직후로 재배치 — page 1 anchor 보장.
        var doc = new PolyDonkyument();
        var section = new Section();
        for (int i = 0; i < 5; i++)
            section.Blocks.Add(Paragraph.Of($"text {i}"));
        section.Blocks.Add(new ShapeObject { Kind = ShapeKind.Rectangle, WidthMm = 30, HeightMm = 20 });
        doc.Sections.Add(section);

        var sec = ReadXml(WritePackage(doc), "Contents/section0.xml");
        var paragraphs = sec.Root!.Elements(Hp + "p").ToList();
        // 첫 단락 = "text 0", 두번째 단락에 도형 hp:rect 가 있어야 함.
        Assert.Contains(paragraphs[0].Descendants(Hp + "t"),
            t => t.Value == "text 0");
        Assert.Single(paragraphs[1].Descendants(Hp + "rect"));
    }

    // ── table dynamic borderFill ───────────────────────────────────────────────

    [Fact]
    public void Table_BorderColorAndThickness_RegisteredInHeader()
    {
        var t = new Table {
            BorderColor = "#FF0000", BorderThicknessPt = 2.0,
        };
        t.Columns.Add(new TableColumn { WidthMm = 30 });
        t.Rows.Add(MakeRow(new[] { "A" }));
        var doc = NewDocWith(t);
        var pkg = WritePackage(doc);
        var hdr = ReadXml(pkg, "Contents/header.xml");

        // 표 자체 borderFill 이 등록되어야 — color=#FF0000 가진 borderFill 존재.
        var redBorder = hdr.Descendants(Hh + "borderFill")
            .FirstOrDefault(bf => bf.Element(Hh + "topBorder")?.Attribute("color")?.Value == "#FF0000");
        Assert.NotNull(redBorder);
    }

    [Fact]
    public void Table_BorderWidth_SnappedToStandardValue()
    {
        // 2pt = 0.7056mm — 비표준. 표준 집합(0.7)으로 snap 되어야.
        var t = new Table { BorderColor = "#000000", BorderThicknessPt = 2.0 };
        t.Columns.Add(new TableColumn { WidthMm = 30 });
        t.Rows.Add(MakeRow(new[] { "A" }));
        var doc = NewDocWith(t);
        var hdr = ReadXml(WritePackage(doc), "Contents/header.xml");

        // 모든 width 값이 표준 집합 안에 있어야.
        var stdSet = new HashSet<string> {
            "0.1 mm", "0.12 mm", "0.15 mm", "0.2 mm", "0.25 mm",
            "0.3 mm", "0.4 mm", "0.5 mm", "0.6 mm", "0.7 mm",
            "1 mm", "1.5 mm", "2 mm", "3 mm", "4 mm", "5 mm" };
        foreach (var w in hdr.Descendants(Hh + "borderFill")
                    .SelectMany(bf => bf.Elements().Where(e =>
                        e.Name.LocalName.EndsWith("Border") || e.Name.LocalName == "diagonal"))
                    .Select(e => e.Attribute("width")?.Value))
        {
            if (w is null) continue;
            Assert.Contains(w, stdSet);
        }
    }

    [Fact]
    public void Table_BackgroundColor_AddedAsFillBrush()
    {
        var t = new Table {
            BorderColor = "#000000", BorderThicknessPt = 0.5,
            BackgroundColor = "#EEEEEE",
        };
        t.Columns.Add(new TableColumn { WidthMm = 30 });
        t.Rows.Add(MakeRow(new[] { "A" }));
        var doc = NewDocWith(t);
        var hdr = ReadXml(WritePackage(doc), "Contents/header.xml");

        // 어떤 borderFill 이라도 winBrush.faceColor=#EEEEEE 를 가진 항목이 있어야.
        var hasFill = hdr.Descendants(Hh + "borderFill")
            .Any(bf => bf.Descendants(Hc + "winBrush")
                .Any(wb => wb.Attribute("faceColor")?.Value == "#EEEEEE"));
        Assert.True(hasFill, "Table BackgroundColor must produce a winBrush faceColor in some borderFill");
    }

    [Fact]
    public void Table_HpTbl_HasFullOuterAttributes()
    {
        var t = new Table();
        t.Columns.Add(new TableColumn { WidthMm = 30 });
        t.Rows.Add(MakeRow(new[] { "A" }));
        var doc = NewDocWith(t);
        var sec = ReadXml(WritePackage(doc), "Contents/section0.xml");
        var tbl = sec.Descendants(Hp + "tbl").Single();

        foreach (var name in new[] {
            "id", "zOrder", "numberingType", "textWrap", "textFlow",
            "lock", "dropcapstyle", "pageBreak", "repeatHeader",
            "rowCnt", "colCnt", "cellSpacing", "borderFillIDRef", "noAdjust" })
            Assert.NotNull(tbl.Attribute(name));
        Assert.Equal("TABLE", tbl.Attribute("numberingType")!.Value);
    }

    [Fact]
    public void Table_Cell_SubList_HasRequiredAttributes()
    {
        // hp:subList 의 10개 속성이 모두 있어야 한컴이 셀 본문 인식.
        var t = new Table();
        t.Columns.Add(new TableColumn { WidthMm = 30 });
        t.Rows.Add(MakeRow(new[] { "A" }));
        var doc = NewDocWith(t);
        var subList = ReadXml(WritePackage(doc), "Contents/section0.xml")
            .Descendants(Hp + "tc").Single()
            .Element(Hp + "subList")!;

        foreach (var name in new[] {
            "id", "textDirection", "lineWrap", "vertAlign",
            "linkListIDRef", "linkListNextIDRef",
            "textWidth", "textHeight", "hasTextRef", "hasNumRef" })
            Assert.NotNull(subList.Attribute(name));
    }

    // ── helpers ────────────────────────────────────────────────────────────────

    private static PolyDonkyument NewDocWithImage(out byte[] tinyPng)
    {
        tinyPng = new byte[]
        {
            0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
            0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
            0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,
            0x08,0x06,0x00,0x00,0x00,0x1F,0x15,0xC4,
            0x89,0x00,0x00,0x00,0x0D,0x49,0x44,0x41,
            0x54,0x78,0x9C,0x62,0x00,0x01,0x00,0x00,
            0x05,0x00,0x01,0x0D,0x0A,0x2D,0xB4,0x00,
            0x00,0x00,0x00,0x49,0x45,0x4E,0x44,0xAE,
            0x42,0x60,0x82,
        };
        var doc = new PolyDonkyument();
        var section = new Section();
        section.Blocks.Add(Paragraph.Of("hello"));
        section.Blocks.Add(new ImageBlock {
            Data = tinyPng, MediaType = "image/png",
            WidthMm = 20, HeightMm = 20,
        });
        doc.Sections.Add(section);
        return doc;
    }

    private static PolyDonkyument DocWithShape(ShapeObject shape)
    {
        var doc = new PolyDonkyument();
        var section = new Section();
        section.Blocks.Add(Paragraph.Of("anchor"));
        section.Blocks.Add(shape);
        doc.Sections.Add(section);
        return doc;
    }

    private static PolyDonkyument NewDocWith(Block block)
    {
        var doc = new PolyDonkyument();
        var section = new Section();
        section.Blocks.Add(block);
        doc.Sections.Add(section);
        return doc;
    }

    private static TableRow MakeRow(string[] cellTexts)
    {
        var row = new TableRow();
        foreach (var text in cellTexts)
            row.Cells.Add(new TableCell { Blocks = { Paragraph.Of(text) } });
        return row;
    }

    private static byte[] WritePackage(PolyDonkyument doc)
    {
        using var ms = new MemoryStream();
        new HwpxWriter().Write(doc, ms);
        return ms.ToArray();
    }

    private static XDocument ReadXml(byte[] hwpxBytes, string entryPath)
    {
        using var ms = new MemoryStream(hwpxBytes);
        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
        var entry = zip.GetEntry(entryPath)
            ?? throw new InvalidOperationException($"missing zip entry: {entryPath}");
        using var stream = entry.Open();
        return XDocument.Load(stream);
    }
}

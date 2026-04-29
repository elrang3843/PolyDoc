using System.Threading;
using System.Windows;
using PolyDonky.App.Pagination;
using PolyDonky.Core;

namespace PolyDonky.App.Tests;

// ── 모델 구조 테스트 (WPF·STA 불필요) ─────────────────────────────────────

public class PaginatedDocumentModelTests
{
    [Fact]
    public void PaginatedPage_PageNumber_IsOneBasedIndex()
    {
        var page = new PaginatedPage { PageIndex = 0 };
        Assert.Equal(1, page.PageNumber);

        var page3 = new PaginatedPage { PageIndex = 2 };
        Assert.Equal(3, page3.PageNumber);
    }

    [Fact]
    public void PaginatedDocument_PageCount_MatchesPagesCount()
    {
        var doc = new PolyDonkyument();
        doc.Sections.Add(new Section());

        var paged = new PaginatedDocument
        {
            Source       = doc,
            PageSettings = new PageSettings(),
            Pages        = new[] { new PaginatedPage { PageIndex = 0 }, new PaginatedPage { PageIndex = 1 } },
        };

        Assert.Equal(2, paged.PageCount);
    }

    [Fact]
    public void BlockOnPage_BodyLocalRect_DefaultsToEmpty()
    {
        var para = new Paragraph();
        var bop  = new BlockOnPage { Source = para, PageIndex = 0 };

        Assert.Equal(Rect.Empty, bop.BodyLocalRect);
    }

    [Fact]
    public void OverlayOnPage_StoredCoordinates_RoundTrip()
    {
        var tb  = new TextBoxObject { OverlayXMm = 30, OverlayYMm = 50, AnchorPageIndex = 1 };
        var oop = new OverlayOnPage { Source = tb, AnchorPageIndex = 1, XMm = 30, YMm = 50 };

        Assert.Equal(30, oop.XMm);
        Assert.Equal(50, oop.YMm);
        Assert.Equal(1,  oop.AnchorPageIndex);
    }
}

// ── IsOverlayMode 판별 테스트 (STA 불필요) ───────────────────────────────

public class OverlayModeDetectionTests
{
    [Fact]
    public void TextBoxObject_IsAlwaysOverlay()
    {
        Assert.True(FlowDocumentPaginationAdapter.IsOverlayMode(new TextBoxObject()));
    }

    [Fact]
    public void ImageBlock_InFrontOfText_IsOverlay()
    {
        var img = new ImageBlock { WrapMode = ImageWrapMode.InFrontOfText };
        Assert.True(FlowDocumentPaginationAdapter.IsOverlayMode(img));
    }

    [Fact]
    public void ImageBlock_Inline_IsNotOverlay()
    {
        var img = new ImageBlock { WrapMode = ImageWrapMode.Inline };
        Assert.False(FlowDocumentPaginationAdapter.IsOverlayMode(img));
    }

    [Fact]
    public void ShapeObject_BehindText_IsOverlay()
    {
        var shp = new ShapeObject { WrapMode = ImageWrapMode.BehindText };
        Assert.True(FlowDocumentPaginationAdapter.IsOverlayMode(shp));
    }

    [Fact]
    public void Table_BlockMode_IsNotOverlay()
    {
        var tbl = new Table { WrapMode = TableWrapMode.Block };
        Assert.False(FlowDocumentPaginationAdapter.IsOverlayMode(tbl));
    }

    [Fact]
    public void Table_NonBlockMode_IsOverlay()
    {
        var tbl = new Table { WrapMode = TableWrapMode.InFrontOfText };
        Assert.True(FlowDocumentPaginationAdapter.IsOverlayMode(tbl));
    }

    [Fact]
    public void Paragraph_IsNeverOverlay()
    {
        Assert.False(FlowDocumentPaginationAdapter.IsOverlayMode(new Paragraph()));
    }
}

// ── FlowDocumentPaginationAdapter 통합 테스트 (STA 필수) ──────────────────

public class FlowDocumentPaginationAdapterTests
{
    // FlowDocument / DocumentPaginator 는 DependencyObject → STA 스레드 필수.
    private static void RunOnSta(Action action)
    {
        Exception? caught = null;
        var t = new Thread(() =>
        {
            try { action(); }
            catch (Exception ex) { caught = ex; }
        });
        t.SetApartmentState(ApartmentState.STA);
        t.Start();
        t.Join();
        if (caught is not null)
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(caught).Throw();
    }

    // ── 기본 구조 ──────────────────────────────────────────────────────────

    [Fact]
    public void Paginate_EmptyDocument_ReturnsAtLeastOnePage()
    {
        RunOnSta(() =>
        {
            var doc = PolyDonkyument.Empty();
            var result = FlowDocumentPaginationAdapter.Paginate(doc);

            Assert.True(result.PageCount >= 1);
            Assert.Equal(result.PageCount, result.Pages.Count);
        });
    }

    [Fact]
    public void Paginate_ReturnsCorrectSource()
    {
        RunOnSta(() =>
        {
            var doc    = PolyDonkyument.Empty();
            var result = FlowDocumentPaginationAdapter.Paginate(doc);

            Assert.Same(doc, result.Source);
        });
    }

    [Fact]
    public void Paginate_ReturnsCorrectPageSettings()
    {
        RunOnSta(() =>
        {
            var page = new PageSettings { SizeKind = PaperSizeKind.A4 };
            var doc  = new PolyDonkyument();
            doc.Sections.Add(new Section { Page = page });

            var result = FlowDocumentPaginationAdapter.Paginate(doc);

            Assert.Same(page, result.PageSettings);
        });
    }

    [Fact]
    public void Paginate_PageSettingsOverride_UsedInsteadOfDocumentPage()
    {
        RunOnSta(() =>
        {
            var docPage      = new PageSettings { SizeKind = PaperSizeKind.A4 };
            var overridePage = new PageSettings { SizeKind = PaperSizeKind.A5 };
            var doc          = new PolyDonkyument();
            doc.Sections.Add(new Section { Page = docPage });

            var result = FlowDocumentPaginationAdapter.Paginate(doc, overridePage);

            Assert.Same(overridePage, result.PageSettings);
        });
    }

    // ── 페이지 인덱스 연속성 ──────────────────────────────────────────────

    [Fact]
    public void Paginate_PagesHaveConsecutiveZeroBasedIndices()
    {
        RunOnSta(() =>
        {
            var doc = BuildMultiParagraphDocument(5);
            var result = FlowDocumentPaginationAdapter.Paginate(doc);

            for (int i = 0; i < result.PageCount; i++)
            {
                Assert.Equal(i,     result.Pages[i].PageIndex);
                Assert.Equal(i + 1, result.Pages[i].PageNumber);
            }
        });
    }

    // ── 본문 블록 배정 ────────────────────────────────────────────────────

    [Fact]
    public void Paginate_SingleParagraph_BlockAssignedToSomePage()
    {
        RunOnSta(() =>
        {
            var para = new Paragraph();
            para.AddText("테스트");
            var doc = WrapInDocument(para);

            var result      = FlowDocumentPaginationAdapter.Paginate(doc);
            var allBlocks   = result.Pages.SelectMany(p => p.BodyBlocks).ToList();

            Assert.Contains(allBlocks, bop => ReferenceEquals(bop.Source, para));
        });
    }

    [Fact]
    public void Paginate_TotalBodyBlockCount_MatchesNonOverlayBlockCount()
    {
        RunOnSta(() =>
        {
            var doc     = BuildMultiParagraphDocument(4);
            var section = doc.Sections[0];
            var nonOverlayCount = section.Blocks.Count(b => !FlowDocumentPaginationAdapter.IsOverlayMode(b));

            var result = FlowDocumentPaginationAdapter.Paginate(doc);
            var totalBodyBlocks = result.Pages.Sum(p => p.BodyBlocks.Count);

            Assert.Equal(nonOverlayCount, totalBodyBlocks);
        });
    }

    // ── BodyLocalRect 측정 ────────────────────────────────────────────────

    [Fact]
    public void Paginate_SingleParagraph_BodyLocalRectIsNotEmpty()
    {
        RunOnSta(() =>
        {
            var para = new Paragraph();
            para.AddText("본문 단락");
            var doc    = WrapInDocument(para);
            var result = FlowDocumentPaginationAdapter.Paginate(doc);

            var bop = result.Pages.SelectMany(p => p.BodyBlocks)
                                  .First(b => ReferenceEquals(b.Source, para));

            Assert.NotEqual(System.Windows.Rect.Empty, bop.BodyLocalRect);
        });
    }

    [Fact]
    public void Paginate_BodyLocalRect_YIsNonNegativeAndWithinBodyHeight()
    {
        RunOnSta(() =>
        {
            var page   = new PageSettings { SizeKind = PaperSizeKind.A4 };
            var para   = new Paragraph();
            para.AddText("범위 검증");
            var doc    = new PolyDonkyument();
            var sec    = new Section { Page = page };
            sec.Blocks.Add(para);
            doc.Sections.Add(sec);

            var result  = FlowDocumentPaginationAdapter.Paginate(doc);
            var bop     = result.Pages.SelectMany(p => p.BodyBlocks)
                                      .First(b => ReferenceEquals(b.Source, para));
            var rect    = bop.BodyLocalRect;

            Assert.False(rect == System.Windows.Rect.Empty);
            Assert.True(rect.Y >= 0,           $"Y={rect.Y} 는 0 이상이어야 함");
            Assert.True(rect.Height > 0,        $"Height={rect.Height} 는 양수여야 함");
            Assert.True(rect.Width  > 0,        $"Width={rect.Width} 는 양수여야 함");
        });
    }

    [Fact]
    public void Paginate_MultiParagraph_SecondBlockHasGreaterOrEqualY()
    {
        RunOnSta(() =>
        {
            var p1 = new Paragraph(); p1.AddText("첫째 단락");
            var p2 = new Paragraph(); p2.AddText("둘째 단락");
            var doc = WrapInDocument(p1, p2);

            var result = FlowDocumentPaginationAdapter.Paginate(doc);

            // 같은 페이지에 있을 때만 Y 순서를 검증
            var page0Blocks = result.Pages[0].BodyBlocks.ToList();
            var bop1 = page0Blocks.FirstOrDefault(b => ReferenceEquals(b.Source, p1));
            var bop2 = page0Blocks.FirstOrDefault(b => ReferenceEquals(b.Source, p2));

            if (bop1 is not null && bop2 is not null
                && bop1.BodyLocalRect != System.Windows.Rect.Empty
                && bop2.BodyLocalRect != System.Windows.Rect.Empty)
            {
                Assert.True(bop2.BodyLocalRect.Y >= bop1.BodyLocalRect.Y,
                    $"두 번째 단락 Y({bop2.BodyLocalRect.Y:F1}) 가 첫째({bop1.BodyLocalRect.Y:F1}) 보다 작다");
            }
        });
    }

    // ── 오버레이 블록 배정 ────────────────────────────────────────────────

    [Fact]
    public void Paginate_TextBoxObject_PlacedOnAnchorPage()
    {
        RunOnSta(() =>
        {
            var tb = new TextBoxObject { AnchorPageIndex = 0, OverlayXMm = 10, OverlayYMm = 20 };
            tb.Content.Clear();
            var tbPara = new Paragraph(); tbPara.AddText("글상자");
            tb.Content.Add(tbPara);

            var doc = new PolyDonkyument();
            var section = new Section();
            var bodyPara = new Paragraph(); bodyPara.AddText("본문");
            section.Blocks.Add(bodyPara);
            section.Blocks.Add(tb);
            doc.Sections.Add(section);

            var result   = FlowDocumentPaginationAdapter.Paginate(doc);
            var overlays = result.Pages[0].OverlayBlocks;

            Assert.Single(overlays, o => ReferenceEquals(o.Source, tb));
            Assert.Equal(10, overlays[0].XMm);
            Assert.Equal(20, overlays[0].YMm);
        });
    }

    [Fact]
    public void Paginate_OverlayOnPage2_PageCountAtLeastTwo()
    {
        RunOnSta(() =>
        {
            var shp = new ShapeObject
            {
                WrapMode        = ImageWrapMode.InFrontOfText,
                AnchorPageIndex = 1,
                OverlayXMm      = 50,
                OverlayYMm      = 60,
                WidthMm         = 40,
                HeightMm        = 20,
            };

            var doc     = new PolyDonkyument();
            var section = new Section();
            var para    = new Paragraph(); para.AddText("본문");
            section.Blocks.Add(para);
            section.Blocks.Add(shp);
            doc.Sections.Add(section);

            var result = FlowDocumentPaginationAdapter.Paginate(doc);

            Assert.True(result.PageCount >= 2, $"AnchorPageIndex=1 이므로 최소 2페이지여야 한다. 실제: {result.PageCount}");
            Assert.Contains(result.Pages[1].OverlayBlocks, o => ReferenceEquals(o.Source, shp));
        });
    }

    [Fact]
    public void Paginate_ImageOverlayNotInBodyBlocks()
    {
        RunOnSta(() =>
        {
            var img = new ImageBlock
            {
                WrapMode        = ImageWrapMode.InFrontOfText,
                AnchorPageIndex = 0,
                WidthMm         = 30,
                HeightMm        = 30,
            };
            var doc     = WrapInDocument(img);
            var result  = FlowDocumentPaginationAdapter.Paginate(doc);
            var allBody = result.Pages.SelectMany(p => p.BodyBlocks);

            Assert.DoesNotContain(allBody, bop => ReferenceEquals(bop.Source, img));
        });
    }

    // ── 다중 섹션 ─────────────────────────────────────────────────────────

    [Fact]
    public void Paginate_MultiSection_OverlaysFromAllSectionsCollected()
    {
        RunOnSta(() =>
        {
            var tb1 = new TextBoxObject { AnchorPageIndex = 0, OverlayXMm = 10, OverlayYMm = 10 };
            var tb2 = new TextBoxObject { AnchorPageIndex = 0, OverlayXMm = 50, OverlayYMm = 50 };

            var doc = new PolyDonkyument();

            var s1 = new Section();
            var p1 = new Paragraph(); p1.AddText("섹션 1");
            s1.Blocks.Add(p1);
            s1.Blocks.Add(tb1);
            doc.Sections.Add(s1);

            var s2 = new Section();
            var p2 = new Paragraph(); p2.AddText("섹션 2");
            s2.Blocks.Add(p2);
            s2.Blocks.Add(tb2);
            doc.Sections.Add(s2);

            var result      = FlowDocumentPaginationAdapter.Paginate(doc);
            var allOverlays = result.Pages.SelectMany(p => p.OverlayBlocks).ToList();

            Assert.Contains(allOverlays, o => ReferenceEquals(o.Source, tb1));
            Assert.Contains(allOverlays, o => ReferenceEquals(o.Source, tb2));
        });
    }

    // ── 헬퍼 ─────────────────────────────────────────────────────────────

    private static PolyDonkyument WrapInDocument(params Block[] blocks)
    {
        var doc     = new PolyDonkyument();
        var section = new Section();
        foreach (var b in blocks) section.Blocks.Add(b);
        doc.Sections.Add(section);
        return doc;
    }

    private static PolyDonkyument BuildMultiParagraphDocument(int count)
    {
        var doc     = new PolyDonkyument();
        var section = new Section();
        for (int i = 0; i < count; i++)
        {
            var p = new Paragraph();
            p.AddText($"단락 {i + 1}");
            section.Blocks.Add(p);
        }
        doc.Sections.Add(section);
        return doc;
    }
}

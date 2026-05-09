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

    // ── 강제 페이지 나누기 (ForcePageBreakBefore) ─────────────────────────────

    [Fact]
    public void Paginate_ForcePageBreakBefore_PutsParagraphOnNextPage()
    {
        RunOnSta(() =>
        {
            var doc     = new PolyDonkyument();
            var section = new Section();

            var p1 = new Paragraph(); p1.AddText("첫 단락");
            section.Blocks.Add(p1);

            var p2 = new Paragraph();
            p2.Style.ForcePageBreakBefore = true;
            p2.AddText("강제 페이지 나누기 후 단락");
            section.Blocks.Add(p2);

            doc.Sections.Add(section);

            var result = FlowDocumentPaginationAdapter.Paginate(doc);

            // p1 은 페이지 0, p2 는 페이지 1 이상에 있어야 한다.
            int p1Page = result.Pages.SelectMany(p => p.BodyBlocks)
                .First(b => ReferenceEquals(b.Source, p1)).PageIndex;
            int p2Page = result.Pages.SelectMany(p => p.BodyBlocks)
                .First(b => ReferenceEquals(b.Source, p2)).PageIndex;

            Assert.Equal(0, p1Page);
            Assert.True(p2Page > p1Page,
                $"강제 페이지 나누기 단락은 직전 단락보다 뒤 페이지에 있어야 함: p1={p1Page}, p2={p2Page}");
            Assert.True(result.PageCount >= 2);
        });
    }

    [Fact]
    public void Paginate_ForcePageBreakBefore_OnFirstParagraph_StaysOnFirstPage()
    {
        RunOnSta(() =>
        {
            var doc     = new PolyDonkyument();
            var section = new Section();

            // 첫 블록에 ForcePageBreakBefore — 직전 블록이 없으므로 페이지 0 유지.
            var p = new Paragraph();
            p.Style.ForcePageBreakBefore = true;
            p.AddText("페이지 0 유지");
            section.Blocks.Add(p);

            doc.Sections.Add(section);

            var result = FlowDocumentPaginationAdapter.Paginate(doc);

            int pPage = result.Pages.SelectMany(b => b.BodyBlocks)
                .First(b => ReferenceEquals(b.Source, p)).PageIndex;
            Assert.Equal(0, pPage);
        });
    }

    [Fact]
    public void Paginate_MultipleForcePageBreaks_EachStartsNewPage()
    {
        RunOnSta(() =>
        {
            var doc     = new PolyDonkyument();
            var section = new Section();

            var p1 = new Paragraph(); p1.AddText("페이지 1 단락");
            section.Blocks.Add(p1);

            var p2 = new Paragraph();
            p2.Style.ForcePageBreakBefore = true;
            p2.AddText("페이지 2 단락");
            section.Blocks.Add(p2);

            var p3 = new Paragraph();
            p3.Style.ForcePageBreakBefore = true;
            p3.AddText("페이지 3 단락");
            section.Blocks.Add(p3);

            doc.Sections.Add(section);

            var result = FlowDocumentPaginationAdapter.Paginate(doc);

            int p1Page = result.Pages.SelectMany(p => p.BodyBlocks)
                .First(b => ReferenceEquals(b.Source, p1)).PageIndex;
            int p2Page = result.Pages.SelectMany(p => p.BodyBlocks)
                .First(b => ReferenceEquals(b.Source, p2)).PageIndex;
            int p3Page = result.Pages.SelectMany(p => p.BodyBlocks)
                .First(b => ReferenceEquals(b.Source, p3)).PageIndex;

            Assert.True(p2Page > p1Page);
            Assert.True(p3Page > p2Page);
            Assert.True(result.PageCount >= 3);
        });
    }

    [Fact]
    public void Paginate_ForcePageBreakBefore_SubsequentParagraphsAlsoOnNewPage()
    {
        // 페이지 나누기 단락 *이후*의 단락들도 새 페이지에 배정되어야 한다.
        // 버그 재현: 모든 콘텐츠가 한 페이지 높이 안에 들어갈 때, ForcePageBreakBefore
        // 단락만 페이지 1로 가고 그 뒤의 단락들이 페이지 0에 남는 문제.
        RunOnSta(() =>
        {
            var doc     = new PolyDonkyument();
            var section = new Section();

            // 페이지 나누기 앞 콘텐츠
            var before1 = new Paragraph(); before1.AddText("나누기 앞 단락 1");
            var before2 = new Paragraph(); before2.AddText("독도는 우리땅");
            section.Blocks.Add(before1);
            section.Blocks.Add(before2);

            // 페이지 나누기 단락 (Ctrl+Enter 로 삽입되는 빈 단락 — ForcePageBreakBefore=true)
            var pageBreakPara = new Paragraph();
            pageBreakPara.Style.ForcePageBreakBefore = true;
            section.Blocks.Add(pageBreakPara);

            // 페이지 나누기 뒤 콘텐츠 — 이것들도 새 페이지에 있어야 한다
            var after1 = new Paragraph(); after1.AddText("나누기 뒤 단락 1");
            var after2 = new Paragraph(); after2.AddText("나누기 뒤 단락 2");
            section.Blocks.Add(after1);
            section.Blocks.Add(after2);

            doc.Sections.Add(section);

            var result = FlowDocumentPaginationAdapter.Paginate(doc);

            var allBlocks = result.Pages.SelectMany(p => p.BodyBlocks).ToList();
            int beforePage = allBlocks.First(b => ReferenceEquals(b.Source, before2)).PageIndex;
            int breakPage  = allBlocks.First(b => ReferenceEquals(b.Source, pageBreakPara)).PageIndex;
            int after1Page = allBlocks.First(b => ReferenceEquals(b.Source, after1)).PageIndex;
            int after2Page = allBlocks.First(b => ReferenceEquals(b.Source, after2)).PageIndex;

            Assert.True(breakPage > beforePage,
                $"페이지 나누기 단락({breakPage})은 앞 단락({beforePage})보다 뒤 페이지여야 함");
            Assert.True(after1Page >= breakPage,
                $"나누기 뒤 단락1({after1Page})은 페이지 나누기({breakPage}) 이상 페이지여야 함");
            Assert.True(after2Page >= breakPage,
                $"나누기 뒤 단락2({after2Page})은 페이지 나누기({breakPage}) 이상 페이지여야 함");
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

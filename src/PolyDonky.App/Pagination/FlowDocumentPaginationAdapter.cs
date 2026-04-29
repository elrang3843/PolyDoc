using System.Windows;
using System.Windows.Controls;
using PolyDonky.App.Services;
using PolyDonky.Core;
using WpfDocs = System.Windows.Documents;

namespace PolyDonky.App.Pagination;

/// <summary>
/// WPF <see cref="WpfDocs.DynamicDocumentPaginator"/> 를 레이아웃 백엔드로 사용하는 페이지 분할 어댑터.
///
/// <para>
/// <b>STA 스레드에서 호출해야 한다</b> — <see cref="WpfDocs.FlowDocument"/> 는 <c>DependencyObject</c>.
/// </para>
///
/// <para>
/// 제약 (Phase 1):
/// <list type="bullet">
///   <item>본문 블록→페이지 매핑은 오프스크린 RichTextBox 의 Y 좌표 기반 근사치다.
///         페이지 경계를 가로지르는 블록은 부정확할 수 있다 (Phase 3 에서 정밀화).</item>
///   <item><see cref="BlockOnPage.BodyLocalRect"/> 는 Phase 3 이전 <c>Rect.Empty</c>.</item>
/// </list>
/// </para>
/// </summary>
public static class FlowDocumentPaginationAdapter
{
    /// <summary>
    /// 문서를 페이지 단위로 분할해 <see cref="PaginatedDocument"/> 로 반환한다.
    /// </summary>
    /// <param name="document">분할할 문서.</param>
    /// <param name="pageSettings">
    /// 페이지 설정 재정의. <c>null</c> 이면 문서 첫 섹션의 <see cref="PageSettings"/> 사용.
    /// </param>
    public static PaginatedDocument Paginate(
        PolyDonkyument document,
        PageSettings?  pageSettings = null)
    {
        ArgumentNullException.ThrowIfNull(document);

        var page = pageSettings
            ?? document.Sections.FirstOrDefault()?.Page
            ?? new PageSettings();
        var geo = new PageGeometry(page);

        // 1. FlowDocument 빌드 (PageHeight·PagePadding 으로 paginator 페이지 구분 설정)
        var fd = FlowDocumentBuilder.Build(document);
        fd.PageWidth   = geo.PageWidthDip;
        fd.PageHeight  = geo.PageHeightDip;
        fd.PagePadding = new Thickness(
            geo.PadLeftDip, geo.PadTopDip, geo.PadRightDip, geo.PadBottomDip);

        // sentinel 블록(PageBreakPadder 삽입 잔존물)이 있으면 제거
        PageBreakPadder.RemoveAll(fd.Blocks);

        // 2. DocumentPaginator 로 정확한 페이지 수 산출
        int pageCount = ComputePageCountSync(fd, geo);

        // 3. 오프스크린 RichTextBox 에서 본문 블록 Y 좌표 측정 → 페이지 배정
        var bodyAssignments = MapBodyBlocksToPages(fd, geo, pageCount);

        // 4. 오버레이 블록(글상자·이미지·도형·오버레이 표) 수집 — AnchorPageIndex 기준
        var overlayAssignments = CollectOverlayBlocks(document);

        // 오버레이의 최대 페이지 인덱스로 pageCount 보정
        if (overlayAssignments.Count > 0)
        {
            int maxOverlayPage = overlayAssignments.Max(o => o.pageIdx) + 1;
            pageCount = Math.Max(pageCount, maxOverlayPage);
        }
        pageCount = Math.Max(1, pageCount);

        // 5. PaginatedPage 조립
        var pages = BuildPages(pageCount, bodyAssignments, overlayAssignments);

        return new PaginatedDocument
        {
            Source       = document,
            PageSettings = page,
            Pages        = pages,
        };
    }

    // ── 페이지 수 계산 ────────────────────────────────────────────────────────

    private static int ComputePageCountSync(WpfDocs.FlowDocument fd, PageGeometry geo)
    {
        try
        {
            var paginator = (WpfDocs.DynamicDocumentPaginator)
                ((WpfDocs.IDocumentPaginatorSource)fd).DocumentPaginator;
            paginator.PageSize = new Size(geo.PageWidthDip, geo.PageHeightDip);

            // GetPage(n) 을 순차 호출해 IsPageCountValid 가 true 가 될 때까지 강제 레이아웃.
            // DocumentPaginator.ComputePageCount() 도 내부적으로 동일한 작업을 수행한다.
            int n = 0;
            while (!paginator.IsPageCountValid && n <= 1_000)
            {
                paginator.GetPage(n);
                n++;
            }

            return paginator.IsPageCountValid
                ? Math.Max(1, paginator.PageCount)
                : Math.Max(1, n);
        }
        catch
        {
            return 1;
        }
    }

    // ── 본문 블록 → 페이지 매핑 ──────────────────────────────────────────────

    private static List<(int pageIdx, Block coreBlock)> MapBodyBlocksToPages(
        WpfDocs.FlowDocument fd, PageGeometry geo, int pageCount)
    {
        var result = new List<(int, Block)>();

        // 오프스크린 RichTextBox 에서 Measure/Arrange → TextPointer.GetCharacterRect 활성화
        var rtb = new RichTextBox
        {
            Document          = fd,
            Padding           = new Thickness(0),
            BorderThickness   = new Thickness(0),
            VerticalAlignment = VerticalAlignment.Top,
        };
        rtb.Measure(new Size(geo.PageWidthDip, double.PositiveInfinity));
        rtb.Arrange(new Rect(rtb.DesiredSize));

        // 스크롤 연속 보기에서 "페이지당 본문 높이" = pageHeight - padTop - padBottom
        double bodyH = geo.PageHeightDip - geo.PadTopDip - geo.PadBottomDip;
        if (bodyH <= 0) bodyH = geo.PageHeightDip;

        foreach (var wpfBlock in FlattenBlocks(fd.Blocks))
        {
            if (wpfBlock.Tag is not Block coreBlock) continue;
            if (IsOverlayMode(coreBlock)) continue;

            double y       = TryGetTopY(wpfBlock);
            int    pageIdx = double.IsNaN(y)
                ? 0
                : Math.Clamp((int)(y / bodyH), 0, pageCount - 1);

            result.Add((pageIdx, coreBlock));
        }

        // RichTextBox 분리 (FlowDocument 재사용을 위해)
        rtb.Document = new WpfDocs.FlowDocument();
        return result;
    }

    /// <summary>
    /// fd.Blocks 를 재귀적으로 열거한다.
    /// <see cref="WpfDocs.List"/> 안의 <see cref="WpfDocs.ListItem"/> → Block 도 포함하므로
    /// 목록 단락이 누락되지 않는다.
    /// </summary>
    private static IEnumerable<WpfDocs.Block> FlattenBlocks(WpfDocs.BlockCollection blocks)
    {
        foreach (var b in blocks)
        {
            yield return b;
            if (b is WpfDocs.List list)
            {
                foreach (var li in list.ListItems)
                    foreach (var nested in FlattenBlocks(li.Blocks))
                        yield return nested;
            }
        }
    }

    private static double TryGetTopY(WpfDocs.Block block)
    {
        try
        {
            var r = block.ContentStart.GetCharacterRect(WpfDocs.LogicalDirection.Forward);
            if (r == Rect.Empty || double.IsInfinity(r.Y) || double.IsNaN(r.Y)) return double.NaN;
            return r.Y;
        }
        catch
        {
            return double.NaN;
        }
    }

    // ── 오버레이 블록 수집 ────────────────────────────────────────────────────

    private static List<(int pageIdx, Block coreBlock, double xMm, double yMm)>
        CollectOverlayBlocks(PolyDonkyument document)
    {
        var result = new List<(int, Block, double, double)>();

        foreach (var section in document.Sections)
        {
            foreach (var block in section.Blocks)
            {
                if (!IsOverlayMode(block)) continue;
                if (block is not IOverlayAnchored anchored) continue;

                result.Add((
                    Math.Max(0, anchored.AnchorPageIndex),
                    block,
                    anchored.OverlayXMm,
                    anchored.OverlayYMm));
            }
        }

        return result;
    }

    /// <summary>블록이 오버레이(본문 흐름 밖) 모드인지 반환한다.</summary>
    public static bool IsOverlayMode(Block block) => block switch
    {
        TextBoxObject                                                    => true,
        ImageBlock  img => img.WrapMode is ImageWrapMode.InFrontOfText
                                        or ImageWrapMode.BehindText,
        ShapeObject shp => shp.WrapMode is ImageWrapMode.InFrontOfText
                                        or ImageWrapMode.BehindText,
        Table       tbl => tbl.WrapMode != TableWrapMode.Block,
        _                                                                => false,
    };

    // ── PaginatedPage 조립 ───────────────────────────────────────────────────

    private static IReadOnlyList<PaginatedPage> BuildPages(
        int                                               pageCount,
        List<(int pageIdx, Block coreBlock)>              bodyAssignments,
        List<(int pageIdx, Block coreBlock, double xMm, double yMm)> overlayAssignments)
    {
        var pages = new PaginatedPage[pageCount];
        for (int i = 0; i < pageCount; i++)
        {
            pages[i] = new PaginatedPage
            {
                PageIndex = i,
                BodyBlocks = bodyAssignments
                    .Where(b => b.pageIdx == i)
                    .Select(b => new BlockOnPage { Source = b.coreBlock, PageIndex = i })
                    .ToArray(),
                OverlayBlocks = overlayAssignments
                    .Where(o => o.pageIdx == i)
                    .Select(o => new OverlayOnPage
                    {
                        Source          = o.coreBlock,
                        AnchorPageIndex = o.pageIdx,
                        XMm             = o.xMm,
                        YMm             = o.yMm,
                    })
                    .ToArray(),
            };
        }
        return pages;
    }
}

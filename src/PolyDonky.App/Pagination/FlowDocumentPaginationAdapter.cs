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
/// 근사치 주의:
/// <list type="bullet">
///   <item>본문 블록→페이지 매핑은 오프스크린 RichTextBox 의 Y 좌표 기반 근사치다.
///         페이지 경계를 가로지르는 블록은 페이지 시작 지점에서 잘릴 수 있다.</item>
///   <item><see cref="BlockOnPage.BodyLocalRect"/> 는 연속 스크롤 공간 기준이며
///         FlowDocument 가 실제로 페이지 나눔을 수행한 좌표와 미묘하게 다를 수 있다.
///         Phase 3c(per-page 편집기 재설계) 에서 정밀화 예정.</item>
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

        // 3. 오프스크린 RichTextBox 에서 본문 블록 Y 좌표 측정 → (페이지, 단) 배정.
        // 측정 폭은 단 폭(geo.ColWidthDip) — 단일 단이면 본문 폭과 동일.
        // 전체 용지 폭으로 두면 줄바꿈이 적게 일어나 Y 좌표가 실제 단 RTB 와 달라진다.
        fd.PageWidth   = geo.ColWidthDip;
        fd.PagePadding = new Thickness(0);
        var bodyAssignments = MapBodyBlocksToPages(fd, geo, pageCount);

        // 본문 블록의 실제 배치 결과로 pageCount 보정.
        // DocumentPaginator(풀 페이지+여백) 와 오프스크린 RTB(단 폭·단 슬롯 높이) 측정이
        // 미세하게 어긋나 블록이 DocumentPaginator 산출 pageCount 를 넘는 페이지로 떨어질 수 있다.
        if (bodyAssignments.Count > 0)
        {
            int maxBodyPage = bodyAssignments.Max(b => b.pageIdx) + 1;
            pageCount = Math.Max(pageCount, maxBodyPage);
        }

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

    // ── 본문 블록 → 페이지·단 매핑 ──────────────────────────────────────────

    private static List<(int pageIdx, int colIdx, Block coreBlock, Rect bodyLocalRect)>
        MapBodyBlocksToPages(WpfDocs.FlowDocument fd, PageGeometry geo, int pageCount)
    {
        var result = new List<(int, int, Block, Rect)>();

        // 연속 스크롤 공간에서 "단 슬롯 높이" = pageHeight - padTop - padBottom
        double bodyH = geo.PageHeightDip - geo.PadTopDip - geo.PadBottomDip;
        if (bodyH <= 0) bodyH = geo.PageHeightDip;

        int    colCount  = geo.ColumnCount;
        double colWidth  = geo.ColWidthDip;

        // 오프스크린 RichTextBox — 측정 폭은 단 폭(colWidth).
        // 다단일 때 전체 본문 폭으로 측정하면 줄바꿈이 적게 일어나
        // Y 좌표가 실제 단 레이아웃과 달라진다.
        var rtb = new RichTextBox
        {
            Document          = fd,
            Padding           = new Thickness(0),
            BorderThickness   = new Thickness(0),
            VerticalAlignment = VerticalAlignment.Top,
        };
        rtb.Measure(new Size(colWidth, double.PositiveInfinity));
        rtb.Arrange(new Rect(rtb.DesiredSize));
        // UpdateLayout() 을 명시적으로 호출해 FlowDocument 내부 레이아웃을 동기적으로 완료.
        // OnLoaded 등 WPF 첫 렌더링 이전 시점에는 Measure/Arrange 만으로 텍스트 레이아웃이
        // 확정되지 않아 GetCharacterRect 가 Y=0 을 반환할 수 있다.
        rtb.UpdateLayout();

        // 단 슬롯별 누적 채움 높이. 슬롯 경계 이동 후 다른 블록과 합산 시
        // bodyH 를 넘기는 경우를 감지해 다음 슬롯으로 밀어낸다.
        var slotFill = new System.Collections.Generic.Dictionary<int, double>();

        foreach (var wpfBlock in FlattenBlocks(fd.Blocks))
        {
            if (wpfBlock.Tag is not Block coreBlock) continue;
            if (IsOverlayMode(coreBlock)) continue;

            double topY    = TryGetTopY(wpfBlock);
            double bottomY = TryGetBottomY(wpfBlock);
            int    pageIdx;
            int    colIdx;

            if (double.IsNaN(topY))
            {
                pageIdx = 0;
                colIdx  = 0;
            }
            else
            {
                double blockH  = (!double.IsNaN(bottomY) && bottomY > topY) ? (bottomY - topY) : 0.0;
                // 연속 스크롤 공간에서 "단 슬롯" 인덱스 (단 슬롯 = 단 × 페이지).
                // 상한 클램프 없음 — 호출자(Paginate) 가 max pageIdx 로 pageCount 를 보정한다.
                int slotTop = Math.Max(0, (int)(topY / bodyH));

                // 블록이 단 슬롯 경계를 넘고 한 슬롯에 들어갈 만큼 작으면 다음 슬롯으로 이동.
                if (!double.IsNaN(bottomY)
                    && bottomY > (slotTop + 1) * bodyH
                    && blockH < bodyH)
                {
                    slotTop += 1;
                }

                // 슬롯 누적 채움이 이 블록을 수용하기에 부족하면 다음 슬롯으로 밀어낸다.
                // bodyH 이상인 블록은 분할 불가이므로 채움 추적 대상에서 제외.
                if (blockH > 0 && blockH < bodyH)
                {
                    while (slotFill.GetValueOrDefault(slotTop, 0.0) + blockH > bodyH)
                        slotTop += 1;
                }

                // 슬롯 채움 갱신 (bodyH 캡 — 단 높이를 초과하는 블록은 슬롯 전체를 포화로 표시)
                if (blockH > 0)
                    slotFill[slotTop] = Math.Min(bodyH,
                        slotFill.GetValueOrDefault(slotTop, 0.0) + Math.Min(blockH, bodyH));

                pageIdx = slotTop / colCount;
                colIdx  = slotTop % colCount;
            }

            var bodyLocalRect = TryGetColumnLocalRect(wpfBlock, pageIdx, colIdx, bodyH, colWidth, colCount);
            result.Add((pageIdx, colIdx, coreBlock, bodyLocalRect));
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

    private static double TryGetBottomY(WpfDocs.Block block)
    {
        try
        {
            var r = block.ContentEnd.GetCharacterRect(WpfDocs.LogicalDirection.Backward);
            if (r == Rect.Empty || double.IsNaN(r.Bottom) || double.IsInfinity(r.Bottom))
                return double.NaN;
            return r.Bottom;
        }
        catch
        {
            return double.NaN;
        }
    }

    /// <summary>
    /// 연속 스크롤 공간에서 블록의 경계 상자를 측정해 해당 단 슬롯 기준 Rect 로 변환한다.
    /// 단 슬롯 경계를 넘어서는 블록은 슬롯 높이로 잘린다.
    /// 측정 실패 시 <see cref="Rect.Empty"/> 반환.
    /// </summary>
    private static Rect TryGetColumnLocalRect(
        WpfDocs.Block block, int pageIdx, int colIdx, double bodyH, double colWidth, int colCount)
    {
        try
        {
            var topRect = block.ContentStart.GetCharacterRect(WpfDocs.LogicalDirection.Forward);
            var botRect = block.ContentEnd.GetCharacterRect(WpfDocs.LogicalDirection.Backward);

            if (topRect == Rect.Empty || double.IsNaN(topRect.Y) || double.IsInfinity(topRect.Y))
                return Rect.Empty;

            double globalTop    = topRect.Y;
            double globalBottom = (botRect != Rect.Empty
                                   && !double.IsNaN(botRect.Bottom)
                                   && !double.IsInfinity(botRect.Bottom))
                ? botRect.Bottom
                : globalTop;

            // 단 슬롯 인덱스 (다단에서 페이지·단을 통합 순서로 열거)
            int    slotIdx     = pageIdx * colCount + colIdx;
            double slotOriginY = slotIdx * bodyH;
            double localTop    = globalTop - slotOriginY;
            // 슬롯 경계를 넘어서는 부분은 잘림
            double localBottom = Math.Min(globalBottom - slotOriginY, bodyH);
            double height      = Math.Max(0, localBottom - localTop);

            return new Rect(0, localTop, colWidth, height);
        }
        catch
        {
            return Rect.Empty;
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
        int                                                                    pageCount,
        List<(int pageIdx, int colIdx, Block coreBlock, Rect bodyLocalRect)>  bodyAssignments,
        List<(int pageIdx, Block coreBlock, double xMm, double yMm)>          overlayAssignments)
    {
        var pages = new PaginatedPage[pageCount];
        for (int i = 0; i < pageCount; i++)
        {
            pages[i] = new PaginatedPage
            {
                PageIndex = i,
                BodyBlocks = bodyAssignments
                    .Where(b => b.pageIdx == i)
                    .Select(b => new BlockOnPage
                    {
                        Source        = b.coreBlock,
                        PageIndex     = i,
                        ColumnIndex   = b.colIdx,
                        BodyLocalRect = b.bodyLocalRect,
                    })
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

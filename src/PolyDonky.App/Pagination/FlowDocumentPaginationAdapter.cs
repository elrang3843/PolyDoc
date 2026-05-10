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
    /// 정밀 페이지 매핑(블록별 GetCharacterRect)을 시도할 최대 블록 수.
    /// 이 값을 초과하면 <see cref="MapBodyBlocksToPages"/> 가 fast-path 로 빠져
    /// 전체 FlowDocument 측정·블록별 좌표 조회를 건너뛴다(매우 큰 HTML 등에서 분 단위 멈춤 방지).
    /// fast-path 결과는 모든 블록을 page 0 에 일괄 배정 — 페이지 구분이 정확하지 않지만
    /// 문서를 즉시 표시할 수 있어 "최소한 본다" 는 UX 가 보장된다.
    /// </summary>
    public const int MaxBlocksForPreciseMapping = 2_500;

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

        // 2. DocumentPaginator 로 정확한 페이지 수 산출.
        // ComputePageCountSync 내부에서 모든 페이지를 GetPage(n) 으로 강제 레이아웃하므로
        // 반환 시점에 paginator 는 완전한 페이지 배치 정보를 갖고 있다.
        var paginator = (WpfDocs.DynamicDocumentPaginator)
            ((WpfDocs.IDocumentPaginatorSource)fd).DocumentPaginator;
        int pageCount = ComputePageCountSync(fd, geo, paginator);

        // 표(Wpf.Table) 전용 조각(fragment) 맵: fd 치수를 colWidth/no-padding 으로 바꾸기 *전*에
        // 완전한 용지 기하로 레이아웃된 paginator 에게 각 표·행이 속한 페이지를 직접 질의한다.
        // 페이지를 넘는 표는 TableRowSplitter 로 행 기준 조각으로 분할한다.
        // Y 좌표 측정 방식은 오프스크린 RTB 에서 표 셀 내부 rect 를 신뢰할 수 없어 여러 번
        // 실패했으므로 paginator 의 확정 값으로 대체한다.
        var tableFragmentMap =
            new System.Collections.Generic.Dictionary<
                WpfDocs.Table,
                System.Collections.Generic.List<(Core.Table coreFragment, int slotIdx)>>();
        foreach (var b in FlattenBlocks(fd.Blocks))
        {
            if (b is not WpfDocs.Table wpfTbl) continue;
            if (b.Tag is not Core.Table coreTbl) continue;
            if (tableFragmentMap.ContainsKey(wpfTbl)) continue;

            var rowGroups = TableRowSplitter.GetRowGroups(wpfTbl, coreTbl, paginator);
            var fragments = TableRowSplitter.BuildFragments(coreTbl, rowGroups);

            tableFragmentMap[wpfTbl] = fragments
                .Select(f => (f.fragment, f.pageIdx * geo.ColumnCount))
                .ToList();
        }

        // 3. 오프스크린 RichTextBox 에서 본문 블록 Y 좌표 측정 → (페이지, 단) 배정.
        // 측정 폭은 단 폭(geo.ColWidthDip) — 단일 단이면 본문 폭과 동일.
        // 전체 용지 폭으로 두면 줄바꿈이 적게 일어나 Y 좌표가 실제 단 RTB 와 달라진다.
        fd.PageWidth   = geo.ColWidthDip;
        fd.PagePadding = new Thickness(0);
        var (bodyAssignments, slotFill, blockMeasurements) = MapBodyBlocksToPages(fd, geo, pageCount, tableFragmentMap);

        // 본문 블록의 실제 배치 결과로 pageCount 보정.
        // DocumentPaginator(풀 페이지+여백) 와 오프스크린 RTB(단 폭·단 슬롯 높이) 측정이
        // 미세하게 어긋나 블록이 DocumentPaginator 산출 pageCount 를 넘는 페이지로 떨어질 수 있다.
        if (bodyAssignments.Count > 0)
        {
            int maxBodyPage = bodyAssignments.Max(b => b.pageIdx) + 1;
            pageCount = Math.Max(pageCount, maxBodyPage);
        }

        // 4a. CSS flex 도형 spacer 기반 오버레이 좌표 해결.
        // HtmlReader 가 생성한 AnchorPageIndex=-2 ShapeObject 에 spacer 의 페이지·Y 를 반영한다.
        ResolveFlexShapeOverlays(document, bodyAssignments, geo);

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
            Source                  = document,
            PageSettings            = page,
            Pages                   = pages,
            SlotMeasuredFillDip     = slotFill,
            DebugBlockMeasurements  = blockMeasurements,
        };
    }

    // ── 페이지 수 계산 ────────────────────────────────────────────────────────

    private static int ComputePageCountSync(
        WpfDocs.FlowDocument fd, PageGeometry geo,
        WpfDocs.DynamicDocumentPaginator? paginator = null)
    {
        try
        {
            paginator ??= (WpfDocs.DynamicDocumentPaginator)
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

    private static (List<(int pageIdx, int colIdx, Block coreBlock, Rect bodyLocalRect)> assignments,
                    System.Collections.Generic.Dictionary<int, double> slotFillOut,
                    List<BlockMeasurementEntry> measurements)
        MapBodyBlocksToPages(
            WpfDocs.FlowDocument fd, PageGeometry geo, int pageCount,
            System.Collections.Generic.Dictionary<
                WpfDocs.Table,
                System.Collections.Generic.List<(Core.Table coreFragment, int slotIdx)>>? tableFragmentMap = null)
    {
        var result       = new List<(int, int, Block, Rect)>();
        var measurements = new List<BlockMeasurementEntry>();

        // 연속 스크롤 공간에서 "단 슬롯 높이" = pageHeight - padTop - padBottom
        double bodyH = geo.PageHeightDip - geo.PadTopDip - geo.PadBottomDip;
        if (bodyH <= 0) bodyH = geo.PageHeightDip;

        int    colCount  = geo.ColumnCount;
        double colWidth  = geo.ColWidthDip;

        // ── Fast path — 매우 큰 문서는 정밀 매핑을 건너뛴다 ─────────────────
        // FlattenBlocks 를 한 번 평탄화해 카운트만 본 뒤, 임계 초과면 모든 블록을 page 0 에 배정.
        // WPF FlowDocument 의 Measure/UpdateLayout 은 블록 수가 많을수록 슈퍼리니어로 느려지며,
        // 위키 페이지 등 큰 HTML 은 여기서 분 단위로 UI 가 멈춘다. fast-path 는 정확한 페이지
        // 경계를 잃지만 — 페이지 수는 ComputePageCountSync 결과를 그대로 사용 — 문서를 즉시
        // 표시할 수 있게 한다.
        var flat = FlattenBlocks(fd.Blocks).ToList();
        if (flat.Count > MaxBlocksForPreciseMapping)
        {
            foreach (var wpfBlock in flat)
            {
                if (wpfBlock.Tag is not Block coreBlock) continue;
                if (IsOverlayMode(coreBlock)) continue;
                result.Add((0, 0, coreBlock, Rect.Empty));
            }
            return (result, new System.Collections.Generic.Dictionary<int, double>(), measurements);
        }

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

        // mm→DIP 변환과 WPF 서브픽셀 반올림으로 인한 경계 오차 흡수용 허용치 (DIP).
        // 블록의 bottomY 가 슬롯 경계를 이 값 이하로 넘어서면 경계 위반으로 보지 않는다.
        // PerPageEditorHost.ClipRenderingTolerance 와 동일한 값을 유지한다.
        const double BoundaryTol = 2.0;

        // 페이지 하단에 남겨두는 최소 여유 공간 (DIP).
        // 오프스크린 RTB 측정값과 실제 렌더 높이 사이의 누적 오차(Padding.Bottom 일부 미반영,
        // 마진 붕괴 근사 등)를 흡수하기 위한 safety margin.
        // 이 값보다 작은 공간이 남으면 다음 블록을 다음 슬롯으로 밀어낸다.
        const double FillSafetyMarginDip = 15.0;

        // 단 슬롯별 누적 채움 높이 (DIP). 의미: 슬롯에 배정된 블록들의 "커서" —
        // 이전 블록 채움 + 블록 간 gap(WPF 마진 붕괴 반영) + 이 블록 높이 를 순차 누적한다.
        // 슬롯 이동(페이지 경계 초과·강제 나누기)이 발생하면 gap=0 으로 리셋해 슬롯 커서가
        // 이전 슬롯의 연속으로 이어지지 않도록 한다.
        // 페이지 경계 결정(fillOverflow), 단락 분할, 디버그 오버레이 표시에 사용된다.
        var slotFill = new System.Collections.Generic.Dictionary<int, double>();

        // 직전 블록이 최종 배정된 슬롯 인덱스. ForcePageBreakBefore 단락을 다음 페이지로
        // 강제 이동시킬 때 기준 페이지를 결정하는 데 사용. -1 = 첫 블록.
        int prevSlot = -1;

        // 강제 페이지 나누기 이후 이어지는 모든 블록의 최소 슬롯 인덱스.
        // ForcePageBreakBefore 블록이 slot N 에 배정되면 minSlot=N 으로 올라가
        // 그 뒤 블록들이 자연 Y 가 slot N 이전이어도 slot N 이상으로 배정되도록 보장한다.
        // (모든 콘텐츠가 한 페이지에 들어갈 만큼 짧아도 페이지 나누기가 정확히 작동하게 함.)
        int minSlot = 0;

        // 직전 블록의 연속 레이아웃 bottomY. 블록 사이 간격(gap)을 계산하는 데 사용한다.
        // gap = topY[i] - prevContBottom[i-1] — 이 값은 WPF 마진 붕괴(margin collapsing)와
        // 행간 여백을 이미 반영한 실제 간격이다.
        double prevContBottom = double.NaN;

        for (int i = 0; i < flat.Count; i++)
        {
            var wpfBlock = flat[i];

            // 이미지/표 캡션 단락은 Core Block 이 없는 WPF 전용 블록 — Core 배정 대상은 아니지만
            // 실제 렌더 높이를 차지하므로 slotFill 과 prevContBottom 은 갱신해야 한다.
            bool isSatellitePara = wpfBlock.Tag is FlowDocumentBuilder.ImageCaptionTag
                                  || wpfBlock.Tag == FlowDocumentBuilder.TableCaptionTag;
            if (isSatellitePara)
            {
                double satTopY    = TryGetTopY(wpfBlock);
                double satBottomY = TryGetBottomY(wpfBlock, colWidth);
                double satBlockH  = (!double.IsNaN(satBottomY) && satBottomY > satTopY) ? satBottomY - satTopY : 0.0;
                if (satBlockH > 0)
                {
                    double satGap = (!double.IsNaN(prevContBottom) && !double.IsNaN(satTopY) && satTopY > prevContBottom)
                        ? satTopY - prevContBottom : 0.0;
                    int satSlot = Math.Max(prevSlot >= 0 ? prevSlot : 0,
                                          (int)(satTopY / bodyH));
                    double satPrevFill = slotFill.GetValueOrDefault(satSlot, 0.0);
                    slotFill[satSlot] = Math.Min(bodyH, satPrevFill + satGap + satBlockH);
                }
                if (!double.IsNaN(satBottomY)) prevContBottom = satBottomY;
                continue;
            }

            if (wpfBlock.Tag is not Block coreBlock) continue;
            if (IsOverlayMode(coreBlock)) continue;

            double topY    = TryGetTopY(wpfBlock);
            double bottomY = TryGetBottomY(wpfBlock, colWidth);

            // gap 은 아래 cascade Y 보정 후에 확정된다.
            double gap;

            // 표(Wpf.Table) 는 paginator 가 확정한 조각(fragment) 목록을 우선 사용한다.
            // 단일 페이지 표 → 조각 1개(원본), 여러 페이지 표 → 행 기준 분할 조각 N개.
            if (wpfBlock is WpfDocs.Table wTbl
                && tableFragmentMap is not null
                && tableFragmentMap.TryGetValue(wTbl, out var tblFragments)
                && tblFragments.Count > 0)
            {
                foreach (var (coreFragment, slotIdx) in tblFragments)
                {
                    int tblSlot = Math.Max(minSlot, slotIdx);
                    result.Add((tblSlot / colCount, tblSlot % colCount, coreFragment, Rect.Empty));
                    prevSlot = tblSlot;
                    minSlot  = Math.Max(minSlot, tblSlot);
                }
                if (!double.IsNaN(bottomY)) prevContBottom = bottomY;
                continue;
            }

            // Y 를 측정할 수 없으면 직전 블록과 같은 슬롯에 배정한다.
            // prevSlot 이 유효하면 그것을 하한으로 사용해 minSlot 이 강제하는 최솟값도 존중한다.
            // (이전에는 무조건 minSlot=0 에 배정해 이미지/도형 직후 h2/h3 가 1페이지로 떨어지는 버그 있었음.)
            if (double.IsNaN(topY))
            {
                int nanSlot = prevSlot >= 0 ? Math.Max(minSlot, prevSlot) : minSlot;
                result.Add((nanSlot / colCount, nanSlot % colCount, coreBlock, Rect.Empty));
                measurements.Add(new BlockMeasurementEntry
                {
                    SlotIdx = nanSlot,
                    Label   = MakeMeasurementLabel(coreBlock, wpfBlock) + "(noY)",
                    TopY    = double.IsNaN(prevContBottom) ? double.NaN : prevContBottom,
                    BottomY = double.NaN,
                    BlockH  = 0,
                    Gap     = 0,
                });
                prevSlot = nanSlot;
                continue;
            }

            double blockH  = (!double.IsNaN(bottomY) && bottomY > topY) ? (bottomY - topY) : 0.0;

            // ── Cascade Y 보정 ────────────────────────────────────────────────────────
            // BUC(BlockUIContainer) 는 off-screen RTB 에서 layout height ≈ 0 이므로
            // BUC 이후 블록의 topY 가 BUC 시작 Y 부근으로 붕괴(collapse)된다.
            // topY < prevContBottom 이면 cascade 오염으로 판단하고 prevContBottom 을
            // 유효 시작 Y 로 사용한다. 이 보정은 slotTop 과 gap 모두에 적용되어
            // slotFill 과소평가(블록이 0 높이로 쌓이는 현상)도 함께 해소한다.
            // — prevSlot 을 하한으로 쓰는 이전 방식은 slotFill 오버플로 후 prevSlot
            //   이 N+1 이 되면 자연 슬롯 N 블록까지 전부 N+1 로 밀어 버리는 부작용이 있었다.
            double effectiveTopY = (!double.IsNaN(prevContBottom) && topY < prevContBottom)
                ? prevContBottom
                : topY;

            gap = (!double.IsNaN(prevContBottom) && effectiveTopY > prevContBottom)
                ? effectiveTopY - prevContBottom
                : 0.0;

            // 연속 스크롤 공간에서 "단 슬롯" 인덱스 (단 슬롯 = 단 × 페이지).
            // minSlot 을 하한으로 적용해 강제 페이지 나누기 이후 블록이 이전 페이지로
            // 돌아가지 않도록 한다. 상한 클램프 없음 — 호출자(Paginate) 가 보정.
            int slotTop = Math.Max(minSlot, (int)(effectiveTopY / bodyH));

            // ── 강제 페이지 나누기 처리 ─────────────────────────────────────────────
            // FlowDocumentBuilder 가 ForcePageBreakBefore=true 를 WPF 의 BreakPageBefore 로
            // 변환하지만, 본문 블록 Y 측정은 무한 높이(rtb.Measure(Size(_, +∞))) 에서 이루어져
            // paginator 의 페이지 나눔이 Y 좌표에 반영되지 않는다. 따라서 Y 기반 슬롯 매핑만으로는
            // 같은 슬롯(=같은 페이지) 에 묶일 수 있다 — 이를 직전 블록 페이지의 다음 페이지 첫 단으로
            // 끌어올려 보정하고, minSlot 을 갱신해 후속 블록도 같은 페이지 이상에 배정한다.
            // 첫 블록(prevSlot=-1) 에서는 적용하지 않는다(0 페이지 유지).
            bool isPageBreak = coreBlock is Paragraph fpara && fpara.Style.ForcePageBreakBefore;
            if (isPageBreak && prevSlot >= 0)
            {
                int forcedSlot = ((prevSlot / colCount) + 1) * colCount;
                if (slotTop < forcedSlot) slotTop = forcedSlot;
                minSlot = slotTop;
                gap = 0.0; // 페이지 나누기 후에는 앞 간격 리셋
            }

            // ── 줄 단위 분할 ──────────────────────────────────────────────────────────
            // 목록 마커가 없는 일반 단락이고, 한 슬롯에 들어갈 수 있는 높이일 때만 시도.
            // blockH >= bodyH 인 초장문 단락은 분할을 생략(구조 복잡도 대비 효용 낮음).
            // 강제 페이지 나누기 단락은 새 페이지 시작에 통째로 배치 — 줄 분할 생략.
            if (!isPageBreak
                && coreBlock is Paragraph corePara
                && corePara.Style.ListMarker == null
                && blockH > 0 && blockH < bodyH
                && wpfBlock is WpfDocs.Paragraph wpfPara)
            {
                double slotBoundaryY = (slotTop + 1) * bodyH;
                bool   crossesBoundary = !double.IsNaN(bottomY) && bottomY > slotBoundaryY + BoundaryTol;
                // 슬롯 커서(이전 블록들이 채운 양) + 현재 블록까지의 간격 + 이 단락 높이가 슬롯을 넘으면 오버플로.
                // cursor = slotFill[slotTop] 은 이전에 이 슬롯에 배정된 블록들의 누적 채움이다.
                double cursor4Para     = slotFill.GetValueOrDefault(slotTop, 0.0);
                bool   fillOverflow    = cursor4Para + gap + blockH > bodyH + BoundaryTol;

                if (crossesBoundary || fillOverflow)
                {
                    // 슬롯 경계에서 분할 — 이를 넘는 콘텐츠는 다음 슬롯의 RTB(높이=bodyH) 에 안 들어간다.
                    // 이전에는 fillOverflow 경우 sum 기반 splitY 를 썼지만 새 max-bottomY 의미론에서는
                    // 슬롯 끝(slotBoundaryY) 이 항상 올바른 분할점.
                    double splitY = slotBoundaryY;

                    int splitCharOffset = FindSplitCharOffset(wpfPara, splitY);
                    int totalChars      = corePara.Runs.Sum(r => r.Text.Length);

                    if (splitCharOffset > 0 && splitCharOffset < totalChars)
                    {
                        var (frag1, frag2) = SplitCoreParagraph(corePara, splitCharOffset);

                        // 첫 조각 → 현재 슬롯. 슬롯을 끝까지 채운 것으로 표시.
                        slotFill[slotTop] = bodyH;
                        var rect1 = TryGetColumnLocalRect(
                            wpfBlock, slotTop / colCount, slotTop % colCount, bodyH, colWidth, colCount);
                        result.Add((slotTop / colCount, slotTop % colCount, frag1, rect1));

                        // 이어지는 조각 → 다음 슬롯 (다음 슬롯에 다른 콘텐츠가 있으면 더 밀어냄)
                        int    nextSlot = slotTop + 1;
                        double frag2H   = Math.Max(0.0, blockH - (splitY - topY));
                        if (frag2H > 0 && frag2H < bodyH)
                        {
                            while (slotFill.GetValueOrDefault(nextSlot, 0.0) + frag2H > bodyH + BoundaryTol)
                                nextSlot++;
                        }
                        if (frag2H > 0)
                            // 커서 누적: nextSlot 에 이미 채워진 양 + frag2H
                            slotFill[nextSlot] = Math.Min(bodyH,
                                slotFill.GetValueOrDefault(nextSlot, 0.0) + frag2H);
                        result.Add((nextSlot / colCount, nextSlot % colCount, frag2, Rect.Empty));

                        prevSlot = nextSlot;
                        if (!double.IsNaN(bottomY)) prevContBottom = bottomY;
                        continue; // 아래 단일 블록 처리 생략
                    }
                }
            }

            // ── 단일 블록 배정 ────────────────────────────────────────────────────────
            // 블록이 단 슬롯 경계를 넘고 한 슬롯에 들어갈 만큼 작으면 다음 슬롯으로 이동.
            // BoundaryTol 이하의 초과는 mm→DIP 변환 오차로 보고 현재 슬롯에 유지한다.
            // 강제 페이지 나누기 단락은 위에서 이미 슬롯을 끌어올렸으므로 이 보정을 건너뛴다
            // (자연 Y 기준 boundary 검사가 강제 슬롯과 어긋나면 잘못된 +1 이 일어날 수 있다).
            //
            // 표(Wpf.Table) 는 한 슬롯 높이를 넘어도 다음 슬롯으로 이동을 허용한다 —
            // 표 분할이 구현되어 있지 않은 현재 상태에서는 "현재 페이지 끝에서 잘려 일부만 보이는"
            // 것보다 "다음 페이지 시작에 통째로 놓이고 끝이 잘리는" 편이 본문 행이 더 많이 보여 낫다.
            bool isTable = wpfBlock is WpfDocs.Table;
            if (!isPageBreak
                && !double.IsNaN(bottomY)
                && bottomY > (slotTop + 1) * bodyH + BoundaryTol
                && (blockH < bodyH || isTable))
            {
                slotTop += 1;
                gap = 0.0; // 슬롯 이동 시 간격 리셋 — 새 슬롯은 이전 슬롯의 연속이 아니다
            }

            // 슬롯 커서(이 슬롯에 이미 배정된 누적 채움) + gap + blockH 가 슬롯 높이를 넘으면
            // 다음 슬롯으로 밀어낸다. bodyH 이상인 블록은 분할 불가이므로 채움 추적 대상에서 제외.
            if (blockH > 0 && blockH < bodyH)
            {
                while (slotFill.GetValueOrDefault(slotTop, 0.0) + gap + blockH > bodyH - FillSafetyMarginDip)
                {
                    slotTop += 1;
                    gap = 0.0; // 슬롯 이동 시 간격 리셋
                }
            }

            // 슬롯 채움 갱신: "이전 커서 + 간격 + 이 블록 높이" 의 순수 누적.
            // 블록이 같은 슬롯에 자연 배치되면 gap 이 실제 WPF 마진·행간을 반영하고,
            // 슬롯 이동 후라면 gap=0 이므로 blockH 만 더해진다.
            // 이전 bottomY 기준 localBottomY 방식은 블록이 이동됐을 때 실제 높이를 과소평가했다.
            if (blockH > 0)
            {
                double prevFill = slotFill.GetValueOrDefault(slotTop, 0.0);
                slotFill[slotTop] = Math.Min(bodyH, prevFill + gap + blockH);
            }

            var bodyLocalRect = TryGetColumnLocalRect(
                wpfBlock, slotTop / colCount, slotTop % colCount, bodyH, colWidth, colCount);
            result.Add((slotTop / colCount, slotTop % colCount, coreBlock, bodyLocalRect));

            // 진단: 블록 측정값 기록 (디버그 오버레이에 표시됨)
            measurements.Add(new BlockMeasurementEntry
            {
                SlotIdx = slotTop,
                Label   = MakeMeasurementLabel(coreBlock, wpfBlock),
                TopY    = topY,
                BottomY = bottomY,
                BlockH  = blockH,
                Gap     = gap,
            });

            prevSlot = slotTop;
            if (!double.IsNaN(bottomY)) prevContBottom = bottomY;
        }

        // RichTextBox 분리 (FlowDocument 재사용을 위해)
        rtb.Document = new WpfDocs.FlowDocument();

        // ── 고아 제목 방지 ──────────────────────────────────────────────────────
        // BUC(이미지·도형·flex 표) 높이 cascade 오류로 인해 제목 단락이
        // 직후 내용보다 앞 페이지에 배정되는 경우를 역방향 스캔으로 교정한다.
        // 역방향이면 연속된 제목 체인(h1→h2→h3→content)도 한 번의 패스로 처리된다.
        for (int oi = result.Count - 2; oi >= 0; oi--)
        {
            (int pageIdx, int colIdx, Block coreBlock, Rect bodyLocalRect) curr = result[oi];
            (int pageIdx, int colIdx, Block coreBlock, Rect bodyLocalRect) next = result[oi + 1];
            if (curr.coreBlock is Paragraph hp
                && hp.Style.Outline != OutlineLevel.Body
                && next.pageIdx > curr.pageIdx)
            {
                result[oi] = (next.pageIdx, next.colIdx, curr.coreBlock, curr.bodyLocalRect);
            }
        }

        return (result, slotFill, measurements);
    }

    /// <summary>
    /// 진단 레이블 생성 — 블록 타입과 내용 일부를 최대 32자로 압축.
    /// </summary>
    private static string MakeMeasurementLabel(Block coreBlock, WpfDocs.Block wpfBlock)
    {
        string typeName;
        string extra = "";
        if (coreBlock is Paragraph corePara3 && corePara3.Style is not null)
        {
            if (corePara3.Style.CodeLanguage is not null)
            {
                string lang = corePara3.Style.CodeLanguage is "" ? "?" : corePara3.Style.CodeLanguage;
                typeName = $"Code({lang})";
            }
            else
            {
                typeName = "Para";
            }
            var text = string.Concat(corePara3.Runs.Select(r => r.Text));
            extra = text.Length > 20 ? ":" + text[..20] : ":" + text;
        }
        else
        {
            typeName = coreBlock switch
            {
                ImageBlock   => "Img",
                ShapeObject  => "Shape",
                Table        => "Table",
                ContainerBlock     => "Container",
                ThematicBreakBlock => "HR",
                TocBlock     => "TOC",
                _            => coreBlock.GetType().Name,
            };
        }

        // WPF 블록 타입 힌트 (BlockUIContainer + 자식 FE 타입 포함)
        string wpfHint;
        if (wpfBlock is WpfDocs.BlockUIContainer buc)
        {
            string childName = buc.Child is null ? "null" : buc.Child.GetType().Name;
            wpfHint = $"(BUC:{childName})";
        }
        else
        {
            wpfHint = wpfBlock is WpfDocs.Paragraph ? "" : $"({wpfBlock.GetType().Name})";
        }

        return $"{typeName}{wpfHint}{extra}";
    }

    // ── 줄 단위 분할 헬퍼 ────────────────────────────────────────────────────────

    /// <summary>
    /// WPF Paragraph 내에서 Y 좌표 splitY 직전까지의 텍스트 문자 수를 반환한다.
    /// 이진 탐색으로 O(log n) 심볼 위치를 찾은 뒤 TextRange 로 문자 수를 센다.
    /// </summary>
    /// <summary>외부(글상자 다단 등)에서 동일 분할 로직 재사용을 위한 internal alias.</summary>
    internal static int FindSplitCharOffsetPublic(WpfDocs.Paragraph wpfPara, double splitY)
        => FindSplitCharOffset(wpfPara, splitY);

    /// <summary>외부에서 동일 분할 로직 재사용을 위한 internal alias.</summary>
    internal static (Paragraph first, Paragraph second) SplitCoreParagraphPublic(
        Paragraph para, int charOffset)
        => SplitCoreParagraph(para, charOffset);

    private static int FindSplitCharOffset(WpfDocs.Paragraph wpfPara, double splitY)
    {
        var start = wpfPara.ContentStart;
        var end   = wpfPara.ContentEnd;
        int total = start.GetOffsetToPosition(end);
        if (total <= 0) return 0;

        // 이진 탐색: 줄의 하단(rect.Y + rect.Height) 이 splitY 이내인 마지막 심볼 위치 찾기.
        // 줄 시작점(rect.Y) 만 비교하면 splitY 직전에서 시작하는 줄이 splitY 를 넘어 끝나도
        // 포함되어 frag1 의 마지막 줄이 슬라이스 RTB(높이=bodyH) 에서 클리핑된다.
        // 줄 하단까지 비교하면 그런 줄은 frag1 에 포함되지 않아 frag2(다음 단/페이지) 의 첫 줄로 간다.
        // 2 DIP 허용치는 mm→DIP 반올림 오차 흡수용 (PerPageEditorHost.ClipRenderingTolerance 와 대응).
        const double LineFitTol = 2.0;
        int lo = 0, hi = total;
        while (lo < hi - 1)
        {
            int mid = (lo + hi) / 2;
            var tp = start.GetPositionAtOffset(mid);
            if (tp == null) { hi = mid; continue; }

            var rect = tp.GetCharacterRect(WpfDocs.LogicalDirection.Forward);
            if (rect == Rect.Empty || double.IsNaN(rect.Y) || double.IsInfinity(rect.Y))
            {
                lo = mid; // 측정 불가 위치는 분할선 이전으로 처리
                continue;
            }
            double lineBottom = rect.Y + (double.IsNaN(rect.Height) ? 0 : rect.Height);
            if (lineBottom <= splitY + LineFitTol) lo = mid;
            else                                   hi = mid;
        }

        var splitPtr = start.GetPositionAtOffset(lo);
        if (splitPtr == null || splitPtr.CompareTo(start) <= 0) return 0;

        // 심볼 위치 → 텍스트 문자 수 (단락 내부 범위이므로 \n 없음)
        try
        {
            return new WpfDocs.TextRange(start, splitPtr).Text.Length;
        }
        catch
        {
            return 0;
        }
    }

    /// <summary>
    /// 단락을 charOffset 문자 위치에서 두 조각으로 나눈다.
    /// 각 조각의 Id 에 "§f0" / "§f1" 접미사를 붙여 ParseAllPageEditors 에서 재결합할 수 있게 한다.
    /// </summary>
    private static (Paragraph first, Paragraph second) SplitCoreParagraph(
        Paragraph para, int charOffset)
    {
        // 원본 Id 가 없으면 임시 그룹 Id 생성 (§g 접두사로 식별)
        string groupId = para.Id is { } origId
            ? origId
            : "§g" + System.Guid.NewGuid().ToString("N")[..8];

        var first  = CloneParaForSplit(para, groupId + "§f0");
        var second = CloneParaForSplit(para, groupId + "§f1");
        // 이어지는 조각의 문단 앞뒤 간격은 제거 (시각적 연속성)
        second.Style.SpaceBeforePt = 0;
        first.Style.SpaceAfterPt  = 0;
        // 두 번째 조각은 목록 마커를 반복 표시하지 않는다 — 한 단락이 두 페이지에 걸쳐 있을 때 불릿/번호는
        // 첫 줄에서만 보여야 한다. QuoteLevel·CodeLanguage 등 단락 전체에 적용되는 속성은 둘 다 보존.
        second.Style.ListMarker = null;

        int remaining = charOffset;
        bool inSecond = false;

        foreach (var run in para.Runs)
        {
            if (inSecond)
            {
                second.Runs.Add(run.Clone());
                continue;
            }

            string text = run.Text;

            if (remaining <= 0)
            {
                second.Runs.Add(run.Clone());
                inSecond = true;
            }
            else if (remaining >= text.Length)
            {
                first.Runs.Add(run.Clone());
                remaining -= text.Length;
            }
            else // 0 < remaining < text.Length: run 을 둘로 쪼갬
            {
                // Run.Clone() 으로 모든 필드 복사 (Url 등 누락 방지) 후 텍스트만 분할.
                var firstRun  = run.Clone();
                firstRun.Text = text[..remaining];
                first.Runs.Add(firstRun);

                var secondRun  = run.Clone();
                secondRun.Text = text[remaining..];
                second.Runs.Add(secondRun);

                remaining = 0;
                inSecond  = true;
            }
        }

        return (first, second);
    }

    private static Paragraph CloneParaForSplit(Paragraph p, string? id) => new()
    {
        Id      = id,
        Status  = p.Status,
        StyleId = p.StyleId,
        // ParagraphStyle.Clone() 으로 모든 필드 복사 — 이전 자체 구현은 ListMarker / QuoteLevel /
        // CodeLanguage 누락. 호출측에서 second.Style.ListMarker = null 로 마커 제거.
        Style   = p.Style.Clone(),
    };

    // Run/RunStyle 복제는 Core 의 정식 Clone() 메서드 사용 — 이전에 이 파일에 있던 자체 구현은
    // Url 등 새로 추가된 필드를 빠뜨려 페이지네이션 분할 시 하이퍼링크가 사라지는 버그가 있었다.

    /// <summary>
    /// fd.Blocks 를 재귀적으로 열거한다.
    /// <see cref="WpfDocs.List"/> 안의 <see cref="WpfDocs.ListItem"/> → Block 도 포함하므로
    /// 목록 단락이 누락되지 않는다.
    /// </summary>
    private static IEnumerable<WpfDocs.Block> FlattenBlocks(WpfDocs.BlockCollection blocks)
    {
        foreach (var b in blocks)
        {
            // Section(=ContainerBlock 렌더 결과) 은 자체를 yield 하지 않고 자식만 재귀.
            // 그래야 Section 안의 단락/표/이미지가 Section 의 단일 Y 좌표 한 점이 아닌
            // 각자의 Y 로 페이지에 배정돼 컨테이너가 길어도 페이지를 넘어 분산된다.
            // 페이지 단위 box 시각화는 같은 페이지에 모든 자식이 들어갔을 때만 보장 — Section 이
            // 페이지 경계를 넘으면 box framing 이 끊긴다 (현재 단계의 trade-off).
            if (b is WpfDocs.Section sect)
            {
                foreach (var nested in FlattenBlocks(sect.Blocks))
                    yield return nested;
                continue;
            }

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

    private static double TryGetBottomY(WpfDocs.Block block, double colWidth)
    {
        try
        {
            // BlockUIContainer: ContentEnd.GetCharacterRect 는 캐럿 높이만 반환하므로
            // 자식 UIElement 의 높이를 직접 측정한다.
            if (block is WpfDocs.BlockUIContainer buc && buc.Child is FrameworkElement fe)
            {
                double topY = TryGetTopY(block);
                if (!double.IsNaN(topY))
                {
                    double h = ResolveFEHeight(fe, colWidth);
                    if (h > 0)
                        // topY = ContentStart 기준 → Margin.Top 이미 포함 → Margin.Bottom 만 추가.
                        return topY + h + block.Margin.Bottom;
                }
                // 높이를 못 구하면 NaN 반환 — caret 높이(≈20 DIP) 폴백을 피해
                // 잘못된 작은 blockH 로 페이지 분할이 어긋나는 일을 막는다.
                return double.NaN;
            }

            // Wpf.Table 은 ContentEnd.GetCharacterRect 가 표 직후 캐럿 위치(≈ 표의 top) 만
            // 반환한다. 셀 내부 단락도 오프스크린 RTB 에서 NaN 을 반환하는 경우가 많아
            // 여기서는 NaN 을 그대로 반환 — 호출측(MapBodyBlocksToPages) 에서 다음 블록
            // topY 대리값으로 보정한다.
            if (block is WpfDocs.Table)
                return double.NaN;

            var r = block.ContentEnd.GetCharacterRect(WpfDocs.LogicalDirection.Backward);
            if (r == Rect.Empty || double.IsNaN(r.Bottom) || double.IsInfinity(r.Bottom))
            {
                // Paragraph 에 InlineUIContainer 가 있으면(코드 블록 줄 번호, 수식 FormulaControl 등)
                // ContentEnd.GetCharacterRect 가 Rect.Empty 를 반환하는 경우가 있다.
                // InlineUIContainer.Child.ActualHeight 와 LineBreak 수를 이용해 높이를 추정한다.
                if (block is WpfDocs.Paragraph para)
                    return TryEstimateParaBottomViaInlines(para);
                return double.NaN;
            }
            return r.Bottom;
        }
        catch
        {
            return double.NaN;
        }
    }

    /// <summary>
    /// <c>FrameworkElement</c> 의 높이를 반환한다. <c>ActualHeight</c> 우선, 0 이면
    /// 명시적 <c>Height</c> → <c>Measure(colWidth, ∞)</c> → <c>DesiredSize.Height</c> 순으로 폴백.
    /// 오프스크린 RTB 에서 <c>BlockUIContainer</c> 자식 UIElement 가 레이아웃 미완료로
    /// <c>ActualHeight=0</c> 을 반환하는 경우에도 올바른 높이를 얻기 위한 보완 측정이다.
    /// </summary>
    private static double ResolveFEHeight(FrameworkElement fe, double colWidth)
    {
        if (!double.IsNaN(fe.ActualHeight) && fe.ActualHeight > 0)
            return fe.ActualHeight;

        // 명시적 Height 프로퍼티 (Image.Height, Border.Height 등)
        if (!double.IsNaN(fe.Height) && fe.Height > 0)
            return fe.Height;

        // 강제 측정 + 배치 — 오프스크린 RTB UpdateLayout 이후에도 UIElement 가 layout-pass 를
        // 받지 못한 경우(FlowDocument 내 BlockUIContainer 에서 흔함)에 대한 보완.
        // Measure 만으로 DesiredSize.Height=0 이 나오는 케이스(빈 데이터 이미지의 StackPanel,
        // SVG Viewbox 래퍼 등)에 대비해 Arrange 까지 강제하면 ActualHeight 가 채워진다.
        // 폭은 단 폭을 상한으로, 명시적 Width 가 있으면 그 값이 우선된다.
        try
        {
            double availW = (!double.IsNaN(fe.Width) && fe.Width > 0) ? fe.Width : colWidth;
            fe.Measure(new Size(availW, double.PositiveInfinity));
            if (fe.DesiredSize.Height > 0)
                return fe.DesiredSize.Height;

            // Arrange 강제 — 일부 컨테이너(StackPanel/Grid)에서 자식 ActualHeight 가
            // Arrange 후에야 채워지는 경우가 있다.
            fe.Arrange(new Rect(0, 0, fe.DesiredSize.Width, fe.DesiredSize.Height));
            if (fe.ActualHeight > 0) return fe.ActualHeight;
        }
        catch { /* 측정 실패 시 0 반환 */ }

        return 0;
    }

    /// <summary>
    /// <c>ContentEnd.GetCharacterRect</c> 가 실패한 <c>Paragraph</c> 의 bottomY 를
    /// <c>InlineUIContainer.Child.ActualHeight</c> 와 <c>LineBreak</c> 수에서 추정한다.
    /// <para>
    /// 코드 블록 줄 번호(<c>BuildCodeBlockWithLineNumbers</c>) 나 수식(<c>FormulaControl</c>) 처럼
    /// 오프스크린 RTB 에서 UIElement 를 포함한 단락이 <c>Rect.Empty</c> 를 반환할 때 사용한다.
    /// </para>
    /// </summary>
    private static double TryEstimateParaBottomViaInlines(WpfDocs.Paragraph para)
    {
        double topY = TryGetTopY(para);
        if (double.IsNaN(topY)) return double.NaN;

        int    lineCount  = 1;
        double lineHeight = 0;

        foreach (var inline in para.Inlines)
        {
            if (inline is WpfDocs.LineBreak)
            {
                lineCount++;
            }
            else if (inline is WpfDocs.InlineUIContainer { Child: FrameworkElement fe })
            {
                double h = fe.ActualHeight;
                if (h > 0 && !double.IsNaN(h) && h > lineHeight) lineHeight = h;
            }
        }

        if (lineHeight <= 0)
        {
            // InlineUIContainer 없음(일반 단락) — FontSize × 기본 행간으로 추정
            lineHeight = (para.FontSize > 0 ? para.FontSize : FlowDocumentBuilder.PtToDip(11)) * 1.2;
        }

        return topY
               + lineCount * lineHeight
               + para.Padding.Top + para.Padding.Bottom
               + para.Margin.Top  + para.Margin.Bottom;
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

            // BlockUIContainer 는 ContentEnd.GetCharacterRect 가 캐럿 높이만 반환할 수 있어
            // 내부 UIElement 높이로 globalBottom 을 계산한다.
            // globalTop = ContentStart.Y 는 이미 Margin.Top 이후 위치 → Margin.Bottom 만 추가.
            double globalBottom = double.NaN;
            if (block is WpfDocs.BlockUIContainer buc2 && buc2.Child is FrameworkElement fe2)
            {
                double feH2 = ResolveFEHeight(fe2, colWidth);
                if (feH2 > 0) globalBottom = globalTop + feH2 + block.Margin.Bottom;
            }
            if (double.IsNaN(globalBottom))
            {
                if (botRect != Rect.Empty
                    && !double.IsNaN(botRect.Bottom)
                    && !double.IsInfinity(botRect.Bottom))
                {
                    globalBottom = botRect.Bottom;
                }
                else
                {
                    // Paragraph 에 InlineUIContainer(코드 블록 줄 번호, 수식 등)가 있으면
                    // ContentEnd.GetCharacterRect 가 Rect.Empty 를 반환할 수 있다.
                    double estimated = block is WpfDocs.Paragraph para2
                        ? TryEstimateParaBottomViaInlines(para2)
                        : double.NaN;
                    globalBottom = !double.IsNaN(estimated) ? estimated : globalTop;
                }
            }

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

    // ── flex 도형 오버레이 좌표 해결 ─────────────────────────────────────────

    /// <summary>
    /// HtmlReader 가 생성한 <c>AnchorPageIndex = -2</c> 도형(CSS flex 순수 도형)의
    /// 페이지 인덱스와 절대 좌표를 spacer 단락의 body 배치 결과로 확정한다.
    /// </summary>
    private static void ResolveFlexShapeOverlays(
        PolyDonkyument                                                              document,
        List<(int pageIdx, int colIdx, Block coreBlock, Rect bodyLocalRect)>  bodyAssignments,
        PageGeometry                                                            geo)
    {
        if (bodyAssignments.Count == 0) return;

        // Core Block → 배치 결과 룩업.
        var lookup = new System.Collections.Generic.Dictionary<Block,
            (int pageIdx, Rect bodyLocalRect)>(bodyAssignments.Count);
        foreach (var a in bodyAssignments)
            lookup.TryAdd(a.coreBlock, (a.pageIdx, a.bodyLocalRect));

        double marginLeftMm = FlowDocumentBuilder.DipToMm(geo.PadLeftDip);
        double marginTopMm  = FlowDocumentBuilder.DipToMm(geo.PadTopDip);

        int    currentSpacerPage  = 0;
        double currentSpacerYMm   = 0; // spacer 콘텐츠 상단의 body 내 Y (mm)

        foreach (var section in document.Sections)
        {
            foreach (var block in section.Blocks)
            {
                if (block is Paragraph p && p.StyleId == "pd-flex-shape-spacer")
                {
                    if (lookup.TryGetValue(p, out var info)
                        && info.bodyLocalRect != Rect.Empty)
                    {
                        currentSpacerPage = info.pageIdx;
                        currentSpacerYMm  = FlowDocumentBuilder.DipToMm(info.bodyLocalRect.Top);
                    }
                    // bodyLocalRect 가 Rect.Empty 이면 이전 값 유지 (fallback).
                }
                else if (block is ShapeObject shape && shape.AnchorPageIndex == -2)
                {
                    shape.AnchorPageIndex = currentSpacerPage;
                    // OverlayXMm/YMm 은 콘텐츠 영역 기준 상대값 → 페이지 절대 좌표로 변환.
                    shape.OverlayXMm += marginLeftMm;
                    shape.OverlayYMm += marginTopMm + currentSpacerYMm;
                }
            }
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

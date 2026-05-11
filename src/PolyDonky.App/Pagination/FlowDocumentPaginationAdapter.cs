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
                // fast-path: ContainerBlock Section 은 자식 코어 블록을 직접 추가.
                if (wpfBlock is WpfDocs.Section fastSect && fastSect.Tag is ContainerBlock)
                {
                    AssignSectionCoreBlocks(fastSect, 0, colCount, result);
                    continue;
                }
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
        // 오프스크린 RTB 측정값과 실제 렌더 높이 사이의 누적 오차(서브픽셀 반올림, 마진 붕괴 근사,
        // 폰트 어센더·디센더 미세 오차 등)를 흡수하기 위한 safety margin.
        // BUC(이미지·SVG·flex) 높이 오차는 Core 모델 폴백(MinBucBlockH 분기)에서 직접 보정하므로
        // 여기서 흡수할 필요가 없어졌다 — 텍스트 측정 잔여 오차 수준인 15 DIP 로 유지.
        const double FillSafetyMarginDip = 15.0;

        // 단 슬롯별 누적 채움 높이 (DIP). 의미: 슬롯에 배정된 블록들의 "커서" —
        // 이전 블록 채움 + 블록 간 gap(WPF 마진 붕괴 반영) + 이 블록 높이 를 순차 누적한다.
        // 슬롯 이동(페이지 경계 초과·강제 나누기)이 발생하면 gap=0 으로 리셋해 슬롯 커서가
        // 이전 슬롯의 연속으로 이어지지 않도록 한다.
        // 페이지 경계 결정(fillOverflow), 단락 분할, 디버그 오버레이 표시에 사용된다.
        var slotFill = new System.Collections.Generic.Dictionary<int, double>();

        // result[i] 에 해당하는 slotFill 기여량 — (슬롯 인덱스, gap+blockH, blockH 단독).
        // 주 루프에서 일반 블록(line 525)의 기여분을 기록하고,
        // 이후 orphan heading scan 이 제목을 다음 페이지로 이동할 때
        // source/target 슬롯의 slotFill 을 사후 보정하는 데 사용한다.
        // (리스트 아이템·NaN 블록 등은 기여분이 없거나 개별 추적이 불필요해 등록하지 않는다.)
        var resultFillContribs = new System.Collections.Generic.Dictionary<int, (int slotIdx, double contribution, double blockHOnly)>();

        // result[i] → measurements[j] 인덱스 매핑.
        // orphan heading scan 이 result[i] 를 이동할 때 measurements[j].SlotIdx 도
        // 함께 갱신해 디버그 오버레이가 실제 배정 페이지를 올바르게 표시하도록 한다.
        // 리스트 아이템·분할 단락 등 measurements 항목이 없는 result 항목은 등록하지 않는다.
        var resultToMeasurement = new System.Collections.Generic.Dictionary<int, int>();

        // 슬롯별 최초 배정된 블록의 off-screen RTB topY.
        // minSlot 강제 배정으로 인해 자연 슬롯보다 높은 슬롯에 블록이 몰릴 때,
        // slotFill 대신 "prevContBottom - slotContentStartY[slot]" 로 실제 시각 채움을
        // 추정해 ContainerBlock overflow 검사에 사용한다.
        var slotContentStartY = new System.Collections.Generic.Dictionary<int, double>();

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

            // ── Wpf.List 전체 처리 ────────────────────────────────────────────────────────
            // 리스트 내 개별 단락의 Y 좌표가 off-screen RTB 에서 collapse 되어 slotFill 이
            // 심각하게 과소평가되는 문제를 방지한다. List 전체의 topY/bottomY 를 단번에 측정해
            // slotFill 에 반영하고, 모든 ListItem CoreBlock 을 동일 슬롯에 일괄 배정한다.
            // 제약: List 가 한 슬롯 높이를 초과할 경우 분할 없이 같은 페이지에 통째로 배정됨
            // (단일 페이지를 넘는 리스트가 없는 현재 문서에서는 허용 가능한 단순화).
            if (wpfBlock is WpfDocs.List listWpf && wpfBlock.Tag is not Block)
            {
                double listTopY    = TryGetTopY(listWpf);
                double listBottomY = TryGetBottomY(listWpf, colWidth);
                if (!double.IsNaN(listTopY))
                {
                    double effListTopY = (!double.IsNaN(prevContBottom) && listTopY < prevContBottom)
                        ? prevContBottom : listTopY;
                    double listGap = (!double.IsNaN(prevContBottom) && effListTopY > prevContBottom)
                        ? effListTopY - prevContBottom : 0.0;
                    double listH = (!double.IsNaN(listBottomY) && listBottomY > listTopY)
                        ? listBottomY - listTopY : 0.0;

                    int listSlot = Math.Max(minSlot, (int)(effListTopY / bodyH));

                    if (listH > 0 && listH < bodyH)
                    {
                        while (slotFill.GetValueOrDefault(listSlot, 0.0) + listGap + listH > bodyH - FillSafetyMarginDip)
                        {
                            listSlot++;
                            listGap = 0.0;
                        }
                    }
                    if (listH > 0)
                        slotFill[listSlot] = Math.Min(bodyH,
                            slotFill.GetValueOrDefault(listSlot, 0.0) + listGap + listH);

                    int listResultFirst = result.Count;
                    AssignListCoreBlocks(listWpf, listSlot, colCount, result);
                    int listResultLast = result.Count - 1;

                    // resultFillContribs 등록: cascade 가 리스트 아이템을 찾아 다음 슬롯으로
                    // 밀 수 있도록 각 아이템에 기여분을 기록한다.
                    // 첫 번째 아이템만 listGap+listH(전체 기여), 나머지는 0.
                    // cascade 루프가 마지막→첫 번째 순으로 이동하다가 첫 번째 아이템을 옮길 때
                    // slotFill 이 listH 만큼 실제로 감소하며, 리스트 전체가 한 단위로 다음 슬롯에 배정된다.
                    for (int ri = listResultFirst; ri <= listResultLast; ri++)
                    {
                        if (ri == listResultFirst)
                            resultFillContribs[ri] = (listSlot, listGap + listH, listH);
                        else
                            resultFillContribs[ri] = (listSlot, 0.0, 0.0);
                        // 리스트 아이템은 모두 같은 measurements 항목(아래 Add 전)을 가리킨다.
                        resultToMeasurement[ri] = measurements.Count;
                    }

                    measurements.Add(new BlockMeasurementEntry
                    {
                        SlotIdx = listSlot,
                        Label   = $"List({listWpf.ListItems.Count}items)",
                        TopY    = listTopY,
                        BottomY = listBottomY,
                        BlockH  = listH,
                        Gap     = listGap,
                    });

                    if (!double.IsNaN(listTopY) && !slotContentStartY.ContainsKey(listSlot))
                        slotContentStartY[listSlot] = listTopY;
                    prevSlot = listSlot;
                    minSlot  = Math.Max(minSlot, listSlot);
                    // cascade 발생 시 effListTopY 를 기준으로 보정 — 개별 블록 갱신과 동일 원칙.
                    if (!double.IsNaN(listBottomY)) prevContBottom = effListTopY + listH;
                    else if (!double.IsNaN(listTopY)) prevContBottom = effListTopY;
                }
                continue;
            }

            // ── Wpf.Section (ContainerBlock) 전체 처리 ──────────────────────────────────
            // ContainerBlock 전체의 topY/bottomY 를 단번에 측정해 단일 단위로 슬롯에 배정.
            // 단일 페이지 이내의 ContainerBlock 이 페이지 경계에 걸리면 통째로 다음 슬롯으로 이동.
            // Wpf.List 전용 핸들러와 동일한 패턴 — 자식 CoreBlock 들은 모두 같은 슬롯에 할당.
            if (wpfBlock is WpfDocs.Section sectWpf && sectWpf.Tag is ContainerBlock)
            {
                double sectTopY    = TryGetTopY(sectWpf);
                double sectBottomY = TryGetBottomY(sectWpf, colWidth);
                if (!double.IsNaN(sectTopY))
                {
                    double effSectTopY = (!double.IsNaN(prevContBottom) && sectTopY < prevContBottom)
                        ? prevContBottom : sectTopY;
                    double sectGap = (!double.IsNaN(prevContBottom) && effSectTopY > prevContBottom)
                        ? effSectTopY - prevContBottom : 0.0;
                    double sectH = (!double.IsNaN(sectBottomY) && sectBottomY > sectTopY)
                        ? sectBottomY - sectTopY : 0.0;

                    int sectSlot = Math.Max(minSlot, (int)(effSectTopY / bodyH));

                    if (sectH > 0 && sectH < bodyH)
                    {
                        // 1차: Y 좌표 기반 검사 — slotFill 과소평가(NaN 블록 등)와 무관하게
                        //       오프스크린 RTB 의 실제 레이아웃 좌표로 경계를 판단한다.
                        //       effSectTopY 는 cascade 보정을 이미 반영한 값이므로
                        //       pageRelY = (현재 슬롯 안에서의 상대 Y).
                        double pageRelY = effSectTopY - (double)sectSlot * bodyH;
                        bool   yOverflow = pageRelY + sectH > bodyH - FillSafetyMarginDip;

                        // 2차: 슬롯 실제 콘텐츠 시작 Y 기반 검사 — minSlot 강제 배정으로
                        //       블록이 자연 슬롯보다 높은 슬롯에 몰릴 때 slotFill 이 실제
                        //       시각 채움을 크게 과소평가하는 케이스를 잡는다.
                        //       prevContBottom - slotFirstContentTopY = 이 슬롯에서 지금까지
                        //       시각적으로 소비된 높이 추정치.
                        bool actualFillOverflow = false;
                        if (!double.IsNaN(prevContBottom) &&
                            slotContentStartY.TryGetValue(sectSlot, out double slotStartContentY))
                        {
                            double slotActualFill = prevContBottom - slotStartContentY;
                            actualFillOverflow = slotActualFill + sectGap + sectH > bodyH - FillSafetyMarginDip;
                        }

                        if (yOverflow || actualFillOverflow)
                        {
                            sectSlot++;
                            sectGap = 0.0;
                        }
                        else
                        {
                            // 3차: slotFill 기반 검사 — cascade 로 effSectTopY 가 높아진 경우 보완.
                            while (slotFill.GetValueOrDefault(sectSlot, 0.0) + sectGap + sectH > bodyH - FillSafetyMarginDip)
                            {
                                sectSlot++;
                                sectGap = 0.0;
                            }
                        }
                    }
                    if (!double.IsNaN(sectTopY) && !slotContentStartY.ContainsKey(sectSlot))
                        slotContentStartY[sectSlot] = sectTopY;
                    if (sectH > 0)
                        slotFill[sectSlot] = Math.Min(bodyH,
                            slotFill.GetValueOrDefault(sectSlot, 0.0) + sectGap + sectH);

                    int sectResultFirst = result.Count;
                    AssignSectionCoreBlocks(sectWpf, sectSlot, colCount, result);
                    int sectResultLast = result.Count - 1;

                    for (int ri = sectResultFirst; ri <= sectResultLast; ri++)
                    {
                        if (ri == sectResultFirst)
                            resultFillContribs[ri] = (sectSlot, sectGap + sectH, sectH);
                        else
                            resultFillContribs[ri] = (sectSlot, 0.0, 0.0);
                        resultToMeasurement[ri] = measurements.Count;
                    }

                    measurements.Add(new BlockMeasurementEntry
                    {
                        SlotIdx = sectSlot,
                        Label   = $"Container({sectWpf.Blocks.Count}blocks)",
                        TopY    = sectTopY,
                        BottomY = sectBottomY,
                        BlockH  = sectH,
                        Gap     = sectGap,
                    });

                    prevSlot = sectSlot;
                    minSlot  = Math.Max(minSlot, sectSlot);
                    if (!double.IsNaN(sectBottomY)) prevContBottom = effSectTopY + sectH;
                    else if (!double.IsNaN(sectTopY)) prevContBottom = effSectTopY;
                }
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
                // BUC(이미지) 등 topY 측정 불가 블록이 슬롯 최초 콘텐츠일 때,
                // slotContentStartY 가 누락되어 ContainerBlock actualFillOverflow 검사가
                // 발동하지 않는 문제를 방지한다. prevContBottom 을 슬롯 진입 위치 대리값으로 사용.
                if (!double.IsNaN(prevContBottom) && !slotContentStartY.ContainsKey(nanSlot))
                    slotContentStartY[nanSlot] = prevContBottom;
                prevSlot = nanSlot;
                continue;
            }

            double blockH  = (!double.IsNaN(bottomY) && bottomY > topY) ? (bottomY - topY) : 0.0;

            // ── BUC 높이 폴백 ───────────────────────────────────────────────────────
            // BlockUIContainer(이미지·SVG·flex 컨테이너) 는 off-screen RTB 에서
            // layout 이 완료되지 않아 height ≈ 0 으로 측정된다.
            // blockH ≈ 0 이면 slotFill 에 기여가 없어 해당 페이지의 채움이 실제보다
            // 훨씬 작게 계산되고, 이후 블록이 계속 같은 페이지로 밀려 누적 오버플로가 발생한다.
            // Core 모델에 치수가 있는 블록(ImageBlock·ShapeObject)은 HeightMm 을 폴백 높이로 사용한다.
            const double MinBucBlockH = 5.0; // DIP — 이 이하면 측정 실패로 판단
            if (blockH < MinBucBlockH && wpfBlock is WpfDocs.BlockUIContainer)
            {
                double fallbackH = coreBlock switch
                {
                    ImageBlock  img   when img.HeightMm > 0   => FlowDocumentBuilder.MmToDip(img.HeightMm),
                    ShapeObject shape when shape.HeightMm > 0 => FlowDocumentBuilder.MmToDip(shape.HeightMm),
                    _                                          => 0.0,
                };
                if (fallbackH > blockH)
                    blockH = fallbackH;
            }

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
                        minSlot  = Math.Max(minSlot, nextSlot); // 분할 후에도 문서 순서 보장
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
            // 다음 슬롯으로 밀어낸다.
            // bodyH 미만 블록: 들어갈 슬롯이 나올 때까지 반복 이동.
            // bodyH 이상 블록(분할 불가): 현재 슬롯에 이미 내용이 있을 때만 한 번 이동.
            //   — 같은 페이지에 작은 블록 + 큰 BUC(SVG/flex 등)가 함께 배정되면
            //     per-page RTB(bodyH+2) 를 초과해 시각적으로 잘리는 문제를 방지한다.
            if (blockH > 0)
            {
                if (blockH < bodyH)
                {
                    while (slotFill.GetValueOrDefault(slotTop, 0.0) + gap + blockH > bodyH - FillSafetyMarginDip)
                    {
                        slotTop += 1;
                        gap = 0.0;
                    }
                }
                else if (slotFill.GetValueOrDefault(slotTop, 0.0) + gap > FillSafetyMarginDip)
                {
                    // 이미 채워진 슬롯에 bodyH 이상 블록이 오면 다음 슬롯으로 이동.
                    slotTop += 1;
                    gap = 0.0;
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
            if (!double.IsNaN(topY) && !slotContentStartY.ContainsKey(slotTop))
                slotContentStartY[slotTop] = topY;
            resultFillContribs[result.Count - 1]  = (slotTop, gap + blockH, blockH);
            resultToMeasurement[result.Count - 1] = measurements.Count; // 아래 measurements.Add 직전

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
            minSlot  = Math.Max(minSlot, slotTop); // 문서 순서 보장: 이후 블록이 앞 슬롯으로 돌아가지 않도록
            // prevContBottom 갱신: effectiveTopY 를 기준으로 보정.
            // cascade 가 발생(effectiveTopY > topY)했을 때 원래 topY 를 쓰면 이후 블록의
            // cascade 체인이 끊어져 slotFill 이 과소평가된다. effectiveTopY 를 기준으로 해야
            // "List → BUC HR → 일반 단락" 같은 연속 cascade 시퀀스가 끊기지 않는다.
            if (!double.IsNaN(topY) && blockH > 0 && (double.IsNaN(bottomY) || bottomY - topY < blockH))
                prevContBottom = effectiveTopY + blockH;
            else if (!double.IsNaN(bottomY))
                prevContBottom = effectiveTopY + (bottomY - topY);
        }

        // RichTextBox 분리 (FlowDocument 재사용을 위해)
        rtb.Document = new WpfDocs.FlowDocument();

        // ── 고아 제목 방지 ──────────────────────────────────────────────────────
        // BUC(이미지·도형·flex 표) 높이 cascade 오류로 인해 제목 단락이
        // 직후 내용보다 앞 페이지에 배정되는 경우를 역방향 스캔으로 교정한다.
        // 역방향이면 연속된 제목 체인(h1→h2→h3→content)도 한 번의 패스로 처리된다.
        //
        // slotFill 사후 보정: 제목을 page N → N+1 으로 이동하면
        //   - page N   slotFill -= (gap + blockH)  ← 제목 기여분 환원
        //   - page N+1 slotFill += blockH          ← 제목이 실제 렌더할 높이
        // target 슬롯이 bodyH - FillSafetyMarginDip 를 넘으면 target 마지막 블록을
        // 다음 슬롯으로 cascade 해 overflow 를 방지한다.
        for (int oi = result.Count - 2; oi >= 0; oi--)
        {
            (int pageIdx, int colIdx, Block coreBlock, Rect bodyLocalRect) curr = result[oi];
            (int pageIdx, int colIdx, Block coreBlock, Rect bodyLocalRect) next = result[oi + 1];
            if (curr.coreBlock is Paragraph hp
                && hp.Style.Outline != OutlineLevel.Body
                && next.pageIdx > curr.pageIdx)
            {
                result[oi] = (next.pageIdx, next.colIdx, curr.coreBlock, curr.bodyLocalRect);

                if (resultFillContribs.TryGetValue(oi, out var hc))
                {
                    int oldSlot = curr.pageIdx * colCount + curr.colIdx;
                    int newSlot = next.pageIdx * colCount + next.colIdx;

                    // source 페이지: 제목 기여분(gap + blockH) 제거
                    slotFill[oldSlot] = Math.Max(0.0,
                        slotFill.GetValueOrDefault(oldSlot, 0.0) - hc.contribution);

                    // target 페이지: 제목 높이 + WPF 마진(SpaceBefore + SpaceAfter) 추가.
                    // blockH 는 TranslatePoint 기반 콘텐츠 높이만이며, WPF 렌더 시
                    // Paragraph.Margin.Top/Bottom 이 별도로 추가된다. 이 마진을 반영하지 않으면
                    // 여러 제목이 같은 페이지에 쌓일 때 slotFill 이 실제보다 과소 계산되어 오버플로가 된다.
                    double hSpaceBefore = FlowDocumentBuilder.PtToDip(
                        hp.Style.SpaceBeforePt > 0 ? hp.Style.SpaceBeforePt :
                        hp.Style.Outline switch
                        {
                            OutlineLevel.H1 => 12.0,
                            OutlineLevel.H2 => 10.0,
                            OutlineLevel.H3 =>  8.0,
                            OutlineLevel.H4 =>  6.0,
                            OutlineLevel.H5 =>  4.0,
                            OutlineLevel.H6 =>  4.0,
                            _               =>  0.0,
                        });
                    double hSpaceAfter = FlowDocumentBuilder.PtToDip(
                        hp.Style.SpaceAfterPt > 0 ? hp.Style.SpaceAfterPt :
                        hp.Style.Outline switch
                        {
                            OutlineLevel.H1 => 6.0,
                            OutlineLevel.H2 => 4.0,
                            OutlineLevel.H3 => 4.0,
                            OutlineLevel.H4 => 2.0,
                            OutlineLevel.H5 => 2.0,
                            OutlineLevel.H6 => 2.0,
                            _               => 0.0,
                        });
                    double headingMargin = hSpaceBefore + hSpaceAfter;
                    double headingTotalH = hc.blockHOnly + headingMargin;

                    slotFill[newSlot] = Math.Min(bodyH,
                        slotFill.GetValueOrDefault(newSlot, 0.0) + headingTotalH);
                    resultFillContribs[oi] = (newSlot, headingTotalH, headingTotalH);

                    // measurements 도 실제 배정 슬롯으로 갱신 — 디버그 오버레이가
                    // 이동된 제목을 올바른 페이지에 표시하도록 한다.
                    if (resultToMeasurement.TryGetValue(oi, out int mIdx))
                    {
                        var om = measurements[mIdx];
                        measurements[mIdx] = new BlockMeasurementEntry
                        {
                            SlotIdx = newSlot,
                            Label   = om.Label + "→",   // 이동됨을 라벨로 표시
                            TopY    = om.TopY,
                            BottomY = om.BottomY,
                            BlockH  = om.BlockH,
                            Gap     = 0,  // 새 페이지 시작 — gap 0 으로 보정
                        };
                    }

                    // target 슬롯 overflow → 마지막 블록을 다음 슬롯으로 반복 cascade.
                    // 단일 if 대신 while 루프로 여러 블록이 밀려야 할 때도 처리한다.
                    while (slotFill.GetValueOrDefault(newSlot, 0.0) > bodyH - FillSafetyMarginDip)
                    {
                        // target 슬롯의 마지막 비-제목 블록 (forward scan; heading=oi 는 제외)
                        int lastOnTarget = -1;
                        for (int j = oi + 1; j < result.Count; j++)
                        {
                            (int rPage, int rCol, Block _, Rect __) = result[j];
                            if (rPage == next.pageIdx && rCol == next.colIdx)
                                lastOnTarget = j;
                            else if (rPage > next.pageIdx)
                                break;
                        }
                        if (lastOnTarget <= oi
                            || !resultFillContribs.TryGetValue(lastOnTarget, out var lc))
                            break; // 더 밀 블록 없음

                        int cascadeSlot = newSlot + 1;
                        slotFill[newSlot] = Math.Max(0.0,
                            slotFill.GetValueOrDefault(newSlot, 0.0) - lc.contribution);
                        slotFill[cascadeSlot] = Math.Min(bodyH,
                            slotFill.GetValueOrDefault(cascadeSlot, 0.0) + lc.blockHOnly);
                        (int lrPage, int lrCol, Block lrCore, Rect lrRect) = result[lastOnTarget];
                        result[lastOnTarget] = (cascadeSlot / colCount, cascadeSlot % colCount,
                            lrCore, lrRect);
                        resultFillContribs[lastOnTarget] = (cascadeSlot, lc.blockHOnly, lc.blockHOnly);
                        // cascade 된 블록의 measurements 도 갱신
                        if (resultToMeasurement.TryGetValue(lastOnTarget, out int cmIdx))
                        {
                            var cm = measurements[cmIdx];
                            measurements[cmIdx] = new BlockMeasurementEntry
                            {
                                SlotIdx = cascadeSlot,
                                Label   = cm.Label + "↓",
                                TopY    = cm.TopY,
                                BottomY = cm.BottomY,
                                BlockH  = cm.BlockH,
                                Gap     = 0,
                            };
                        }
                    }
                }
            }
        }

        // ── Y-span 사후 보정 ────────────────────────────────────────────────────
        // off-screen RTB 에서 측정한 블록들의 Y 범위(maxBottomY - minTopY)가 bodyH 를 넘으면
        // per-page RTB 에서 마지막 블록이 잘린다. slotFill 이 gap 리셋·ContainerBlock 마진 등으로
        // 실제 높이를 과소평가하기 때문에 발생한다.
        // 슬롯별 Y 범위를 직접 계산해 bodyH 초과 슬롯에서 마지막 블록을 다음 슬롯으로 cascade 한다.
        // guard: ySpan > 2×bodyH 는 제목 이동 등으로 topY·bottomY 가 서로 다른 페이지에 걸쳐
        // 수집된 경우로, 이 보정 대상이 아니다.
        {
            var slotMinTopY    = new Dictionary<int, double>();
            var slotMaxBottomY = new Dictionary<int, double>();
            foreach (var m in measurements)
            {
                if (double.IsNaN(m.TopY) || double.IsNaN(m.BottomY)) continue;
                int s = m.SlotIdx;
                if (!slotMinTopY.ContainsKey(s) || m.TopY < slotMinTopY[s])
                    slotMinTopY[s] = m.TopY;
                if (!slotMaxBottomY.ContainsKey(s) || m.BottomY > slotMaxBottomY[s])
                    slotMaxBottomY[s] = m.BottomY;
            }
            foreach (int s in slotMinTopY.Keys.OrderBy(k => k))
            {
                double ySpan = slotMaxBottomY.GetValueOrDefault(s) - slotMinTopY[s];
                if (ySpan <= bodyH || ySpan > 2.0 * bodyH) continue;

                while (ySpan > bodyH)
                {
                    // 이 슬롯에 배정된 마지막 블록을 찾는다.
                    int lastOnSlot = -1;
                    for (int j = 0; j < result.Count; j++)
                    {
                        (int rPage, int rCol, Block _, Rect __) = result[j];
                        if (rPage * colCount + rCol == s) lastOnSlot = j;
                    }
                    if (lastOnSlot < 0 || !resultFillContribs.TryGetValue(lastOnSlot, out var lc))
                        break;

                    int nextS = s + 1;
                    slotFill[s]     = Math.Max(0.0, slotFill.GetValueOrDefault(s)     - lc.contribution);
                    slotFill[nextS] = Math.Min(bodyH, slotFill.GetValueOrDefault(nextS) + lc.blockHOnly);
                    (int pg, int col, Block cb, Rect r) = result[lastOnSlot];
                    result[lastOnSlot] = (nextS / colCount, nextS % colCount, cb, r);
                    resultFillContribs[lastOnSlot] = (nextS, lc.blockHOnly, lc.blockHOnly);

                    if (resultToMeasurement.TryGetValue(lastOnSlot, out int mIdx))
                    {
                        var m = measurements[mIdx];
                        measurements[mIdx] = new BlockMeasurementEntry
                        {
                            SlotIdx = nextS,
                            Label   = m.Label + "↓",
                            TopY    = m.TopY,
                            BottomY = m.BottomY,
                            BlockH  = m.BlockH,
                            Gap     = 0,
                        };
                        // 이 슬롯의 maxBottomY 를 재계산해 while 조건을 갱신한다.
                        slotMaxBottomY[s] = measurements
                            .Where(mm => mm.SlotIdx == s && !double.IsNaN(mm.BottomY))
                            .Select(mm => mm.BottomY)
                            .DefaultIfEmpty(slotMinTopY[s])
                            .Max();
                    }
                    ySpan = slotMaxBottomY.GetValueOrDefault(s) - slotMinTopY[s];
                }
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
            // ContainerBlock 으로 태깅된 Section 은 자체를 yield — 주 루프의
            // Section 전용 핸들러가 단일 단위로 측정/배정한다.
            // Tag 가 없는 일반 Section(스타일 래퍼 등)은 종전처럼 재귀.
            if (b is WpfDocs.Section sect)
            {
                if (sect.Tag is ContainerBlock)
                {
                    yield return sect;
                    continue;
                }
                foreach (var nested in FlattenBlocks(sect.Blocks))
                    yield return nested;
                continue;
            }

            yield return b;
            // Wpf.List 자식 단락은 MapBodyBlocksToPages 의 List 전용 핸들러에서 직접 처리.
            // 여기서 개별 yield 하면 Y collapse 로 slotFill 이 과소평가되므로 재귀하지 않는다.
        }
    }

    /// <summary>
    /// Wpf.List 의 모든 ListItem CoreBlock 을 지정 슬롯에 재귀적으로 배정한다.
    /// 중첩 List 와 Section 래퍼를 투명하게 처리한다.
    /// </summary>
    private static void AssignListCoreBlocks(
        WpfDocs.List list,
        int slot,
        int colCount,
        List<(int pageIdx, int colIdx, Block coreBlock, Rect bodyLocalRect)> result)
    {
        foreach (var li in list.ListItems)
        {
            foreach (var block in li.Blocks)
            {
                if (block is WpfDocs.List nestedList)
                {
                    AssignListCoreBlocks(nestedList, slot, colCount, result);
                }
                else if (block is WpfDocs.Section sect)
                {
                    foreach (var sectBlock in sect.Blocks)
                    {
                        if (sectBlock is WpfDocs.List nestedList2)
                            AssignListCoreBlocks(nestedList2, slot, colCount, result);
                        else if (sectBlock.Tag is Block cb && !IsOverlayMode(cb))
                            result.Add((slot / colCount, slot % colCount, cb, Rect.Empty));
                    }
                }
                else if (block.Tag is Block coreBlock && !IsOverlayMode(coreBlock))
                {
                    result.Add((slot / colCount, slot % colCount, coreBlock, Rect.Empty));
                }
            }
        }
    }

    /// <summary>
    /// Wpf.Section (ContainerBlock 렌더 결과) 의 모든 직계·재귀 CoreBlock 을 지정 슬롯에 배정한다.
    /// 중첩 Section 과 List 를 투명하게 처리한다.
    /// </summary>
    private static void AssignSectionCoreBlocks(
        WpfDocs.Section sect,
        int slot,
        int colCount,
        List<(int pageIdx, int colIdx, Block coreBlock, Rect bodyLocalRect)> result)
    {
        foreach (var child in sect.Blocks)
        {
            if (child is WpfDocs.Section nestedSect)
            {
                AssignSectionCoreBlocks(nestedSect, slot, colCount, result);
            }
            else if (child is WpfDocs.List nestedList)
            {
                AssignListCoreBlocks(nestedList, slot, colCount, result);
            }
            else if (child.Tag is Block cb && !IsOverlayMode(cb))
            {
                result.Add((slot / colCount, slot % colCount, cb, Rect.Empty));
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
                if (block is Paragraph p &&
                    (p.StyleId == "pd-flex-shape-spacer" || p.StyleId == "image-spacer"))
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
                else if (block is ImageBlock imgBlock && imgBlock.AnchorPageIndex == -2)
                {
                    imgBlock.AnchorPageIndex = currentSpacerPage;
                    imgBlock.OverlayXMm     += marginLeftMm;
                    imgBlock.OverlayYMm     += marginTopMm + currentSpacerYMm;
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

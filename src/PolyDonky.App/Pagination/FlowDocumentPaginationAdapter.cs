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
            return result;
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

        // 단 슬롯별 누적 채움 높이. 슬롯 경계 이동 후 다른 블록과 합산 시
        // bodyH 를 넘기는 경우를 감지해 다음 슬롯으로 밀어낸다.
        var slotFill = new System.Collections.Generic.Dictionary<int, double>();

        // 직전 블록이 최종 배정된 슬롯 인덱스. ForcePageBreakBefore 단락을 다음 페이지로
        // 강제 이동시킬 때 기준 페이지를 결정하는 데 사용. -1 = 첫 블록.
        int prevSlot = -1;

        // 강제 페이지 나누기 이후 이어지는 모든 블록의 최소 슬롯 인덱스.
        // ForcePageBreakBefore 블록이 slot N 에 배정되면 minSlot=N 으로 올라가
        // 그 뒤 블록들이 자연 Y 가 slot N 이전이어도 slot N 이상으로 배정되도록 보장한다.
        // (모든 콘텐츠가 한 페이지에 들어갈 만큼 짧아도 페이지 나누기가 정확히 작동하게 함.)
        int minSlot = 0;

        foreach (var wpfBlock in FlattenBlocks(fd.Blocks))
        {
            if (wpfBlock.Tag is not Block coreBlock) continue;
            if (IsOverlayMode(coreBlock)) continue;

            double topY    = TryGetTopY(wpfBlock);
            double bottomY = TryGetBottomY(wpfBlock);

            // Y 를 측정할 수 없으면 minSlot 슬롯에 배정하고 다음 블록으로.
            if (double.IsNaN(topY))
            {
                result.Add((minSlot / colCount, minSlot % colCount, coreBlock, Rect.Empty));
                prevSlot = minSlot;
                continue;
            }

            double blockH  = (!double.IsNaN(bottomY) && bottomY > topY) ? (bottomY - topY) : 0.0;
            // 연속 스크롤 공간에서 "단 슬롯" 인덱스 (단 슬롯 = 단 × 페이지).
            // minSlot 을 하한으로 적용해 강제 페이지 나누기 이후 블록이 이전 페이지로
            // 돌아가지 않도록 한다. 상한 클램프 없음 — 호출자(Paginate) 가 보정.
            int slotTop = Math.Max(minSlot, (int)(topY / bodyH));

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
                bool   fillOverflow    = slotFill.GetValueOrDefault(slotTop, 0.0) + blockH > bodyH + BoundaryTol;

                if (crossesBoundary || fillOverflow)
                {
                    // 분할 Y: 자연 슬롯 경계(crossesBoundary) 우선, 아니면 누적 채움 기반.
                    double splitY = crossesBoundary
                        ? slotBoundaryY
                        : topY + Math.Max(0.0, bodyH - slotFill.GetValueOrDefault(slotTop, 0.0));

                    int splitCharOffset = FindSplitCharOffset(wpfPara, splitY);
                    int totalChars      = corePara.Runs.Sum(r => r.Text.Length);

                    if (splitCharOffset > 0 && splitCharOffset < totalChars)
                    {
                        var (frag1, frag2) = SplitCoreParagraph(corePara, splitCharOffset);

                        // 첫 조각 → 현재 슬롯
                        double frag1H = Math.Max(0.0, splitY - topY);
                        slotFill[slotTop] = Math.Min(bodyH,
                            slotFill.GetValueOrDefault(slotTop, 0.0) + frag1H);
                        var rect1 = TryGetColumnLocalRect(
                            wpfBlock, slotTop / colCount, slotTop % colCount, bodyH, colWidth, colCount);
                        result.Add((slotTop / colCount, slotTop % colCount, frag1, rect1));

                        // 이어지는 조각 → 다음 슬롯 (누적 채움 검사)
                        int    nextSlot = slotTop + 1;
                        double frag2H   = blockH - frag1H;
                        if (frag2H > 0 && frag2H < bodyH)
                        {
                            while (slotFill.GetValueOrDefault(nextSlot, 0.0) + frag2H > bodyH + BoundaryTol)
                                nextSlot++;
                        }
                        if (frag2H > 0)
                            slotFill[nextSlot] = Math.Min(bodyH,
                                slotFill.GetValueOrDefault(nextSlot, 0.0) + Math.Min(frag2H, bodyH));
                        result.Add((nextSlot / colCount, nextSlot % colCount, frag2, Rect.Empty));

                        prevSlot = nextSlot;
                        continue; // 아래 단일 블록 처리 생략
                    }
                }
            }

            // ── 단일 블록 배정 (기존 로직) ────────────────────────────────────────────
            // 블록이 단 슬롯 경계를 넘고 한 슬롯에 들어갈 만큼 작으면 다음 슬롯으로 이동.
            // BoundaryTol 이하의 초과는 mm→DIP 변환 오차로 보고 현재 슬롯에 유지한다.
            // 강제 페이지 나누기 단락은 위에서 이미 슬롯을 끌어올렸으므로 이 보정을 건너뛴다
            // (자연 Y 기준 boundary 검사가 강제 슬롯과 어긋나면 잘못된 +1 이 일어날 수 있다).
            if (!isPageBreak
                && !double.IsNaN(bottomY)
                && bottomY > (slotTop + 1) * bodyH + BoundaryTol
                && blockH < bodyH)
            {
                slotTop += 1;
            }

            // 슬롯 누적 채움이 이 블록을 수용하기에 부족하면 다음 슬롯으로 밀어낸다.
            // bodyH 이상인 블록은 분할 불가이므로 채움 추적 대상에서 제외.
            if (blockH > 0 && blockH < bodyH)
            {
                while (slotFill.GetValueOrDefault(slotTop, 0.0) + blockH > bodyH + BoundaryTol)
                    slotTop += 1;
            }

            // 슬롯 채움 갱신 (bodyH 캡 — 단 높이를 초과하는 블록은 슬롯 전체를 포화로 표시)
            if (blockH > 0)
                slotFill[slotTop] = Math.Min(bodyH,
                    slotFill.GetValueOrDefault(slotTop, 0.0) + Math.Min(blockH, bodyH));

            var bodyLocalRect = TryGetColumnLocalRect(
                wpfBlock, slotTop / colCount, slotTop % colCount, bodyH, colWidth, colCount);
            result.Add((slotTop / colCount, slotTop % colCount, coreBlock, bodyLocalRect));
            prevSlot = slotTop;
        }

        // RichTextBox 분리 (FlowDocument 재사용을 위해)
        rtb.Document = new WpfDocs.FlowDocument();
        return result;
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
        // CodeLanguage / IsThematicBreak 누락. 호출측에서 second.Style.ListMarker = null 로 마커 제거.
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

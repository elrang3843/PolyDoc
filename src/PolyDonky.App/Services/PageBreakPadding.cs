using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Threading;
using PolyDonky.Core;
using Wpf = System.Windows.Documents;

namespace PolyDonky.App.Services;

/// <summary>
/// 합성 페이지 갭 패딩 단락 sentinel — Tag 비교용 마커.
/// 이 마커가 박힌 BlockUIContainer 는 시각 페이지 사이 갭(padBottom + interPageGap + padTop)
/// 을 차지하는 *순수 시각 장치* 이므로 IWPF 직렬화·클립보드·plain-text 추출에서 모두 제외된다.
/// </summary>
public sealed class PageBreakPaddingMarker
{
    public static readonly PageBreakPaddingMarker Instance = new();
    private PageBreakPaddingMarker() { }
}

/// <summary>
/// 본문 RichTextBox 안에 페이지 사이 시각적 갭 + 다음 페이지 padTop 을 차지하는
/// 합성 빈 단락(BlockUIContainer)을 자동 삽입·갱신한다.
///
/// 작동 원리:
/// 1. <see cref="Schedule"/> 호출 시 200ms 디바운스 후 <see cref="Repaginate"/> 실행.
/// 2. 기존 sentinel 단락 모두 제거 후, 본문 블록을 위에서부터 측정.
/// 3. 누적 Y 가 페이지 본문 하단을 넘는 블록 직전에 합성 단락 삽입.
///    합성 단락 높이 = padBottom + InterPageGap + padTop.
/// 4. 다음 페이지 경계로 옮겨가며 반복.
///
/// 한계:
/// - 표 행이 페이지 경계를 가로지르면 표 자체는 분할되지 않고 다음 페이지 시작점으로 밀림.
/// - 매우 긴 단락(1페이지 분량 초과) 도 분할 없이 다음 페이지 시작점으로 밀림.
/// 이는 옵션 C 의 알려진 한계.
/// </summary>
public sealed class PageBreakPadder
{
    private readonly RichTextBox _editor;
    private readonly Func<PageGeometry?> _geometryProvider;
    private readonly DispatcherTimer _timer;
    private bool _running;

    /// <summary>
    /// <paramref name="suppressTextChanged"/> 콜백은 Repaginate 가 본문을 수정하는 동안 외부의
    /// TextChanged 가드를 켜두는 데 사용 (true=on, false=off).
    /// </summary>
    public PageBreakPadder(RichTextBox editor, Func<PageGeometry?> geometryProvider,
                           Action<bool>? suppressTextChanged = null)
    {
        _editor = editor;
        _geometryProvider = geometryProvider;
        _suppressTextChanged = suppressTextChanged;
        _timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(200) };
        _timer.Tick += (_, _) => { _timer.Stop(); SafeRepaginate(); };
    }

    private readonly Action<bool>? _suppressTextChanged;

    /// <summary>본문 변경 후 200ms 안에 추가 변경이 없으면 재페이지네이션.</summary>
    public void Schedule()
    {
        _timer.Stop();
        _timer.Start();
    }

    /// <summary>지금 즉시 재페이지네이션 (페이지 설정 변경 등 동기 트리거용).</summary>
    public void RunNow()
    {
        _timer.Stop();
        SafeRepaginate();
    }

    /// <summary>본문 블록 컬렉션에서 sentinel 단락만 모두 제거. 외부 Save/Parse 경로에서 사용.</summary>
    public static void RemoveAll(BlockCollection blocks)
    {
        var toRemove = blocks.Where(b => ReferenceEquals(b.Tag, PageBreakPaddingMarker.Instance)).ToList();
        foreach (var b in toRemove) blocks.Remove(b);
    }

    /// <summary>해당 Block 이 합성 페이지 갭 패딩이면 true.</summary>
    public static bool IsPagePadding(Wpf.Block? block)
        => block is not null && ReferenceEquals(block.Tag, PageBreakPaddingMarker.Instance);

    private void SafeRepaginate()
    {
        if (_running) return;
        _running = true;
        _suppressTextChanged?.Invoke(true);
        try { Repaginate(); }
        catch { /* 측정 실패는 무시 — 다음 디바운스에서 재시도 */ }
        finally
        {
            _suppressTextChanged?.Invoke(false);
            _running = false;
        }
    }

    private void Repaginate()
    {
        var pg = _geometryProvider();
        if (pg is null) return;

        var doc = _editor.Document;
        // 1. 기존 sentinel 제거
        RemoveAll(doc.Blocks);

        // BodyEditor 가 측정될 만큼 layout pass 가 끝났는지 확인
        if (!_editor.IsLoaded || _editor.ActualWidth < 1) return;

        double pageHeight    = pg.PageHeightDip;
        double padTop        = pg.PadTopDip;
        double padBottom     = pg.PadBottomDip;
        double gap           = PageGeometry.InterPageGapDip;
        double bodyHeight    = pageHeight - padTop - padBottom;
        // 페이지 N (1-base) 의 PaperHost 좌표 — corner 와 본문 시작/끝.
        // 모든 페이지가 같은 stride (pageHeight + gap) 로 떨어져 있다.
        double pageStride = pageHeight + gap;

        if (bodyHeight <= 0) return;

        // 2. 페이지 경계마다 합성 패딩 단락 삽입.
        // BodyEditor 의 Padding 이 (padL, padT, padR, padB) 이라 첫 페이지 본문은 Y=padTop 부터.
        // 페이지 N (1-base) 의 본문 끝 Y = N*pageStride - padBottom - gap (= 정확히 페이지 N 의 corner + bodyHeight + padTop).
        // 합성 패딩 높이는 *가변* — 첫 overflow 블록이 다음 페이지 본문 시작 Y(targetY) 에
        // 정확히 안착하도록 (targetY - currentBlockY) 로 계산한다. 고정 (padBottom+gap+padTop)
        // 으로 두면 straddle 블록 이동 시 페이지 N+1 본문 시작 위치가 어긋난다.
        int pageIndex = 0;
        while (true)
        {
            double pageBodyEndY = pageIndex * pageStride + padTop + bodyHeight;
            double nextPageBodyStartY = (pageIndex + 1) * pageStride + padTop;

            // GetCharacterRect 는 layout 갱신이 필요할 수 있어 UpdateLayout 먼저.
            _editor.UpdateLayout();

            // sentinel 제외한 블록 리스트
            var liveBlocks = doc.Blocks.Where(b => !IsPagePadding(b)).ToList();
            if (liveBlocks.Count == 0) break;

            Wpf.Block? firstOverflow = null;
            for (int i = 0; i < liveBlocks.Count; i++)
            {
                var b = liveBlocks[i];

                double startY = TryGetTopY(b);
                if (double.IsNaN(startY)) continue;

                // 시작 Y 가 이미 페이지 끝을 넘으면 → 이 블록부터 다음 페이지로 가야 함
                if (startY >= pageBodyEndY)
                {
                    firstOverflow = b;
                    break;
                }

                // 시작은 이 페이지 안이지만 *끝* 이 페이지 경계를 넘는지 확인.
                // 다음 블록의 시작 Y = 현재 블록의 끝 Y. 마지막 블록이면 ContentEnd.Bottom 사용.
                // 끝이 넘으면 (단락이 페이지 경계를 가로지르면) 통째로 다음 페이지로 밀어낸다 (블록 분할은 안 함).
                double endY = double.NaN;
                if (i + 1 < liveBlocks.Count)
                {
                    endY = TryGetTopY(liveBlocks[i + 1]);
                }
                if (double.IsNaN(endY))
                {
                    try
                    {
                        var endRect = b.ContentEnd.GetCharacterRect(LogicalDirection.Backward);
                        if (!double.IsInfinity(endRect.Y) && !double.IsNaN(endRect.Y))
                            endY = endRect.Y + endRect.Height;
                    }
                    catch { }
                }
                if (!double.IsNaN(endY) && endY > pageBodyEndY)
                {
                    firstOverflow = b;
                    break;
                }
            }

            if (firstOverflow is null) break; // 더 이상 넘는 블록 없음 — 종료

            // 가변 패딩 높이 — 이 블록이 다음 페이지 본문 시작 Y 에 정확히 안착하도록.
            double currentY = TryGetTopY(firstOverflow);
            if (double.IsNaN(currentY)) break;
            double H = nextPageBodyStartY - currentY;
            // 최소 보정 — 음수/너무 작은 값 방어 (블록이 이미 다음 페이지 본문 시작 Y 를 넘어 있는 비정상 상황).
            if (H < 1.0) H = padBottom + gap + padTop;

            var padding = BuildSyntheticPaddingBlock(H);
            doc.Blocks.InsertBefore(firstOverflow, padding);

            pageIndex++;
            if (pageIndex > 1000) break; // 안전 한계 — 무한 루프 방지
        }
    }

    /// <summary>블록의 화면상 위쪽 Y 좌표 (BodyEditor 좌표계). 측정 불가 시 NaN.</summary>
    private static double TryGetTopY(Wpf.Block b)
    {
        try
        {
            var r = b.ContentStart.GetCharacterRect(LogicalDirection.Forward);
            if (double.IsInfinity(r.Y) || double.IsNaN(r.Y)) return double.NaN;
            return r.Y;
        }
        catch
        {
            return double.NaN;
        }
    }

    private static Wpf.Block BuildSyntheticPaddingBlock(double heightDip)
    {
        var spacer = new System.Windows.Controls.Border
        {
            Height = heightDip,
            Background = System.Windows.Media.Brushes.Transparent,
            IsHitTestVisible = false,
            Focusable = false,
        };
        return new BlockUIContainer(spacer)
        {
            Tag = PageBreakPaddingMarker.Instance,
            Margin = new Thickness(0),
            Padding = new Thickness(0),
        };
    }
}

using System.Linq;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Threading;
using PolyDonky.Core;

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
    public static bool IsPagePadding(Block? block)
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
        double syntheticH    = padBottom + gap + padTop;

        if (bodyHeight <= 0) return;

        // 2. 페이지 경계마다 합성 패딩 단락 삽입.
        // BodyEditor 의 Padding 이 (padL, padT, padR, padB) 이라 첫 페이지 본문은 Y=padTop 부터.
        // 페이지 N (1-base) 의 본문 끝 Y = padTop + N*bodyHeight + (N-1)*syntheticH.
        int pageIndex = 0;
        while (true)
        {
            double pageBodyEndY = padTop + (pageIndex + 1) * bodyHeight + pageIndex * syntheticH;

            // 본문을 끝에서부터 훑어 이 Y 를 넘어가는 첫 블록을 찾음.
            // GetCharacterRect 는 layout 갱신이 필요할 수 있어 UpdateLayout 먼저.
            _editor.UpdateLayout();

            Block? firstOverflow = null;
            foreach (var b in doc.Blocks)
            {
                Rect rect;
                try { rect = b.ContentStart.GetCharacterRect(LogicalDirection.Forward); }
                catch { continue; }
                if (double.IsInfinity(rect.Y) || double.IsNaN(rect.Y)) continue;
                // 블록 시작 Y 가 페이지 끝을 이미 넘었으면 이 블록은 다음 페이지로 밀려야 함.
                if (rect.Y >= pageBodyEndY)
                {
                    firstOverflow = b;
                    break;
                }
            }

            if (firstOverflow is null) break; // 더 이상 넘는 블록 없음 — 종료

            var padding = BuildSyntheticPaddingBlock(syntheticH);
            doc.Blocks.InsertBefore(firstOverflow, padding);

            pageIndex++;
            if (pageIndex > 1000) break; // 안전 한계 — 무한 루프 방지
        }
    }

    private static Block BuildSyntheticPaddingBlock(double heightDip)
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

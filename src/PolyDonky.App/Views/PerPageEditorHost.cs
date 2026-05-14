using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using PolyDonky.App.Pagination;
using PolyDonky.App.Services;

namespace PolyDonky.App.Views;

/// <summary>
/// 페이지별·단별 RichTextBox 를 Canvas 위에 배치하는 호스트.
/// <see cref="PerPageDocumentSplitter.Split"/> 결과를 받아 슬라이스마다 RTB 를 하나씩 생성한다.
/// <para>
/// 단일 단: RTB 1개/페이지, 전체 페이지 크기 + 여백(padding).
/// 다단: RTB N개/페이지 (N = 단 수), 각 RTB 는 단 폭 × 본문 높이, 여백·단 오프셋은 Canvas 위치로.
/// </para>
/// <para>STA 스레드 전용.</para>
/// </summary>
public sealed class PerPageEditorHost : Canvas
{
    private readonly List<RichTextBox> _pageEditors = new();
    private int                        _physicalPageCount;

    /// <summary>현재 키보드 포커스를 가진 페이지 RTB.</summary>
    public RichTextBox? ActiveEditor { get; private set; }

    /// <summary>첫 번째 RTB. 없으면 null.</summary>
    public RichTextBox? FirstEditor => _pageEditors.Count > 0 ? _pageEditors[0] : null;

    /// <summary>물리 페이지 수 (단 수와 무관하게 실제 페이지 수).</summary>
    public int PageCount => _physicalPageCount;

    /// <summary>생성된 모든 RTB 목록 (페이지 순, 단 순).</summary>
    public IReadOnlyList<RichTextBox> PageEditors => _pageEditors;

    /// <summary>임의 RTB 의 텍스트가 변경되면 발화.</summary>
    public event TextChangedEventHandler? PageTextChanged;

    /// <summary>모든 RTB 의 FlowDocument.Blocks 를 순서대로 열거.</summary>
    public IEnumerable<System.Windows.Documents.Block> AllBlocks
        => _pageEditors.SelectMany(e => e.Document.Blocks);

    /// <summary>
    /// 슬라이스 목록을 받아 기존 RTB 를 모두 교체한다.
    /// </summary>
    /// <param name="slices"><see cref="PerPageDocumentSplitter.Split"/> 결과.</param>
    /// <param name="geo">현재 페이지 기하 정보.</param>
    /// <param name="configure">각 RTB 생성 후 호출되는 콜백 — 이벤트 구독·속성 설정.</param>
    /// <summary>
    /// mm→DIP 변환과 WPF 장치 픽셀 반올림의 미세한 차이로 인해
    /// RTB 의 실제 콘텐츠 표시 영역(Height - Padding)이 페이지네이터가 계산한
    /// bodyH 보다 1-2 DIP 작을 수 있다. 이 차이가 마지막 줄의 하단을
    /// ScrollBarVisibility.Hidden/Disabled 로 잘라낸다.
    /// 2 DIP 여유를 단일 단 RTB 하단 패딩 축소 / 다단 RTB 높이 확장으로 흡수한다.
    /// </summary>
    private const double ClipRenderingTolerance = 2.0;

    public void SetupPages(
        IReadOnlyList<PerPageDocumentSlice> slices,
        PageGeometry                        geo,
        Action<RichTextBox>?                configure = null)
    {
        foreach (var e in _pageEditors) e.TextChanged -= OnPageTextChanged;
        _pageEditors.Clear();
        Children.Clear();
        ActiveEditor        = null;
        _physicalPageCount  = 0;

        if (slices.Count == 0) return;

        _physicalPageCount = slices.Max(s => s.PageIndex) + 1;

        // 페이지별 누적 Y 좌표 — 각 페이지가 다른 PageSettings(높이)를 가질 수 있다.
        // 키 = pageIndex, 값 = 해당 페이지 상단 Y (Canvas 좌표).
        var pageTopY = new Dictionary<int, double>();
        double cumulY = 0;
        for (int pi = 0; pi < _physicalPageCount; pi++)
        {
            pageTopY[pi] = cumulY;
            // 이 페이지의 PageSettings: 해당 pageIndex 의 첫 슬라이스에서 가져온다.
            var pageSlice = slices.FirstOrDefault(s => s.PageIndex == pi);
            var pageGeoForHeight = pageSlice is not null
                ? new PageGeometry(pageSlice.PageSettings)
                : geo;
            cumulY += pageGeoForHeight.PageStrideDip;
        }

        for (int i = 0; i < slices.Count; i++)
        {
            var slice    = slices[i];
            var sliceGeo = new PageGeometry(slice.PageSettings);

            // 단일 단 / 다단 모두 동일하게 — 콘텐츠 영역(BodyWidth × BodyHeight)만 차지하는 RTB 를
            // Canvas 좌표로 페이지 여백 안쪽에 배치한다. 단일 단에서 RTB.Width = PageWidth + Padding
            // = 여백 으로 두면 측정 RTB(Padding=0, fd.PageWidth=ColWidth)와 콘텐츠 영역 layout 이
            // 미세하게 어긋나(WPF 가 RTB.Padding 과 fd.PageWidth 를 다르게 처리) pagination 측정값
            // 과 실렌더 길이가 달라져 페이지 끝에서 클리핑이 일어났다. 두 RTB 의 layout 조건을
            // 동일하게 맞추는 것이 측정 정확도의 전제 — 단일 단도 다단과 동일 패턴으로 통일.
            var rtb = new RichTextBox
            {
                Document      = slice.FlowDocument,
                Width         = slice.BodyWidthDip,
                // ClipRenderingTolerance 만큼 높이를 늘려 mm→DIP 반올림 오차로 인한
                // 마지막 줄 하단 클리핑을 방지한다. 추가 영역은 하단 여백 안에 위치.
                Height        = slice.BodyHeightDip + ClipRenderingTolerance,
                Padding       = new Thickness(0),
                // Disabled 로 설정해야 RTB 가 슬롯 높이를 초과하는 콘텐츠를 내부 스크롤로
                // 처리하지 않는다(Hidden 이면 스크롤이 발생해 편집창과 인쇄 미리보기가 달라진다).
                VerticalScrollBarVisibility   = ScrollBarVisibility.Disabled,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                AcceptsReturn     = true,
                AcceptsTab        = true,
                Background        = Brushes.Transparent,
                BorderThickness   = new Thickness(0),
                FlowDirection     = FlowDirection.LeftToRight,
                IsDocumentEnabled = true,
            };
            double xPos = sliceGeo.PadLeftDip + slice.XOffsetDip;
            double yPos = pageTopY[slice.PageIndex] + sliceGeo.PadTopDip;
            SetLeft(rtb, xPos);
            SetTop (rtb, yPos);

            rtb.PreviewMouseLeftButtonDown += (_, _) => ActiveEditor = rtb;
            rtb.GotKeyboardFocus           += (_, _) => ActiveEditor = rtb;
            rtb.TextChanged                += OnPageTextChanged;

            configure?.Invoke(rtb);

            Children.Add(rtb);
            _pageEditors.Add(rtb);
        }

        ActiveEditor = _pageEditors[0];
        Width  = geo.PageWidthDip;
        Height = cumulY;
    }

    private void OnPageTextChanged(object sender, TextChangedEventArgs e)
        => PageTextChanged?.Invoke(sender, e);
}

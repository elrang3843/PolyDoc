using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using PolyDonky.App.Pagination;
using PolyDonky.App.Services;
using PolyDonky.Core;

namespace PolyDonky.App.Views;

/// <summary>
/// 페이지별·단별 RichTextBox 를 Canvas 위에 배치하는 호스트.
/// <see cref="PerPageDocumentSplitter.Split"/> 결과를 받아 슬라이스마다 RTB 를 하나씩 생성한다.
/// <para>
/// 단일 단: RTB 1개/페이지, 전체 페이지 크기 + 여백(padding).
/// 다단: RTB N개/페이지 (N = 단 수), 각 RTB 는 단 폭 × 본문 높이, 여백·단 오프셋은 Canvas 위치로.
/// </para>
/// <para>
/// 각주 레이아웃: 본문 RTB 높이는 각주 크기만큼 줄어든다. 각주 구분선·내용은
/// RTB 바로 아래(= 꼬리말 위)에 배치돼 꼬리말 영역을 침범하지 않는다.
/// </para>
/// <para>STA 스레드 전용.</para>
/// </summary>
public sealed class PerPageEditorHost : Canvas
{
    private readonly List<RichTextBox> _pageEditors = new();
    private int                        _physicalPageCount;
    private int                        _endnotePageStartIndex = -1;

    /// <summary>현재 키보드 포커스를 가진 페이지 RTB.</summary>
    public RichTextBox? ActiveEditor { get; private set; }

    /// <summary>첫 번째 RTB. 없으면 null.</summary>
    public RichTextBox? FirstEditor => _pageEditors.Count > 0 ? _pageEditors[0] : null;

    /// <summary>물리 페이지 수 (미주 페이지 포함 전체 페이지 수).</summary>
    public int PageCount => _physicalPageCount;

    /// <summary>미주 페이지 제외 본문 페이지 수.</summary>
    public int BodyPageCount => HasEndnotePage ? _physicalPageCount - 1 : _physicalPageCount;

    /// <summary>미주 페이지가 있으면 true.</summary>
    public bool HasEndnotePage => _endnotePageStartIndex >= 0;

    /// <summary>미주 페이지의 첫 번째 PageIndex. 없으면 -1.</summary>
    public int EndnotePageStartIndex => _endnotePageStartIndex;

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
        ActiveEditor           = null;
        _physicalPageCount     = 0;
        _endnotePageStartIndex = -1;

        if (slices.Count == 0) return;

        _physicalPageCount     = slices.Max(s => s.PageIndex) + 1;
        _endnotePageStartIndex = slices.FirstOrDefault(s => s.IsEndnotePage)?.PageIndex ?? -1;

        // 페이지별 PageGeometry 캐시 및 누적 Y 좌표 계산.
        var pageGeos = new Dictionary<int, PageGeometry>(_physicalPageCount);
        var pageTopY = new Dictionary<int, double>(_physicalPageCount);
        double cumulY = 0;
        for (int pi = 0; pi < _physicalPageCount; pi++)
        {
            var pageSlice = slices.FirstOrDefault(s => s.PageIndex == pi);
            var pageGeo = pageSlice is not null
                ? new PageGeometry(pageSlice.PageSettings)
                : geo;
            pageGeos[pi] = pageGeo;
            pageTopY[pi] = cumulY;
            cumulY += pageGeo.PageStrideDip;
        }

        for (int i = 0; i < slices.Count; i++)
        {
            var slice    = slices[i];
            var sliceGeo = pageGeos[slice.PageIndex];

            double xPos = sliceGeo.PadLeftDip + slice.XOffsetDip;
            double yPos = pageTopY[slice.PageIndex] + sliceGeo.PadTopDip;

            // 본문 RTB
            var rtb = new RichTextBox
            {
                Document      = slice.FlowDocument,
                Width         = slice.BodyWidthDip,
                // ClipRenderingTolerance: mm→DIP 반올림 오차로 인한 마지막 줄 클리핑 방지.
                Height        = slice.BodyHeightDip + ClipRenderingTolerance,
                Padding       = new Thickness(0),
                VerticalScrollBarVisibility   = ScrollBarVisibility.Disabled,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                AcceptsReturn     = true,
                AcceptsTab        = true,
                Background        = Brushes.Transparent,
                BorderThickness   = new Thickness(0),
                FlowDirection     = FlowDirection.LeftToRight,
                IsDocumentEnabled = true,
            };

            // 미주 페이지는 편집 불가.
            if (slice.IsEndnotePage)
            {
                rtb.IsReadOnly = true;
                rtb.Focusable  = false;
            }

            SetLeft(rtb, xPos);
            SetTop (rtb, yPos);

            if (!slice.IsEndnotePage)
            {
                rtb.PreviewMouseLeftButtonDown += (_, _) => ActiveEditor = rtb;
                rtb.GotKeyboardFocus           += (_, _) => ActiveEditor = rtb;
                rtb.TextChanged                += OnPageTextChanged;
            }

            configure?.Invoke(rtb);
            Children.Add(rtb);

            if (!slice.IsEndnotePage)
                _pageEditors.Add(rtb);

            // 각주 영역 — 본문 RTB 바로 아래, 꼬리말 위.
            // yPos + BodyHeightDip + ClipRenderingTolerance = RTB 하단 좌표.
            if (slice.PageFootnotes.Count > 0 && slice.FootnoteAreaHeightDip > 0)
            {
                double fnY = yPos + slice.BodyHeightDip + ClipRenderingTolerance;
                var fnPanel = BuildFootnotesPanel(slice, sliceGeo);
                SetLeft(fnPanel, xPos);
                SetTop (fnPanel, fnY);
                Children.Add(fnPanel);
            }
        }

        if (_pageEditors.Count > 0)
            ActiveEditor = _pageEditors[0];

        Width  = geo.PageWidthDip;
        Height = cumulY;
    }

    private void OnPageTextChanged(object sender, TextChangedEventArgs e)
        => PageTextChanged?.Invoke(sender, e);

    // ── 각주 패널 렌더링 ──────────────────────────────────────────────────────

    private static StackPanel BuildFootnotesPanel(PerPageDocumentSlice slice, PageGeometry sliceGeo)
    {
        var panel = new StackPanel
        {
            Width       = slice.BodyWidthDip,
            Orientation = Orientation.Vertical,
            Background  = Brushes.Transparent,
        };

        // 구분선
        panel.Children.Add(new Border
        {
            Height          = 1,
            Margin          = new Thickness(0, 4, 0, 4),
            BorderBrush     = new SolidColorBrush(System.Windows.Media.Color.FromRgb(180, 180, 180)),
            BorderThickness = new Thickness(0, 1, 0, 0),
            Width           = Math.Min(slice.BodyWidthDip / 3.0, 100),
            HorizontalAlignment = HorizontalAlignment.Left,
        });

        // 각주 본문 RTB (실제 측정에 사용한 것과 동일한 FlowDocument)
        var fnNums = slice.PageFootnotes.Count > 0
            ? BuildLocalFnNums(slice.PageFootnotes)
            : null;

        var fd = PerPageDocumentSplitter.BuildFootnoteFlowDocument(
            slice.PageFootnotes, fnNums, slice.BodyWidthDip);

        var noteRtb = new RichTextBox
        {
            Document                      = fd,
            IsReadOnly                    = true,
            BorderThickness               = new Thickness(0),
            Background                    = Brushes.Transparent,
            Padding                       = new Thickness(0),
            VerticalScrollBarVisibility   = ScrollBarVisibility.Disabled,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
            Width                         = slice.BodyWidthDip,
            Focusable                     = false,
        };
        panel.Children.Add(noteRtb);

        return panel;
    }

    // BuildFootnotesPanel 에서 사용할 로컬 번호 맵 (1-based 슬라이스 내 순번)
    private static IReadOnlyDictionary<string, int> BuildLocalFnNums(
        IReadOnlyList<FootnoteEntry> footnotes)
    {
        var d = new Dictionary<string, int>(footnotes.Count);
        for (int i = 0; i < footnotes.Count; i++)
            d[footnotes[i].Id] = i + 1;
        return d;
    }
}

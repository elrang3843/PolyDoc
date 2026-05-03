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

        bool multiColumn = slices[0].ColumnCount > 1;
        _physicalPageCount = slices.Max(s => s.PageIndex) + 1;

        for (int i = 0; i < slices.Count; i++)
        {
            var slice = slices[i];
            RichTextBox rtb;

            if (!multiColumn)
            {
                // 단일 단 — 기존 동작: 전체 페이지 크기 + 여백 padding
                rtb = new RichTextBox
                {
                    Document      = slice.FlowDocument,
                    Width         = geo.PageWidthDip,
                    Height        = geo.PageHeightDip,
                    Padding       = new Thickness(geo.PadLeftDip, geo.PadTopDip,
                                                  geo.PadRightDip, geo.PadBottomDip),
                    VerticalScrollBarVisibility   = ScrollBarVisibility.Hidden,
                    HorizontalScrollBarVisibility = ScrollBarVisibility.Hidden,
                    AcceptsReturn     = true,
                    AcceptsTab        = true,
                    Background        = Brushes.Transparent,
                    BorderThickness   = new Thickness(0),
                    FlowDirection     = FlowDirection.LeftToRight,
                    IsDocumentEnabled = true,
                };
                SetLeft(rtb, 0);
                SetTop (rtb, slice.PageIndex * geo.PageStrideDip);
            }
            else
            {
                // 다단 — 단 폭 × 본문 높이, 여백·단 오프셋은 Canvas 위치로.
                // Disabled 로 설정해야 RTB 가 단 높이를 초과하는 콘텐츠를 내부 스크롤로
                // 처리하지 않는다(Hidden 이면 스크롤이 발생해 편집창과 인쇄 미리보기가 달라진다).
                rtb = new RichTextBox
                {
                    Document      = slice.FlowDocument,
                    Width         = slice.BodyWidthDip,
                    Height        = slice.BodyHeightDip,
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
                double xPos = geo.PadLeftDip  + slice.XOffsetDip;
                double yPos = slice.PageIndex * geo.PageStrideDip + geo.PadTopDip;
                SetLeft(rtb, xPos);
                SetTop (rtb, yPos);
            }

            rtb.PreviewMouseLeftButtonDown += (_, _) => ActiveEditor = rtb;
            rtb.GotKeyboardFocus           += (_, _) => ActiveEditor = rtb;
            rtb.TextChanged                += OnPageTextChanged;

            configure?.Invoke(rtb);

            Children.Add(rtb);
            _pageEditors.Add(rtb);
        }

        ActiveEditor = _pageEditors[0];
        Width  = geo.PageWidthDip;
        Height = geo.TotalHeightDip(_physicalPageCount);
    }

    private void OnPageTextChanged(object sender, TextChangedEventArgs e)
        => PageTextChanged?.Invoke(sender, e);
}

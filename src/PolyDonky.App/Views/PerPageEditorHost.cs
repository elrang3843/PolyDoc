using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using PolyDonky.App.Pagination;
using PolyDonky.App.Services;

namespace PolyDonky.App.Views;

/// <summary>
/// 페이지별 RichTextBox 를 수직으로 배치하는 Canvas.
/// <see cref="PerPageDocumentSplitter.Split"/> 결과를 받아 슬라이스마다 RTB 를 하나씩 생성한다.
/// <para>STA 스레드 전용.</para>
/// </summary>
public sealed class PerPageEditorHost : Canvas
{
    private readonly List<RichTextBox> _pageEditors = new();

    /// <summary>현재 키보드 포커스를 가진 페이지 RTB.</summary>
    public RichTextBox? ActiveEditor { get; private set; }

    /// <summary>첫 번째 페이지 RTB. 페이지가 없으면 null.</summary>
    public RichTextBox? FirstEditor => _pageEditors.Count > 0 ? _pageEditors[0] : null;

    /// <summary>현재 구성된 페이지 수.</summary>
    public int PageCount => _pageEditors.Count;

    /// <summary>개별 페이지 RTB 목록 (본문 파싱용).</summary>
    public IReadOnlyList<RichTextBox> PageEditors => _pageEditors;

    /// <summary>임의 페이지 RTB 의 텍스트가 변경되면 발화.</summary>
    public event TextChangedEventHandler? PageTextChanged;

    /// <summary>모든 페이지 RTB 의 FlowDocument.Blocks 를 페이지 순서대로 열거.</summary>
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
        ActiveEditor = null;

        if (slices.Count == 0) return;

        for (int i = 0; i < slices.Count; i++)
        {
            var rtb = new RichTextBox
            {
                Document      = slices[i].FlowDocument,
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

            // ActiveEditor を MouseDown 시 즉시 갱신 — GotKeyboardFocus 보다 먼저 실행됨.
            rtb.PreviewMouseLeftButtonDown += (_, _) => ActiveEditor = rtb;
            rtb.GotKeyboardFocus           += (_, _) => ActiveEditor = rtb;
            rtb.TextChanged                += OnPageTextChanged;

            configure?.Invoke(rtb);

            SetLeft(rtb, 0);
            SetTop (rtb, i * geo.PageStrideDip);
            Children.Add(rtb);
            _pageEditors.Add(rtb);
        }

        ActiveEditor = _pageEditors[0];
        Width  = geo.PageWidthDip;
        Height = geo.TotalHeightDip(slices.Count);
    }

    private void OnPageTextChanged(object sender, TextChangedEventArgs e)
        => PageTextChanged?.Invoke(sender, e);
}

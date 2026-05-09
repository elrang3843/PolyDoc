using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using PolyDonky.Core;
using WpfShapes = System.Windows.Shapes;

namespace PolyDonky.App.Views;

/// <summary>
/// 글상자 내부 다단 편집을 위한 Canvas 기반 호스트.
/// <see cref="PerPageEditorHost"/> 의 단순화 버전 — 페이지 개념 없이 단(column)만 다룬다.
/// <para>
/// 단별로 별도 <see cref="RichTextBox"/> 를 만들어 가로로 배치한다.
/// 단 경계 좌/우 화살표 이동, 단별 텍스트 선택은 부모(<see cref="TextBoxOverlay"/>)
/// 가 이벤트를 받아 처리한다.
/// </para>
/// </summary>
public sealed class TextBoxColumnHost : Canvas
{
    private readonly List<RichTextBox> _editors = new();
    private readonly List<WpfShapes.Line> _dividerLines = new();

    /// <summary>현재 키보드 포커스를 가진 단 RTB.</summary>
    public RichTextBox? ActiveEditor { get; private set; }

    /// <summary>첫 번째 단 RTB (포커스 폴백용).</summary>
    public RichTextBox? FirstEditor => _editors.Count > 0 ? _editors[0] : null;

    /// <summary>모든 단 RTB 목록 (단 인덱스 순).</summary>
    public IReadOnlyList<RichTextBox> Editors => _editors;

    /// <summary><paramref name="editor"/> 의 단 인덱스 (없으면 -1).</summary>
    public int IndexOf(RichTextBox editor)
    {
        for (int i = 0; i < _editors.Count; i++)
            if (ReferenceEquals(_editors[i], editor)) return i;
        return -1;
    }

    /// <summary>임의 단 RTB 의 텍스트가 변경되면 발화.</summary>
    public event TextChangedEventHandler? ColumnTextChanged;

    /// <summary>단별 RTB 들을 슬라이스로 재구성한다.</summary>
    /// <param name="slices">단 분배 결과.</param>
    /// <param name="configure">각 RTB 생성 후 호출되는 콜백 — 이벤트 구독·속성 설정.</param>
    public void SetupColumns(
        IReadOnlyList<TextBoxColumnLayout.ColumnSlice> slices,
        Action<RichTextBox>?                           configure = null)
    {
        foreach (var e in _editors) e.TextChanged -= OnAnyTextChanged;
        _editors.Clear();
        Children.Clear();
        ActiveEditor = null;

        if (slices.Count == 0) return;

        foreach (var slice in slices)
        {
            // 클리핑 여유 — 본문 다단(PerPageEditorHost.ClipRenderingTolerance) 와 동일.
            const double ClipTol = 2.0;
            var rtb = new RichTextBox
            {
                Document        = slice.FlowDocument,
                Width           = slice.WidthDip,
                Height          = slice.HeightDip + ClipTol,
                Padding         = new Thickness(0),
                VerticalScrollBarVisibility   = ScrollBarVisibility.Disabled,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                AcceptsReturn   = true,
                AcceptsTab      = true,
                Background      = Brushes.Transparent,
                BorderThickness = new Thickness(0),
                FlowDirection   = FlowDirection.LeftToRight,
                IsDocumentEnabled = true,
            };
            SetLeft(rtb, slice.XOffsetDip);
            SetTop (rtb, 0);

            rtb.PreviewMouseLeftButtonDown += (_, _) => ActiveEditor = rtb;
            rtb.GotKeyboardFocus           += (_, _) => ActiveEditor = rtb;
            rtb.TextChanged                += OnAnyTextChanged;

            configure?.Invoke(rtb);

            Children.Add(rtb);
            _editors.Add(rtb);
        }

        ActiveEditor = _editors[0];

        // 호스트 자체 크기 — 모든 단을 감싸도록.
        if (slices.Count > 0)
        {
            var last = slices[^1];
            Width  = last.XOffsetDip + last.WidthDip;
            Height = slices.Max(s => s.HeightDip);
        }
    }

    /// <summary>
    /// 단 사이 구분선을 그리거나 지운다.
    /// <paramref name="slices"/> 는 마지막 <see cref="SetupColumns"/> 호출과 동일해야 한다.
    /// </summary>
    public void UpdateDividerLines(
        IReadOnlyList<TextBoxColumnLayout.ColumnSlice> slices,
        bool               visible,
        string             colorHex,
        double             thicknessPt,
        ColumnDividerStyle style)
    {
        foreach (var line in _dividerLines) Children.Remove(line);
        _dividerLines.Clear();

        if (!visible || slices.Count < 2) return;

        var brush = ParseBrush(colorHex);
        double thickDip = thicknessPt * (96.0 / 72.0);
        DoubleCollection? dash = style switch
        {
            ColumnDividerStyle.Dashed => new DoubleCollection { 4, 3 },
            ColumnDividerStyle.Dotted => new DoubleCollection { 1, 2 },
            _                         => null,
        };
        if (style == ColumnDividerStyle.None) return;

        double hostHeight = slices.Max(s => s.HeightDip);

        for (int i = 0; i < slices.Count - 1; i++)
        {
            var curr = slices[i];
            var next = slices[i + 1];
            double midX = (curr.XOffsetDip + curr.WidthDip + next.XOffsetDip) / 2.0;

            var line = new WpfShapes.Line
            {
                X1               = midX,
                Y1               = 0,
                X2               = midX,
                Y2               = hostHeight,
                Stroke           = brush,
                StrokeThickness  = thickDip,
                IsHitTestVisible = false,
                SnapsToDevicePixels = true,
            };
            if (dash is not null) line.StrokeDashArray = dash;
            Children.Add(line);
            _dividerLines.Add(line);
        }
    }

    private static Brush ParseBrush(string? hex)
    {
        if (!string.IsNullOrWhiteSpace(hex))
        {
            try
            {
                var s = hex.Trim();
                if (!s.StartsWith('#')) s = '#' + s;
                var c = (System.Windows.Media.Color)ColorConverter.ConvertFromString(s)!;
                return new SolidColorBrush(c);
            }
            catch { }
        }
        return new SolidColorBrush(System.Windows.Media.Color.FromRgb(0x88, 0x88, 0x88));
    }

    private void OnAnyTextChanged(object sender, TextChangedEventArgs e)
        => ColumnTextChanged?.Invoke(sender, e);
}

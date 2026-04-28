using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace PolyDonky.App.Controls;

/// <summary>색상 입력 컨트롤. 직접 입력 + OS 색상 선택기 + 없음(지우기) 버튼을 제공한다.</summary>
public partial class ColorPickerBox : UserControl
{
    // ── DependencyProperties ─────────────────────────────────────────────

    public static readonly DependencyProperty ColorTextProperty =
        DependencyProperty.Register(
            nameof(ColorText), typeof(string), typeof(ColorPickerBox),
            new FrameworkPropertyMetadata(
                string.Empty,
                FrameworkPropertyMetadataOptions.BindsTwoWayByDefault,
                OnColorTextChanged));

    public static readonly DependencyProperty AllowEmptyProperty =
        DependencyProperty.Register(
            nameof(AllowEmpty), typeof(bool), typeof(ColorPickerBox),
            new PropertyMetadata(true, OnAllowEmptyChanged));

    /// <summary>현재 색상 문자열 (#RRGGBB 또는 이름). 빈 문자열 = 색상 없음.</summary>
    public string ColorText
    {
        get => (string)GetValue(ColorTextProperty);
        set => SetValue(ColorTextProperty, value);
    }

    /// <summary>true 이면 "없음" 버튼을 표시한다.</summary>
    public bool AllowEmpty
    {
        get => (bool)GetValue(AllowEmptyProperty);
        set => SetValue(AllowEmptyProperty, value);
    }

    private bool _syncing;

    public ColorPickerBox()
    {
        InitializeComponent();
    }

    // ── 속성 변경 콜백 ───────────────────────────────────────────────────

    private static void OnColorTextChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        var ctrl = (ColorPickerBox)d;
        if (ctrl._syncing) return;
        ctrl._syncing = true;
        ctrl.ColorTextBox.Text = (string?)e.NewValue ?? string.Empty;
        ctrl.UpdateSwatch();
        ctrl._syncing = false;
    }

    private static void OnAllowEmptyChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        => ((ColorPickerBox)d).ClearButton.Visibility =
            (bool)e.NewValue ? Visibility.Visible : Visibility.Collapsed;

    // ── 이벤트 핸들러 ────────────────────────────────────────────────────

    private void OnTextChanged(object sender, TextChangedEventArgs e)
    {
        if (_syncing) return;
        _syncing = true;
        SetCurrentValue(ColorTextProperty, ColorTextBox.Text);
        UpdateSwatch();
        _syncing = false;
    }

    private void OnPickClick(object sender, RoutedEventArgs e)
    {
        var dlg = new System.Windows.Forms.ColorDialog
        {
            FullOpen = true,
            AnyColor = true,
        };

        if (TryParseColor(ColorText, out var current))
            dlg.Color = System.Drawing.Color.FromArgb(current.A, current.R, current.G, current.B);

        if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            var c = dlg.Color;
            ColorText = $"#{c.R:X2}{c.G:X2}{c.B:X2}";
        }
    }

    private void OnClearClick(object sender, RoutedEventArgs e)
        => ColorText = string.Empty;

    // ── 스와치 갱신 ──────────────────────────────────────────────────────

    private void UpdateSwatch()
    {
        Swatch.Background = TryParseColor(ColorTextBox.Text, out var c)
            ? new SolidColorBrush(c)
            : null;
    }

    // ── 정적 유틸 (다른 창에서도 사용 가능) ─────────────────────────────

    /// <summary>hex 문자열을 WPF Color 로 파싱. 실패하면 false.</summary>
    public static bool TryParseColor(string? hex, out Color color)
    {
        color = default;
        if (string.IsNullOrWhiteSpace(hex)) return false;
        try
        {
            if (System.Windows.Media.ColorConverter.ConvertFromString(hex) is Color c)
            {
                color = c;
                return true;
            }
        }
        catch { }
        return false;
    }
}

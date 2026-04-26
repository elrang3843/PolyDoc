using System.Windows;
using System.Windows.Media;
using PolyDoc.Core;
using WpfMedia = System.Windows.Media;

namespace PolyDoc.App.Views;

/// <summary>글상자 테두리·배경·여백·정렬 편집 대화상자.</summary>
public partial class TextBoxPropertiesWindow : Window
{
    public string?       ResultBorderColor      { get; private set; }
    public double        ResultBorderThicknessPt { get; private set; }
    public string?       ResultBackgroundColor  { get; private set; }
    public double        ResultPaddingTopMm     { get; private set; }
    public double        ResultPaddingBottomMm  { get; private set; }
    public double        ResultPaddingLeftMm    { get; private set; }
    public double        ResultPaddingRightMm   { get; private set; }
    public TextBoxHAlign ResultHAlign           { get; private set; }
    public TextBoxVAlign ResultVAlign           { get; private set; }

    public TextBoxPropertiesWindow(TextBoxObject model)
    {
        InitializeComponent();

        TxtBorderColor.Text     = model.BorderColor ?? "#000000";
        TxtBorderThickness.Text = model.BorderThicknessPt.ToString("0.##");
        TxtBackgroundColor.Text = model.BackgroundColor ?? "";

        TxtPaddingTop.Text    = model.PaddingTopMm.ToString("0.##");
        TxtPaddingBottom.Text = model.PaddingBottomMm.ToString("0.##");
        TxtPaddingLeft.Text   = model.PaddingLeftMm.ToString("0.##");
        TxtPaddingRight.Text  = model.PaddingRightMm.ToString("0.##");

        CboHAlign.SelectedIndex = (int)model.HAlign;
        CboVAlign.SelectedIndex = (int)model.VAlign;

        RefreshColorButton(BtnBorderColorPick,     TxtBorderColor.Text);
        RefreshColorButton(BtnBackgroundColorPick, TxtBackgroundColor.Text);
    }

    // ── 색 텍스트 변경 → 미리보기 버튼 갱신 ────────────────────────

    private void OnBorderColorChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        => RefreshColorButton(BtnBorderColorPick, TxtBorderColor.Text);

    private void OnBackgroundColorChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        => RefreshColorButton(BtnBackgroundColorPick, TxtBackgroundColor.Text);

    private static void RefreshColorButton(System.Windows.Controls.Button btn, string? hex)
    {
        if (string.IsNullOrWhiteSpace(hex))
        {
            btn.Background = System.Windows.Media.Brushes.Transparent;
            return;
        }
        try
        {
            var s = hex.Trim();
            if (!s.StartsWith('#')) s = '#' + s;
            var c = (WpfMedia.Color)WpfMedia.ColorConverter.ConvertFromString(s)!;
            btn.Background = new SolidColorBrush(c);
        }
        catch { btn.Background = System.Windows.Media.Brushes.Transparent; }
    }

    // ── 색 선택 Picker ───────────────────────────────────────────────

    private void OnBorderColorPickClick(object sender, RoutedEventArgs e)
        => PickColor(TxtBorderColor);

    private void OnBackgroundColorPickClick(object sender, RoutedEventArgs e)
        => PickColor(TxtBackgroundColor);

    private void PickColor(System.Windows.Controls.TextBox target)
    {
        using var dlg = new System.Windows.Forms.ColorDialog { FullOpen = true, AnyColor = true };
        if (TryParseWpfColor(target.Text, out var current))
            dlg.Color = System.Drawing.Color.FromArgb(current.A, current.R, current.G, current.B);
        if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            var p = dlg.Color;
            target.Text = $"#{p.R:X2}{p.G:X2}{p.B:X2}";
        }
    }

    private static bool TryParseWpfColor(string? hex, out WpfMedia.Color color)
    {
        color = default;
        if (string.IsNullOrWhiteSpace(hex)) return false;
        var s = hex.Trim();
        if (!s.StartsWith('#')) s = '#' + s;
        try { color = (WpfMedia.Color)WpfMedia.ColorConverter.ConvertFromString(s)!; return true; }
        catch { return false; }
    }

    // ── 확인 / 취소 ─────────────────────────────────────────────────

    private void OnOk(object sender, RoutedEventArgs e)
    {
        ResultBorderColor = string.IsNullOrWhiteSpace(TxtBorderColor.Text)
            ? null : TxtBorderColor.Text.Trim();

        ResultBorderThicknessPt = double.TryParse(TxtBorderThickness.Text, out var pt) && pt >= 0
            ? pt : 1.0;

        ResultBackgroundColor = string.IsNullOrWhiteSpace(TxtBackgroundColor.Text)
            ? null : TxtBackgroundColor.Text.Trim();

        ResultPaddingTopMm    = ParseMm(TxtPaddingTop.Text);
        ResultPaddingBottomMm = ParseMm(TxtPaddingBottom.Text);
        ResultPaddingLeftMm   = ParseMm(TxtPaddingLeft.Text);
        ResultPaddingRightMm  = ParseMm(TxtPaddingRight.Text);

        ResultHAlign = (TextBoxHAlign)Math.Clamp(CboHAlign.SelectedIndex, 0, 3);
        ResultVAlign = (TextBoxVAlign)Math.Clamp(CboVAlign.SelectedIndex, 0, 2);

        DialogResult = true;
    }

    private void OnCancel(object sender, RoutedEventArgs e) => DialogResult = false;

    private static double ParseMm(string? s)
        => double.TryParse(s?.Trim(), out var v) && v >= 0 ? v : 2.0;
}

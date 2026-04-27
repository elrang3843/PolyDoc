using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using PolyDonky.Core;
using WpfMedia = System.Windows.Media;

namespace PolyDonky.App.Views;

/// <summary>글상자 테두리·배경·여백·정렬 + 모양별 형태 파라미터 편집 대화상자.</summary>
public partial class TextBoxPropertiesWindow : Window
{
    public TextBoxShape  ResultShape             { get; private set; }
    public string?       ResultBorderColor       { get; private set; }
    public double        ResultBorderThicknessPt { get; private set; }
    public string?       ResultBackgroundColor   { get; private set; }
    public double        ResultPaddingTopMm      { get; private set; }
    public double        ResultPaddingBottomMm   { get; private set; }
    public double        ResultPaddingLeftMm     { get; private set; }
    public double        ResultPaddingRightMm    { get; private set; }
    public TextBoxHAlign ResultHAlign            { get; private set; }
    public TextBoxVAlign ResultVAlign            { get; private set; }

    public SpeechPointerDirection ResultSpeechDirection    { get; private set; }
    public int                    ResultCloudPuffCount     { get; private set; }
    public int                    ResultSpikeCount         { get; private set; }
    public int                    ResultLightningBendCount { get; private set; }
    public double                 ResultPieStartAngleDeg   { get; private set; }
    public double                 ResultPieSweepAngleDeg   { get; private set; }
    public double                 ResultRotationAngleDeg   { get; private set; }
    public TextOrientation        ResultTextOrientation    { get; private set; }
    public TextProgression        ResultTextProgression    { get; private set; }

    public TextBoxPropertiesWindow(TextBoxObject model)
    {
        InitializeComponent();

        CboShape.SelectedIndex  = (int)model.Shape;
        TxtBorderColor.Text     = model.BorderColor ?? "#000000";
        TxtBorderThickness.Text = model.BorderThicknessPt.ToString("0.##");
        TxtBackgroundColor.Text = model.BackgroundColor ?? "";

        TxtPaddingTop.Text    = model.PaddingTopMm.ToString("0.##");
        TxtPaddingBottom.Text = model.PaddingBottomMm.ToString("0.##");
        TxtPaddingLeft.Text   = model.PaddingLeftMm.ToString("0.##");
        TxtPaddingRight.Text  = model.PaddingRightMm.ToString("0.##");

        CboHAlign.SelectedIndex = (int)model.HAlign;
        CboVAlign.SelectedIndex = (int)model.VAlign;

        // 모양별 파라미터
        SelectSpeechDirection(model.SpeechDirection);
        TxtCloudPuffs.Text     = model.CloudPuffCount.ToString();
        TxtSpikeCount.Text     = model.SpikeCount.ToString();
        TxtLightningBends.Text = model.LightningBendCount.ToString();
        TxtPieStartAngle.Text  = model.PieStartAngleDeg.ToString("0.##");
        TxtPieSweepAngle.Text  = model.PieSweepAngleDeg.ToString("0.##");
        TxtRotationAngle.Text  = model.RotationAngleDeg.ToString("0.##");

        CboTextOrientation.SelectedIndex = (int)model.TextOrientation;
        CboTextProgression.SelectedIndex = (int)model.TextProgression;

        UpdateShapePanelVisibility();

        RefreshColorButton(BtnBorderColorPick,     TxtBorderColor.Text);
        RefreshColorButton(BtnBackgroundColorPick, TxtBackgroundColor.Text);
    }

    // ── 모양 변경 → 해당 파라미터 패널만 표시 ──────────────────────

    private void OnShapeChanged(object sender, SelectionChangedEventArgs e)
        => UpdateShapePanelVisibility();

    private void UpdateShapePanelVisibility()
    {
        var shape = (TextBoxShape)System.Math.Clamp(CboShape.SelectedIndex, 0, 6);
        PnlSpeech.Visibility    = shape == TextBoxShape.Speech    ? Visibility.Visible : Visibility.Collapsed;
        PnlCloud.Visibility     = shape == TextBoxShape.Cloud     ? Visibility.Visible : Visibility.Collapsed;
        PnlSpiky.Visibility     = shape == TextBoxShape.Spiky     ? Visibility.Visible : Visibility.Collapsed;
        PnlLightning.Visibility = shape == TextBoxShape.Lightning ? Visibility.Visible : Visibility.Collapsed;
        PnlPie.Visibility       = shape == TextBoxShape.Pie       ? Visibility.Visible : Visibility.Collapsed;
    }

    private void SelectSpeechDirection(SpeechPointerDirection dir)
    {
        var rb = dir switch
        {
            SpeechPointerDirection.TopLeft     => RbSpeechTL,
            SpeechPointerDirection.Top         => RbSpeechT,
            SpeechPointerDirection.TopRight    => RbSpeechTR,
            SpeechPointerDirection.Left        => RbSpeechL,
            SpeechPointerDirection.Right       => RbSpeechR,
            SpeechPointerDirection.BottomLeft  => RbSpeechBL,
            SpeechPointerDirection.Bottom      => RbSpeechB,
            SpeechPointerDirection.BottomRight => RbSpeechBR,
            _                                  => RbSpeechB,
        };
        rb.IsChecked = true;
    }

    private SpeechPointerDirection GetSelectedSpeechDirection()
    {
        if (RbSpeechTL.IsChecked == true) return SpeechPointerDirection.TopLeft;
        if (RbSpeechT .IsChecked == true) return SpeechPointerDirection.Top;
        if (RbSpeechTR.IsChecked == true) return SpeechPointerDirection.TopRight;
        if (RbSpeechL .IsChecked == true) return SpeechPointerDirection.Left;
        if (RbSpeechR .IsChecked == true) return SpeechPointerDirection.Right;
        if (RbSpeechBL.IsChecked == true) return SpeechPointerDirection.BottomLeft;
        if (RbSpeechBR.IsChecked == true) return SpeechPointerDirection.BottomRight;
        return SpeechPointerDirection.Bottom;
    }

    // ── 색 텍스트 변경 → 미리보기 버튼 갱신 ────────────────────────

    private void OnBorderColorChanged(object sender, TextChangedEventArgs e)
        => RefreshColorButton(BtnBorderColorPick, TxtBorderColor.Text);

    private void OnBackgroundColorChanged(object sender, TextChangedEventArgs e)
        => RefreshColorButton(BtnBackgroundColorPick, TxtBackgroundColor.Text);

    private static void RefreshColorButton(Button btn, string? hex)
    {
        if (string.IsNullOrWhiteSpace(hex))
        {
            btn.Background = Brushes.Transparent;
            return;
        }
        try
        {
            var s = hex.Trim();
            if (!s.StartsWith('#')) s = '#' + s;
            var c = (WpfMedia.Color)WpfMedia.ColorConverter.ConvertFromString(s)!;
            btn.Background = new SolidColorBrush(c);
        }
        catch { btn.Background = Brushes.Transparent; }
    }

    // ── 색 선택 Picker ───────────────────────────────────────────────

    private void OnBorderColorPickClick(object sender, RoutedEventArgs e)
        => PickColor(TxtBorderColor);

    private void OnBackgroundColorPickClick(object sender, RoutedEventArgs e)
        => PickColor(TxtBackgroundColor);

    private void PickColor(TextBox target)
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
        ResultShape = (TextBoxShape)System.Math.Clamp(CboShape.SelectedIndex, 0, 6);

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

        ResultHAlign = (TextBoxHAlign)System.Math.Clamp(CboHAlign.SelectedIndex, 0, 3);
        ResultVAlign = (TextBoxVAlign)System.Math.Clamp(CboVAlign.SelectedIndex, 0, 2);

        ResultSpeechDirection    = GetSelectedSpeechDirection();
        ResultCloudPuffCount     = ParseInt(TxtCloudPuffs.Text,     10, 6, 32);
        ResultSpikeCount         = ParseInt(TxtSpikeCount.Text,     12, 5, 24);
        ResultLightningBendCount = ParseInt(TxtLightningBends.Text,  2, 1,  5);
        ResultPieStartAngleDeg   = ParseAngle(TxtPieStartAngle.Text);
        ResultPieSweepAngleDeg   = ParsePieSweep(TxtPieSweepAngle.Text);
        ResultRotationAngleDeg   = ParseAngle(TxtRotationAngle.Text);
        ResultTextOrientation    = (TextOrientation)System.Math.Clamp(CboTextOrientation.SelectedIndex, 0, 1);
        ResultTextProgression    = (TextProgression)System.Math.Clamp(CboTextProgression.SelectedIndex, 0, 1);

        DialogResult = true;
    }

    private static double ParseAngle(string? s)
        => double.TryParse(s?.Trim(), out var v) ? System.Math.Clamp(v, -360, 360) : 0;

    private static double ParsePieSweep(string? s)
        => double.TryParse(s?.Trim(), out var v) ? System.Math.Clamp(v, 5, 355) : 270;

    private void OnCancel(object sender, RoutedEventArgs e) => DialogResult = false;

    private static double ParseMm(string? s)
        => double.TryParse(s?.Trim(), out var v) && v >= 0 ? v : 2.0;

    private static int ParseInt(string? s, int fallback, int min, int max)
        => int.TryParse(s?.Trim(), out var v) ? System.Math.Clamp(v, min, max) : fallback;
}

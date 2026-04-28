using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using PolyDonky.Core;

namespace PolyDonky.App.Views;

public partial class ShapePropertiesWindow : Window
{
    private readonly ShapeObject _shape;

    public ShapePropertiesWindow(ShapeObject shape)
    {
        _shape = shape;
        InitializeComponent();
        Loaded += (_, _) => Populate();
    }

    private void Populate()
    {
        // ── 배치 ──────────────────────────────────────────────────────────
        SelectComboByTag(CboWrapMode, _shape.WrapMode.ToString());
        SelectComboByTag(CboHAlign,   _shape.HAlign.ToString());

        // ── 크기 ──────────────────────────────────────────────────────────
        TxtWidth.Text  = _shape.WidthMm.ToString("0.##");
        TxtHeight.Text = _shape.HeightMm.ToString("0.##");

        // ── 오버레이 위치 ─────────────────────────────────────────────────
        TxtOverlayX.Text = _shape.OverlayXMm.ToString("0.##");
        TxtOverlayY.Text = _shape.OverlayYMm.ToString("0.##");

        // ── 선 ────────────────────────────────────────────────────────────
        StrokeColorPicker.ColorText = _shape.StrokeColor ?? "#000000";
        TxtStrokePt.Text            = _shape.StrokeThicknessPt.ToString("0.##");
        SelectComboByTag(CboStrokeDash, _shape.StrokeDash.ToString());

        // ── 채우기 ────────────────────────────────────────────────────────
        FillColorPicker.ColorText = _shape.FillColor ?? string.Empty;
        TxtFillOpacity.Text       = _shape.FillOpacity.ToString("0.##");

        // ── 화살촉 ────────────────────────────────────────────────────────
        SelectComboByTag(CboStartArrow, _shape.StartArrow.ToString());
        SelectComboByTag(CboEndArrow,   _shape.EndArrow.ToString());
        bool isLineKind = _shape.Kind is ShapeKind.Line or ShapeKind.Polyline or ShapeKind.Spline;
        GrpArrows.IsEnabled = isLineKind;

        // ── 모양 파라미터 ─────────────────────────────────────────────────
        TxtSideCount.Text  = _shape.SideCount.ToString();
        TxtInnerRadius.Text = _shape.InnerRadiusRatio.ToString("0.##");
        TxtCornerRadius.Text = _shape.CornerRadiusMm.ToString("0.##");
        TxtRotation.Text   = _shape.RotationAngleDeg.ToString("0.##");

        bool hasSides      = _shape.Kind is ShapeKind.RegularPolygon or ShapeKind.Star;
        bool hasInner      = _shape.Kind == ShapeKind.Star;
        bool hasCornerRad  = _shape.Kind == ShapeKind.RoundedRect;
        LblSideCount.IsEnabled   = hasSides;
        TxtSideCount.IsEnabled   = hasSides;
        LblInnerRadius.IsEnabled = hasInner;
        TxtInnerRadius.IsEnabled = hasInner;
        LblCornerRadius.IsEnabled = hasCornerRad;
        TxtCornerRadius.IsEnabled = hasCornerRad;

        // ── 레이블 ────────────────────────────────────────────────────────
        TxtLabelText.Text             = _shape.LabelText  ?? string.Empty;
        TxtLabelFont.Text             = _shape.LabelFontFamily ?? string.Empty;
        TxtLabelSize.Text             = _shape.LabelFontSizePt.ToString("0.##");
        LabelColorPicker.ColorText    = _shape.LabelColor ?? string.Empty;
        ChkLabelBold.IsChecked   = _shape.LabelBold;
        ChkLabelItalic.IsChecked = _shape.LabelItalic;
    }

    private void OnOk(object sender, RoutedEventArgs e)
    {
        if (!Apply()) return;
        DialogResult = true;
    }

    private bool Apply()
    {
        // ── 배치 ──────────────────────────────────────────────────────────
        if (GetComboTag(CboWrapMode) is string wrapStr &&
            Enum.TryParse<ImageWrapMode>(wrapStr, out var wrap))
            _shape.WrapMode = wrap;

        if (GetComboTag(CboHAlign) is string hAlignStr &&
            Enum.TryParse<ImageHAlign>(hAlignStr, out var hAlign))
            _shape.HAlign = hAlign;

        // ── 크기 ──────────────────────────────────────────────────────────
        if (double.TryParse(TxtWidth.Text,  out double w) && w > 0) _shape.WidthMm  = w;
        if (double.TryParse(TxtHeight.Text, out double h) && h > 0) _shape.HeightMm = h;

        // ── 오버레이 위치 ─────────────────────────────────────────────────
        if (double.TryParse(TxtOverlayX.Text, out double ox)) _shape.OverlayXMm = ox;
        if (double.TryParse(TxtOverlayY.Text, out double oy)) _shape.OverlayYMm = oy;

        // ── 선 ────────────────────────────────────────────────────────────
        _shape.StrokeColor = string.IsNullOrWhiteSpace(StrokeColorPicker.ColorText) ? "#000000" : StrokeColorPicker.ColorText.Trim();
        if (double.TryParse(TxtStrokePt.Text, out double sp) && sp >= 0) _shape.StrokeThicknessPt = sp;
        if (GetComboTag(CboStrokeDash) is string dashStr &&
            Enum.TryParse<StrokeDash>(dashStr, out var dash))
            _shape.StrokeDash = dash;

        // ── 채우기 ────────────────────────────────────────────────────────
        _shape.FillColor = string.IsNullOrWhiteSpace(FillColorPicker.ColorText) ? null : FillColorPicker.ColorText.Trim();
        if (double.TryParse(TxtFillOpacity.Text, out double fo))
            _shape.FillOpacity = Math.Clamp(fo, 0, 1);

        // ── 화살촉 ────────────────────────────────────────────────────────
        if (GetComboTag(CboStartArrow) is string saStr && Enum.TryParse<ShapeArrow>(saStr, out var sa))
            _shape.StartArrow = sa;
        if (GetComboTag(CboEndArrow) is string eaStr && Enum.TryParse<ShapeArrow>(eaStr, out var ea))
            _shape.EndArrow = ea;

        // ── 모양 파라미터 ─────────────────────────────────────────────────
        if (int.TryParse(TxtSideCount.Text, out int sc) && sc >= 3) _shape.SideCount = sc;
        if (double.TryParse(TxtInnerRadius.Text, out double ir))
            _shape.InnerRadiusRatio = Math.Clamp(ir, 0.1, 0.9);
        if (double.TryParse(TxtCornerRadius.Text, out double cr) && cr >= 0) _shape.CornerRadiusMm = cr;
        if (double.TryParse(TxtRotation.Text, out double rot)) _shape.RotationAngleDeg = rot;

        // ── 레이블 ────────────────────────────────────────────────────────
        _shape.LabelText       = string.IsNullOrWhiteSpace(TxtLabelText.Text) ? null : TxtLabelText.Text;
        _shape.LabelFontFamily = string.IsNullOrWhiteSpace(TxtLabelFont.Text) ? null : TxtLabelFont.Text.Trim();
        if (double.TryParse(TxtLabelSize.Text, out double ls) && ls > 0) _shape.LabelFontSizePt = ls;
        _shape.LabelColor  = string.IsNullOrWhiteSpace(LabelColorPicker.ColorText) ? null : LabelColorPicker.ColorText.Trim();
        _shape.LabelBold   = ChkLabelBold.IsChecked   == true;
        _shape.LabelItalic = ChkLabelItalic.IsChecked == true;

        _shape.Status = NodeStatus.Modified;
        return true;
    }

    private static void SelectComboByTag(ComboBox cbo, string? tag)
    {
        if (tag is null) return;
        foreach (ComboBoxItem item in cbo.Items)
        {
            if (item.Tag?.ToString() == tag)
            {
                cbo.SelectedItem = item;
                return;
            }
        }
        if (cbo.Items.Count > 0) cbo.SelectedIndex = 0;
    }

    private static string? GetComboTag(ComboBox cbo)
        => (cbo.SelectedItem as ComboBoxItem)?.Tag?.ToString();
}

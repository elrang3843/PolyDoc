using System.Windows;
using PolyDonky.App.Controls;
using PolyDonky.Core;

namespace PolyDonky.App.Views;

/// <summary>표 속성 다이얼로그. 표 정렬·배경·여백·테두리를 편집한다.</summary>
public partial class TablePropertiesWindow : Window
{
    private readonly Table _table;

    public TablePropertiesWindow(Table table)
    {
        InitializeComponent();
        _table = table;
        LoadValues();
    }

    // ── 초기화 ───────────────────────────────────────────────────────────

    private void LoadValues()
    {
        switch (_table.HAlign)
        {
            case TableHAlign.Center: AlignCenterRadio.IsChecked = true; break;
            case TableHAlign.Right:  AlignRightRadio.IsChecked  = true; break;
            default:                 AlignLeftRadio.IsChecked   = true; break;
        }

        BgColorPicker.ColorText = _table.BackgroundColor ?? string.Empty;

        CellPadTopBox.Text    = _table.DefaultCellPaddingTopMm    > 0 ? _table.DefaultCellPaddingTopMm.ToString("F1")    : "1.0";
        CellPadBottomBox.Text = _table.DefaultCellPaddingBottomMm > 0 ? _table.DefaultCellPaddingBottomMm.ToString("F1") : "1.0";
        CellPadLeftBox.Text   = _table.DefaultCellPaddingLeftMm   > 0 ? _table.DefaultCellPaddingLeftMm.ToString("F1")   : "1.5";
        CellPadRightBox.Text  = _table.DefaultCellPaddingRightMm  > 0 ? _table.DefaultCellPaddingRightMm.ToString("F1")  : "1.5";

        OuterMarginTopBox.Text    = _table.OuterMarginTopMm    > 0 ? _table.OuterMarginTopMm.ToString("F1")    : "0";
        OuterMarginBottomBox.Text = _table.OuterMarginBottomMm > 0 ? _table.OuterMarginBottomMm.ToString("F1") : "0";
        OuterMarginLeftBox.Text   = _table.OuterMarginLeftMm   > 0 ? _table.OuterMarginLeftMm.ToString("F1")   : "0";
        OuterMarginRightBox.Text  = _table.OuterMarginRightMm  > 0 ? _table.OuterMarginRightMm.ToString("F1")  : "0";

        BorderThicknessBox.Text = _table.BorderThicknessPt > 0 ? _table.BorderThicknessPt.ToString("F2") : "0";
        BorderColorPicker.ColorText = _table.BorderColor ?? string.Empty;
    }

    // ── 확인 ─────────────────────────────────────────────────────────────

    private void OnOkClick(object sender, RoutedEventArgs e)
    {
        if (!TryParseNonNeg(CellPadTopBox.Text,    out double cpt) ||
            !TryParseNonNeg(CellPadBottomBox.Text, out double cpb) ||
            !TryParseNonNeg(CellPadLeftBox.Text,   out double cpl) ||
            !TryParseNonNeg(CellPadRightBox.Text,  out double cpr))
        {
            MessageBox.Show(this, "기본 셀 안여백은 0 이상의 숫자(mm)로 입력하세요.", "표 속성",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (!TryParseNonNeg(OuterMarginTopBox.Text,    out double omt) ||
            !TryParseNonNeg(OuterMarginBottomBox.Text, out double omb) ||
            !TryParseNonNeg(OuterMarginLeftBox.Text,   out double oml) ||
            !TryParseNonNeg(OuterMarginRightBox.Text,  out double omr))
        {
            MessageBox.Show(this, "바깥 여백은 0 이상의 숫자(mm)로 입력하세요.", "표 속성",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (!double.TryParse(BorderThicknessBox.Text.Trim(), out double borderPt) || borderPt < 0)
        {
            MessageBox.Show(this, "외곽선 두께는 0 이상의 숫자(pt)로 입력하세요.", "표 속성",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            BorderThicknessBox.Focus();
            return;
        }

        string bgColor = BgColorPicker.ColorText.Trim();
        if (bgColor.Length > 0 && !ColorPickerBox.TryParseColor(bgColor, out _))
        {
            MessageBox.Show(this, "배경색을 올바른 색상 값으로 입력하세요 (예: #FFCC00).", "표 속성",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            BgColorPicker.Focus();
            return;
        }

        string borderColor = BorderColorPicker.ColorText.Trim();
        if (borderColor.Length > 0 && !ColorPickerBox.TryParseColor(borderColor, out _))
        {
            MessageBox.Show(this, "외곽선 색상을 올바른 색상 값으로 입력하세요 (예: #C8C8C8).", "표 속성",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            BorderColorPicker.Focus();
            return;
        }

        _table.HAlign = AlignCenterRadio.IsChecked == true ? TableHAlign.Center
                      : AlignRightRadio.IsChecked  == true ? TableHAlign.Right
                      : TableHAlign.Left;

        _table.BackgroundColor = bgColor.Length > 0 ? bgColor : null;

        _table.DefaultCellPaddingTopMm    = cpt;
        _table.DefaultCellPaddingBottomMm = cpb;
        _table.DefaultCellPaddingLeftMm   = cpl;
        _table.DefaultCellPaddingRightMm  = cpr;

        _table.OuterMarginTopMm    = omt;
        _table.OuterMarginBottomMm = omb;
        _table.OuterMarginLeftMm   = oml;
        _table.OuterMarginRightMm  = omr;

        _table.BorderThicknessPt = borderPt;
        _table.BorderColor       = borderColor.Length > 0 ? borderColor : null;

        DialogResult = true;
        Close();
    }

    private void OnCancelClick(object sender, RoutedEventArgs e) => Close();

    // ── 유틸 ─────────────────────────────────────────────────────────────

    private static bool TryParseNonNeg(string text, out double value)
        => double.TryParse(text.Trim(), out value) && value >= 0;
}

using System.Windows;
using PolyDonky.App.Controls;
using PolyDonky.Core;

namespace PolyDonky.App.Views;

/// <summary>셀 속성 다이얼로그. 배경색·텍스트 정렬·여백·테두리를 편집한다.</summary>
public partial class CellPropertiesWindow : Window
{
    private readonly TableCell _cell;

    public CellPropertiesWindow(TableCell cell)
    {
        InitializeComponent();
        _cell = cell;
        LoadValues();
    }

    // ── 초기화 ───────────────────────────────────────────────────────────

    private void LoadValues()
    {
        // 면별 편집 컨트롤 라벨.
        SideTop.Label    = "위";
        SideBottom.Label = "아래";
        SideLeft.Label   = "왼쪽";
        SideRight.Label  = "오른쪽";

        BgColorPicker.ColorText     = _cell.BackgroundColor ?? string.Empty;
        BorderColorPicker.ColorText = _cell.BorderColor ?? "#C8C8C8";

        switch (_cell.TextAlign)
        {
            case CellTextAlign.Center:  AlignCenterRadio.IsChecked  = true; break;
            case CellTextAlign.Right:   AlignRightRadio.IsChecked   = true; break;
            case CellTextAlign.Justify: AlignJustifyRadio.IsChecked = true; break;
            default:                    AlignLeftRadio.IsChecked    = true; break;
        }

        PadTopBox.Text    = _cell.PaddingTopMm    > 0 ? _cell.PaddingTopMm.ToString("F1")    : "0";
        PadBottomBox.Text = _cell.PaddingBottomMm > 0 ? _cell.PaddingBottomMm.ToString("F1") : "0";
        PadLeftBox.Text   = _cell.PaddingLeftMm   > 0 ? _cell.PaddingLeftMm.ToString("F1")   : "0";
        PadRightBox.Text  = _cell.PaddingRightMm  > 0 ? _cell.PaddingRightMm.ToString("F1")  : "0";

        BorderThicknessBox.Text = _cell.BorderThicknessPt > 0
            ? _cell.BorderThicknessPt.ToString("F2") : "0.75";

        // 면별 테두리.
        SideTop.Value    = _cell.BorderTop;
        SideBottom.Value = _cell.BorderBottom;
        SideLeft.Value   = _cell.BorderLeft;
        SideRight.Value  = _cell.BorderRight;

        bool anyPerSide = _cell.BorderTop is not null || _cell.BorderBottom is not null
                       || _cell.BorderLeft is not null || _cell.BorderRight is not null;
        PerSideToggle.IsChecked = anyPerSide;
        PerSidePanel.Visibility = anyPerSide ? Visibility.Visible : Visibility.Collapsed;
        PerSideToggle.Content   = anyPerSide ? "면별 설정 접기 ▴" : "면별 설정 펼치기 ▾";
    }

    private void OnPerSideToggleClick(object sender, RoutedEventArgs e)
    {
        bool open = PerSideToggle.IsChecked == true;
        PerSidePanel.Visibility = open ? Visibility.Visible : Visibility.Collapsed;
        PerSideToggle.Content   = open ? "면별 설정 접기 ▴" : "면별 설정 펼치기 ▾";
    }

    // ── 확인 ─────────────────────────────────────────────────────────────

    private void OnOkClick(object sender, RoutedEventArgs e)
    {
        if (!TryParseNonNegMm(PadTopBox.Text,    out double padTop)    ||
            !TryParseNonNegMm(PadBottomBox.Text, out double padBottom) ||
            !TryParseNonNegMm(PadLeftBox.Text,   out double padLeft)   ||
            !TryParseNonNegMm(PadRightBox.Text,  out double padRight))
        {
            MessageBox.Show(this, "여백은 0 이상의 숫자(mm)로 입력하세요.", "셀 속성",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (!double.TryParse(BorderThicknessBox.Text.Trim(), out double borderPt) || borderPt < 0)
        {
            MessageBox.Show(this, "테두리 두께는 0 이상의 숫자(pt)로 입력하세요.", "셀 속성",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            BorderThicknessBox.Focus();
            return;
        }

        string bgColor = BgColorPicker.ColorText.Trim();
        if (bgColor.Length > 0 && !ColorPickerBox.TryParseColor(bgColor, out _))
        {
            MessageBox.Show(this, "배경색을 올바른 색상 값으로 입력하세요 (예: #FFCC00).", "셀 속성",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            BgColorPicker.Focus();
            return;
        }

        string borderColor = BorderColorPicker.ColorText.Trim();
        if (borderColor.Length > 0 && !ColorPickerBox.TryParseColor(borderColor, out _))
        {
            MessageBox.Show(this, "테두리 색상을 올바른 색상 값으로 입력하세요 (예: #C8C8C8).", "셀 속성",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            BorderColorPicker.Focus();
            return;
        }

        _cell.TextAlign = AlignCenterRadio.IsChecked  == true ? CellTextAlign.Center
                        : AlignRightRadio.IsChecked   == true ? CellTextAlign.Right
                        : AlignJustifyRadio.IsChecked == true ? CellTextAlign.Justify
                        : CellTextAlign.Left;

        _cell.PaddingTopMm    = padTop;
        _cell.PaddingBottomMm = padBottom;
        _cell.PaddingLeftMm   = padLeft;
        _cell.PaddingRightMm  = padRight;

        _cell.BorderThicknessPt = borderPt;
        _cell.BorderColor       = borderColor.Length > 0 ? borderColor : null;
        _cell.BackgroundColor   = bgColor.Length > 0 ? bgColor : null;

        // 면별 테두리 — 유효성 검사 후 적용.
        foreach (var editor in new[] { SideTop, SideBottom, SideLeft, SideRight })
        {
            if (!editor.TryValidate(out string? sideError))
            {
                MessageBox.Show(this, sideError ?? "면별 테두리 입력이 잘못되었습니다.", "셀 속성",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
        }
        _cell.BorderTop    = SideTop.Value;
        _cell.BorderBottom = SideBottom.Value;
        _cell.BorderLeft   = SideLeft.Value;
        _cell.BorderRight  = SideRight.Value;

        DialogResult = true;
        Close();
    }

    private void OnCancelClick(object sender, RoutedEventArgs e) => Close();

    // ── 유틸 ─────────────────────────────────────────────────────────────

    private static bool TryParseNonNegMm(string text, out double value)
        => double.TryParse(text.Trim(), out value) && value >= 0;
}

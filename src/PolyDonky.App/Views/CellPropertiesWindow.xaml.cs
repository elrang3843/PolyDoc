using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
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
        // 배경색
        BgColorBox.Text = _cell.BackgroundColor ?? string.Empty;

        // 텍스트 정렬
        switch (_cell.TextAlign)
        {
            case CellTextAlign.Center:  AlignCenterRadio.IsChecked  = true; break;
            case CellTextAlign.Right:   AlignRightRadio.IsChecked   = true; break;
            case CellTextAlign.Justify: AlignJustifyRadio.IsChecked = true; break;
            default:                    AlignLeftRadio.IsChecked    = true; break;
        }

        // 여백 (0 이하면 렌더러 기본값을 사용 — 빈 문자열로 표시해 사용자가 직접 입력하게 한다)
        PadTopBox.Text    = _cell.PaddingTopMm    > 0 ? _cell.PaddingTopMm.ToString("F1")    : "0";
        PadBottomBox.Text = _cell.PaddingBottomMm > 0 ? _cell.PaddingBottomMm.ToString("F1") : "0";
        PadLeftBox.Text   = _cell.PaddingLeftMm   > 0 ? _cell.PaddingLeftMm.ToString("F1")   : "0";
        PadRightBox.Text  = _cell.PaddingRightMm  > 0 ? _cell.PaddingRightMm.ToString("F1")  : "0";

        // 테두리
        BorderThicknessBox.Text = _cell.BorderThicknessPt > 0
            ? _cell.BorderThicknessPt.ToString("F2") : "0.75";
        BorderColorBox.Text = _cell.BorderColor ?? "#C8C8C8";
    }

    // ── 색상 스와치 실시간 갱신 ───────────────────────────────────────────

    private void OnBgColorChanged(object sender, TextChangedEventArgs e)
        => BgColorSwatch.Background = TryParseBrush(BgColorBox.Text);

    private void OnBorderColorChanged(object sender, TextChangedEventArgs e)
        => BorderColorSwatch.Background = TryParseBrush(BorderColorBox.Text);

    private void OnBgColorClear(object sender, RoutedEventArgs e)
    {
        BgColorBox.Text = string.Empty;
        BgColorSwatch.Background = null;
    }

    // ── 확인 ─────────────────────────────────────────────────────────────

    private void OnOkClick(object sender, RoutedEventArgs e)
    {
        // 여백 검증
        if (!TryParseNonNegMm(PadTopBox.Text,    out double padTop)    ||
            !TryParseNonNegMm(PadBottomBox.Text, out double padBottom) ||
            !TryParseNonNegMm(PadLeftBox.Text,   out double padLeft)   ||
            !TryParseNonNegMm(PadRightBox.Text,  out double padRight))
        {
            MessageBox.Show(this, "여백은 0 이상의 숫자(mm)로 입력하세요.", "셀 속성",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        // 테두리 두께 검증
        if (!double.TryParse(BorderThicknessBox.Text.Trim(), out double borderPt) || borderPt < 0)
        {
            MessageBox.Show(this, "테두리 두께는 0 이상의 숫자(pt)로 입력하세요.", "셀 속성",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            BorderThicknessBox.Focus();
            return;
        }

        // 배경색 검증 (빈 문자열 = 없음은 허용)
        string bgColor = BgColorBox.Text.Trim();
        if (bgColor.Length > 0 && TryParseBrush(bgColor) is null)
        {
            MessageBox.Show(this, "배경색을 올바른 색상 값으로 입력하세요 (예: #FFCC00).", "셀 속성",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            BgColorBox.Focus();
            return;
        }

        // 테두리 색상 검증
        string borderColor = BorderColorBox.Text.Trim();
        if (borderColor.Length > 0 && TryParseBrush(borderColor) is null)
        {
            MessageBox.Show(this, "테두리 색상을 올바른 색상 값으로 입력하세요 (예: #C8C8C8).", "셀 속성",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            BorderColorBox.Focus();
            return;
        }

        // 모델에 반영
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

        DialogResult = true;
        Close();
    }

    private void OnCancelClick(object sender, RoutedEventArgs e) => Close();

    // ── 유틸 ─────────────────────────────────────────────────────────────

    private static SolidColorBrush? TryParseBrush(string? hex)
    {
        if (string.IsNullOrWhiteSpace(hex)) return null;
        try
        {
            var c = (System.Windows.Media.Color)ColorConverter.ConvertFromString(hex)!;
            return new SolidColorBrush(c);
        }
        catch { return null; }
    }

    private static bool TryParseNonNegMm(string text, out double value)
        => double.TryParse(text.Trim(), out value) && value >= 0;
}

using System.Windows;
using PolyDonky.App.Controls;
using PolyDonky.Core;

namespace PolyDonky.App.Views;

/// <summary>표 속성 다이얼로그. 배치 방식·정렬·배경·여백·테두리를 편집한다.</summary>
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
        // 면별 편집 컨트롤 라벨 — 디자이너에서 설정하지 않고 코드비하인드에서 한국어 라벨 부여.
        SideTop.Label    = "위";
        SideBottom.Label = "아래";
        SideLeft.Label   = "왼쪽";
        SideRight.Label  = "오른쪽";
        SideInnerH.Label = "중앙 가로";
        SideInnerV.Label = "중앙 세로";

        // 배치 모드
        switch (_table.WrapMode)
        {
            case TableWrapMode.InFrontOfText: WrapFrontRadio.IsChecked  = true; break;
            case TableWrapMode.BehindText:    WrapBehindRadio.IsChecked = true; break;
            case TableWrapMode.Fixed:         WrapFixedRadio.IsChecked  = true; break;
            default:                          WrapBlockRadio.IsChecked  = true; break;
        }
        UpdateOverlayVisibility();

        OverlayXBox.Text    = _table.OverlayXMm.ToString("F1");
        OverlayYBox.Text    = _table.OverlayYMm.ToString("F1");
        AnchorPageBox.Text  = (_table.AnchorPageIndex + 1).ToString();

        // 표 정렬
        switch (_table.HAlign)
        {
            case TableHAlign.Center: AlignCenterRadio.IsChecked = true; break;
            case TableHAlign.Right:  AlignRightRadio.IsChecked  = true; break;
            default:                 AlignLeftRadio.IsChecked   = true; break;
        }

        // 표 크기
        TableWidthBox.Text  = _table.WidthMm  > 0 ? _table.WidthMm.ToString("F1")  : "0";
        TableHeightBox.Text = _table.HeightMm > 0 ? _table.HeightMm.ToString("F1") : "0";

        BgColorPicker.ColorText = _table.BackgroundColor ?? string.Empty;

        CellPadTopBox.Text    = _table.DefaultCellPaddingTopMm    > 0 ? _table.DefaultCellPaddingTopMm.ToString("F1")    : Table.FallbackCellPaddingVerticalMm.ToString("F1");
        CellPadBottomBox.Text = _table.DefaultCellPaddingBottomMm > 0 ? _table.DefaultCellPaddingBottomMm.ToString("F1") : Table.FallbackCellPaddingVerticalMm.ToString("F1");
        CellPadLeftBox.Text   = _table.DefaultCellPaddingLeftMm   > 0 ? _table.DefaultCellPaddingLeftMm.ToString("F1")   : Table.FallbackCellPaddingHorizontalMm.ToString("F1");
        CellPadRightBox.Text  = _table.DefaultCellPaddingRightMm  > 0 ? _table.DefaultCellPaddingRightMm.ToString("F1")  : Table.FallbackCellPaddingHorizontalMm.ToString("F1");

        OuterMarginTopBox.Text    = _table.OuterMarginTopMm    > 0 ? _table.OuterMarginTopMm.ToString("F1")    : "0";
        OuterMarginBottomBox.Text = _table.OuterMarginBottomMm > 0 ? _table.OuterMarginBottomMm.ToString("F1") : "0";
        OuterMarginLeftBox.Text   = _table.OuterMarginLeftMm   > 0 ? _table.OuterMarginLeftMm.ToString("F1")   : "0";
        OuterMarginRightBox.Text  = _table.OuterMarginRightMm  > 0 ? _table.OuterMarginRightMm.ToString("F1")  : "0";

        BorderThicknessBox.Text     = _table.BorderThicknessPt > 0 ? _table.BorderThicknessPt.ToString("F2") : "0";
        BorderColorPicker.ColorText = _table.BorderColor ?? string.Empty;

        // 면별 외곽선
        SideTop.Value    = _table.BorderTop;
        SideBottom.Value = _table.BorderBottom;
        SideLeft.Value   = _table.BorderLeft;
        SideRight.Value  = _table.BorderRight;
        SideInnerH.Value = _table.InnerBorderHorizontal;
        SideInnerV.Value = _table.InnerBorderVertical;

        // 면별 설정이 하나라도 있으면 펼친 상태로 시작.
        bool anyPerSide = _table.BorderTop is not null || _table.BorderBottom is not null
                       || _table.BorderLeft is not null || _table.BorderRight is not null
                       || _table.InnerBorderHorizontal is not null
                       || _table.InnerBorderVertical is not null;
        PerSideToggle.IsChecked = anyPerSide;
        PerSidePanel.Visibility = anyPerSide ? Visibility.Visible : Visibility.Collapsed;
        PerSideToggle.Content   = anyPerSide ? "면별 설정 접기 ▴" : "면별 설정 펼치기 ▾";

        // 페이지 분할
        RepeatHeaderRowsCheck.IsChecked = _table.RepeatHeaderRowsOnBreak;
        HeaderColumnCountBox.Text       = _table.HeaderColumnCount.ToString();

        // 테두리 병합
        BorderCollapseCheck.IsChecked = _table.BorderCollapse;
    }

    private void OnPerSideToggleClick(object sender, RoutedEventArgs e)
    {
        bool open = PerSideToggle.IsChecked == true;
        PerSidePanel.Visibility = open ? Visibility.Visible : Visibility.Collapsed;
        PerSideToggle.Content   = open ? "면별 설정 접기 ▴" : "면별 설정 펼치기 ▾";
    }

    // ── 배치 모드 전환 ────────────────────────────────────────────────────

    private void OnWrapModeChanged(object sender, RoutedEventArgs e) => UpdateOverlayVisibility();

    private void UpdateOverlayVisibility()
    {
        bool isBlock = WrapBlockRadio.IsChecked == true;
        GrpOverlayPos.Visibility = isBlock ? Visibility.Collapsed : Visibility.Visible;
        GrpHAlign.Visibility     = isBlock ? Visibility.Visible   : Visibility.Collapsed;
    }

    // ── 확인 ─────────────────────────────────────────────────────────────

    private void OnOkClick(object sender, RoutedEventArgs e)
    {
        bool isOverlay = WrapBlockRadio.IsChecked != true;

        if (isOverlay)
        {
            if (!double.TryParse(OverlayXBox.Text.Trim(), out double ox) ||
                !double.TryParse(OverlayYBox.Text.Trim(), out double oy))
            {
                MessageBox.Show(this, "위치(X/Y)를 숫자(mm)로 입력하세요.", "표 속성",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (!int.TryParse(AnchorPageBox.Text.Trim(), out int anchorPage) || anchorPage < 1)
            {
                MessageBox.Show(this, "고정 페이지는 1 이상의 정수로 입력하세요.", "표 속성",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                AnchorPageBox.Focus();
                return;
            }
            _table.OverlayXMm      = ox;
            _table.OverlayYMm      = oy;
            _table.AnchorPageIndex = anchorPage - 1;
        }

        if (!TryParseNonNeg(TableWidthBox.Text,  out double tableW) ||
            !TryParseNonNeg(TableHeightBox.Text, out double tableH))
        {
            MessageBox.Show(this, "너비/높이는 0 이상의 숫자(mm)로 입력하세요.", "표 속성",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

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

        // 배치 모드
        _table.WrapMode = WrapFrontRadio.IsChecked  == true ? TableWrapMode.InFrontOfText
                        : WrapBehindRadio.IsChecked == true ? TableWrapMode.BehindText
                        : WrapFixedRadio.IsChecked  == true ? TableWrapMode.Fixed
                        : TableWrapMode.Block;

        _table.HAlign = AlignCenterRadio.IsChecked == true ? TableHAlign.Center
                      : AlignRightRadio.IsChecked  == true ? TableHAlign.Right
                      : TableHAlign.Left;

        _table.WidthMm  = tableW;
        _table.HeightMm = tableH;

        _table.BackgroundColor = bgColor.Length > 0 ? bgColor : null;

        _table.DefaultCellPaddingTopMm    = cpt;
        _table.DefaultCellPaddingBottomMm = cpb;
        _table.DefaultCellPaddingLeftMm   = cpl;
        _table.DefaultCellPaddingRightMm  = cpr;

        _table.OuterMarginTopMm    = omt;
        _table.OuterMarginBottomMm = omb;
        _table.OuterMarginLeftMm   = oml;
        _table.OuterMarginRightMm  = omr;

        if (!int.TryParse(HeaderColumnCountBox.Text.Trim(), out int headerColCount) || headerColCount < 0)
        {
            MessageBox.Show(this, "헤더 열 수는 0 이상의 정수로 입력하세요.", "표 속성",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            HeaderColumnCountBox.Focus();
            return;
        }

        _table.BorderThicknessPt = borderPt;
        _table.BorderColor       = borderColor.Length > 0 ? borderColor : null;

        // 면별 외곽선 — 유효성 검사 후 적용.
        foreach (var editor in new[] { SideTop, SideBottom, SideLeft, SideRight, SideInnerH, SideInnerV })
        {
            if (!editor.TryValidate(out string? sideError))
            {
                MessageBox.Show(this, sideError ?? "면별 외곽선 입력이 잘못되었습니다.", "표 속성",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
        }
        _table.BorderTop             = SideTop.Value;
        _table.BorderBottom          = SideBottom.Value;
        _table.BorderLeft            = SideLeft.Value;
        _table.BorderRight           = SideRight.Value;
        _table.InnerBorderHorizontal = SideInnerH.Value;
        _table.InnerBorderVertical   = SideInnerV.Value;

        // 페이지 분할
        _table.RepeatHeaderRowsOnBreak = RepeatHeaderRowsCheck.IsChecked == true;
        _table.HeaderColumnCount       = headerColCount;

        // 테두리 병합
        _table.BorderCollapse = BorderCollapseCheck.IsChecked == true;

        DialogResult = true;
        Close();
    }

    private void OnCancelClick(object sender, RoutedEventArgs e) => Close();

    // ── 유틸 ─────────────────────────────────────────────────────────────

    private static bool TryParseNonNeg(string text, out double value)
        => double.TryParse(text.Trim(), out value) && value >= 0;
}

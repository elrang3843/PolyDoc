using System;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using PolyDonky.Core;

namespace PolyDonky.App.Controls;

/// <summary>
/// 표/셀 외곽선의 한 면(상·하·좌·우, 가운데 가로/세로) 을 편집하는 재사용 가능한 컨트롤.
/// "설정" 체크박스가 꺼져 있으면 면 지정 = null (공통값/cascade 상위가 적용됨).
/// 켜져 있으면 두께·색상·선 종류로 면 지정.
/// </summary>
public partial class BorderSideEditor : UserControl
{
    public BorderSideEditor()
    {
        InitializeComponent();
        UpdateEnabledState();
    }

    /// <summary>좌측 라벨(예: "위", "왼쪽", "가운데 가로") 을 설정한다.</summary>
    public string Label
    {
        get => LabelText.Text;
        set => LabelText.Text = value;
    }

    /// <summary>현재 편집중인 면의 spec. null = "설정" 체크 해제 상태.</summary>
    public CellBorderSide? Value
    {
        get
        {
            if (EnableCheck.IsChecked != true) return null;
            if (!double.TryParse(ThicknessBox.Text.Trim(),
                    NumberStyles.Any, CultureInfo.InvariantCulture, out double thk) || thk < 0)
                thk = 0;
            string? color = ColorPicker.ColorText.Trim();
            if (string.IsNullOrEmpty(color)) color = null;
            BorderLineStyle ls = BorderLineStyle.Solid;
            if (StyleCombo.SelectedItem is ComboBoxItem ci && ci.Tag is string tag
                && Enum.TryParse(tag, out BorderLineStyle parsed))
                ls = parsed;
            return new CellBorderSide(thk, color, ls);
        }
        set
        {
            if (value is null)
            {
                EnableCheck.IsChecked = false;
                UpdateEnabledState();
                return;
            }
            var v = value.Value;
            EnableCheck.IsChecked = true;
            ThicknessBox.Text     = v.ThicknessPt.ToString("0.##", CultureInfo.InvariantCulture);
            ColorPicker.ColorText = v.Color ?? string.Empty;
            for (int i = 0; i < StyleCombo.Items.Count; i++)
            {
                if (StyleCombo.Items[i] is ComboBoxItem ci
                    && ci.Tag is string tag
                    && Enum.TryParse(tag, out BorderLineStyle parsed)
                    && parsed == v.LineStyle)
                {
                    StyleCombo.SelectedIndex = i;
                    break;
                }
            }
            UpdateEnabledState();
        }
    }

    /// <summary>유효성 검사 — true 면 OK, false 면 사유 메시지 반환.</summary>
    public bool TryValidate(out string? error)
    {
        error = null;
        if (EnableCheck.IsChecked != true) return true;
        var raw = ThicknessBox.Text.Trim();
        if (!double.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out double thk) || thk < 0)
        {
            error = $"{Label}: 두께는 0 이상의 숫자(pt)로 입력하세요.";
            return false;
        }
        var color = ColorPicker.ColorText.Trim();
        if (color.Length > 0 && !ColorPickerBox.TryParseColor(color, out _))
        {
            error = $"{Label}: 색상을 올바른 색상 값으로 입력하세요 (예: #C8C8C8).";
            return false;
        }
        return true;
    }

    private void OnEnableChanged(object sender, RoutedEventArgs e) => UpdateEnabledState();

    private void UpdateEnabledState()
    {
        bool on = EnableCheck.IsChecked == true;
        ThicknessBox.IsEnabled = on;
        ColorPicker.IsEnabled  = on;
        StyleCombo.IsEnabled   = on;
    }
}

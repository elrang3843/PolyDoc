using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using PolyDonky.Core;

namespace PolyDonky.App.Views;

/// <summary>
/// 그림(ImageBlock) 속성 편집 다이얼로그. 크기·정렬·여백·테두리·설명 변경.
/// OK 클릭 시 <see cref="ImageBlock"/> 모델을 직접 갱신하고 DialogResult = true 반환.
/// 호출측에서 FlowDocument 의 BlockUIContainer 를 재빌드해 반영해야 한다.
/// </summary>
public partial class ImagePropertiesWindow : Window
{
    private readonly ImageBlock _image;
    private bool _suppressSync;

    private static readonly (ImageWrapMode Value, string Label)[] WrapOptions =
    {
        (ImageWrapMode.Inline,        "본문 흐름 (텍스트 위·아래)"),
        (ImageWrapMode.AsText,        "텍스트로 처리 — 글자처럼 한 줄에 인라인"),
        (ImageWrapMode.WrapLeft,      "왼쪽 배치 — 텍스트가 오른쪽으로 흐름"),
        (ImageWrapMode.WrapRight,     "오른쪽 배치 — 텍스트가 왼쪽으로 흐름"),
        (ImageWrapMode.InFrontOfText, "텍스트 앞으로 — 그림이 위에 겹침"),
        (ImageWrapMode.BehindText,    "텍스트 뒤로 — 그림이 아래에 깔림"),
    };

    public ImagePropertiesWindow(ImageBlock image)
    {
        InitializeComponent();
        _image = image;
        Loaded += OnLoaded;
    }

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        _suppressSync = true;

        WidthBox.Text  = _image.WidthMm.ToString("F1");
        HeightBox.Text = _image.HeightMm.ToString("F1");

        AlignLeft.IsChecked   = _image.HAlign == ImageHAlign.Left;
        AlignCenter.IsChecked = _image.HAlign == ImageHAlign.Center;
        AlignRight.IsChecked  = _image.HAlign == ImageHAlign.Right;

        foreach (var (_, label) in WrapOptions)
            WrapCombo.Items.Add(label);
        WrapCombo.SelectedIndex = Array.FindIndex(WrapOptions, w => w.Value == _image.WrapMode) is var i and >= 0 ? i : 0;

        OverlayXBox.Text = _image.OverlayXMm.ToString("F1");
        OverlayYBox.Text = _image.OverlayYMm.ToString("F1");
        UpdateOverlayFieldsVisibility();

        MarginTopBox.Text    = _image.MarginTopMm.ToString("F1");
        MarginBottomBox.Text = _image.MarginBottomMm.ToString("F1");

        BorderColorBox.Text     = _image.BorderColor ?? string.Empty;
        BorderThicknessBox.Text = _image.BorderThicknessPt.ToString("F1");

        DescriptionBox.Text = _image.Description ?? string.Empty;

        _suppressSync = false;
        UpdateBorderPreview();
    }

    private void OnWidthChanged(object sender, TextChangedEventArgs e)
    {
        if (_suppressSync || LockRatioCheck.IsChecked != true) return;
        if (!double.TryParse(WidthBox.Text, out double w) || w <= 0) return;
        if (_image.WidthMm <= 0 || _image.HeightMm <= 0) return;

        _suppressSync = true;
        HeightBox.Text = (w * _image.HeightMm / _image.WidthMm).ToString("F1");
        _suppressSync = false;
    }

    private void OnHeightChanged(object sender, TextChangedEventArgs e)
    {
        if (_suppressSync || LockRatioCheck.IsChecked != true) return;
        if (!double.TryParse(HeightBox.Text, out double h) || h <= 0) return;
        if (_image.WidthMm <= 0 || _image.HeightMm <= 0) return;

        _suppressSync = true;
        WidthBox.Text = (h * _image.WidthMm / _image.HeightMm).ToString("F1");
        _suppressSync = false;
    }

    private void OnBorderColorChanged(object sender, TextChangedEventArgs e)
        => UpdateBorderPreview();

    private void OnWrapModeChanged(object sender, SelectionChangedEventArgs e)
    {
        if (!_suppressSync) UpdateOverlayFieldsVisibility();
    }

    private void UpdateOverlayFieldsVisibility()
    {
        if (WrapCombo.SelectedIndex < 0) return;
        var mode = WrapOptions[WrapCombo.SelectedIndex].Value;
        bool isOverlay = mode is ImageWrapMode.InFrontOfText or ImageWrapMode.BehindText;
        var visibility = isOverlay ? Visibility.Visible : Visibility.Collapsed;

        OverlayXLabel.Visibility = visibility;
        OverlayXBox.Visibility   = visibility;
        OverlayYLabel.Visibility = visibility;
        OverlayYBox.Visibility   = visibility;
        OverlayMmLabel.Visibility = visibility;
    }

    private void UpdateBorderPreview()
    {
        try
        {
            var c = (System.Windows.Media.Color)ColorConverter.ConvertFromString(BorderColorBox.Text)!;
            BorderColorPreview.Fill = new SolidColorBrush(c);
        }
        catch
        {
            BorderColorPreview.Fill = Brushes.Transparent;
        }
    }

    private void OnOkClick(object sender, RoutedEventArgs e)
    {
        if (!double.TryParse(WidthBox.Text,  out double w) || w <= 0 ||
            !double.TryParse(HeightBox.Text, out double h) || h <= 0)
        {
            MessageBox.Show(this, "너비와 높이를 올바르게 입력하세요 (단위: mm).", "그림 속성",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (!double.TryParse(MarginTopBox.Text,      out double mt)) mt = 0;
        if (!double.TryParse(MarginBottomBox.Text,   out double mb)) mb = 0;
        if (!double.TryParse(BorderThicknessBox.Text, out double bt)) bt = 0;

        _image.WidthMm  = w;
        _image.HeightMm = h;
        _image.HAlign   = AlignCenter.IsChecked == true ? ImageHAlign.Center
                        : AlignRight.IsChecked  == true ? ImageHAlign.Right
                        : ImageHAlign.Left;
        if (WrapCombo.SelectedIndex >= 0)
            _image.WrapMode = WrapOptions[WrapCombo.SelectedIndex].Value;
        if (double.TryParse(OverlayXBox.Text, out double ox)) _image.OverlayXMm = Math.Max(0, ox);
        if (double.TryParse(OverlayYBox.Text, out double oy)) _image.OverlayYMm = Math.Max(0, oy);
        _image.MarginTopMm    = Math.Max(0, mt);
        _image.MarginBottomMm = Math.Max(0, mb);

        var colorText = BorderColorBox.Text.Trim();
        _image.BorderThicknessPt = Math.Max(0, bt);
        _image.BorderColor = bt > 0 && colorText.Length > 0 ? colorText : null;

        _image.Description = DescriptionBox.Text.Trim() is { Length: > 0 } d ? d : null;

        DialogResult = true;
        Close();
    }

    private void OnCancelClick(object sender, RoutedEventArgs e) => Close();
}

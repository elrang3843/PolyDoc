using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using PolyDonky.App.Services;
using PolyDonky.Core;
using SR = PolyDonky.App.Properties.Resources;
using Wpf = System.Windows.Documents;
using WpfMedia = System.Windows.Media;

namespace PolyDonky.App.Views;

/// <summary>
/// 개요 수준별 서식 편집 다이얼로그.
/// 글자 모양 / 문단 모양 편집은 기존 CharFormatWindow / ParaFormatWindow 를 단독 모드로 호출해 재사용.
/// </summary>
public partial class OutlineStyleWindow : Window
{
    private OutlineStyleSet _styleSet;
    private bool _suppress;

    // 변경을 ViewModel 에 넘기는 콜백 — 창을 닫지 않고 실시간으로 문서에 반영할 수도 있다.
    public event EventHandler<OutlineStyleSet>? StyleApplied;

    public OutlineStyleWindow(OutlineStyleSet current)
    {
        _styleSet = DeepClone(current);
        InitializeComponent();
        Loaded += OnLoaded;
    }

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        PopulatePresets();
        PopulateLevels();
        PopulateNumberStyles();
        CboLevel.SelectedIndex = 0;
    }

    // ── 초기화 ──────────────────────────────────────────────────

    private void PopulatePresets()
    {
        var presetNames = new[]
        {
            SR.FormatOutlinePresetDefault,
            SR.FormatOutlinePresetAcademic,
            SR.FormatOutlinePresetBusiness,
            SR.FormatOutlinePresetModern,
        };
        foreach (var name in presetNames)
            CboPreset.Items.Add(name);

        // 현재 세트 이름과 일치하는 프리셋 선택
        int match = presetNames.ToList().FindIndex(n => n == _styleSet.Name);
        _suppress = true;
        CboPreset.SelectedIndex = match >= 0 ? match : -1;
        _suppress = false;
    }

    private void PopulateLevels()
    {
        CboLevel.Items.Add(new ComboBoxItem { Content = SR.FormatOutlineLevelBody, Tag = OutlineLevel.Body });
        for (int i = 1; i <= 6; i++)
            CboLevel.Items.Add(new ComboBoxItem { Content = $"H{i}", Tag = (OutlineLevel)i });
    }

    private void PopulateNumberStyles()
    {
        var items = new[]
        {
            (SR.FormatOutlineNumberNone,       NumberingStyle.None),
            (SR.FormatOutlineNumberDecimal,    NumberingStyle.Decimal),
            (SR.FormatOutlineNumberAlphaLower, NumberingStyle.AlphaLower),
            (SR.FormatOutlineNumberAlphaUpper, NumberingStyle.AlphaUpper),
            (SR.FormatOutlineNumberRomanLower, NumberingStyle.RomanLower),
            (SR.FormatOutlineNumberRomanUpper, NumberingStyle.RomanUpper),
            (SR.FormatOutlineNumberHangul,     NumberingStyle.HangulSyllable),
        };
        foreach (var (label, style) in items)
            CboNumberStyle.Items.Add(new ComboBoxItem { Content = label, Tag = style });
    }

    // ── 현재 수준 가져오기 ──────────────────────────────────────

    private OutlineLevel CurrentLevel =>
        CboLevel.SelectedItem is ComboBoxItem ci && ci.Tag is OutlineLevel ol ? ol : OutlineLevel.Body;

    private OutlineLevelStyle CurrentLevelStyle => _styleSet.GetLevel(CurrentLevel);

    // ── 수준 변경 시 UI 갱신 ────────────────────────────────────

    private void OnLevelChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_suppress) return;
        LoadLevelUI();
    }

    private void LoadLevelUI()
    {
        _suppress = true;
        try
        {
            var ls = CurrentLevelStyle;
            RefreshSummaries(ls);

            // 번호 스타일
            var nStyle = ls.Numbering.Style;
            for (int i = 0; i < CboNumberStyle.Items.Count; i++)
            {
                if (CboNumberStyle.Items[i] is ComboBoxItem ci && ci.Tag is NumberingStyle ns && ns == nStyle)
                {
                    CboNumberStyle.SelectedIndex = i;
                    break;
                }
            }
            if (CboNumberStyle.SelectedIndex < 0) CboNumberStyle.SelectedIndex = 0;

            TxtPrefix.Text = ls.Numbering.Prefix;
            TxtSuffix.Text = ls.Numbering.Suffix;
            PanelPrefixSuffix.IsEnabled = ls.Numbering.Style != NumberingStyle.None;

            // 테두리
            ChkBorderTop.IsChecked    = ls.Border.ShowTop;
            ChkBorderBottom.IsChecked = ls.Border.ShowBottom;
            TxtBorderColor.Text       = ls.Border.Color ?? "";
            if (TryParseColor(ls.Border.Color, out var bc))
                BorderColorSwatch.Background = new WpfMedia.SolidColorBrush(bc);
            else
                BorderColorSwatch.Background = null;

            // 배경색
            bool bgNone = string.IsNullOrEmpty(ls.BackgroundColor);
            ChkBgNone.IsChecked   = bgNone;
            PanelBgColor.IsEnabled = !bgNone;
            TxtBgColor.Text = ls.BackgroundColor ?? "";
            if (TryParseColor(ls.BackgroundColor, out var bgc))
                BgColorSwatch.Background = new WpfMedia.SolidColorBrush(bgc);
            else
                BgColorSwatch.Background = null;
        }
        finally
        {
            _suppress = false;
        }
        UpdatePreview();
    }

    private void RefreshSummaries(OutlineLevelStyle ls)
    {
        var c = ls.Char;
        var parts = new System.Collections.Generic.List<string>();
        if (!string.IsNullOrEmpty(c.FontFamily)) parts.Add(c.FontFamily);
        parts.Add($"{c.FontSizePt:0.#}pt");
        if (c.Bold)   parts.Add("굵게");
        if (c.Italic) parts.Add("기울임");
        if (c.Foreground is { } fg) parts.Add($"#{fg.R:X2}{fg.G:X2}{fg.B:X2}");
        TxtCharSummary.Text = string.Join(", ", parts);

        var p = ls.Para;
        var pParts = new System.Collections.Generic.List<string>();
        pParts.Add(p.Alignment switch
        {
            Alignment.Center      => "가운데",
            Alignment.Right       => "오른쪽",
            Alignment.Justify     => "양쪽",
            Alignment.Distributed => "배분",
            _                     => "왼쪽",
        });
        pParts.Add($"{p.LineHeightFactor * 100:0}%");
        if (p.SpaceBeforePt > 0) pParts.Add($"↑{p.SpaceBeforePt:0.#}pt");
        if (p.SpaceAfterPt  > 0) pParts.Add($"↓{p.SpaceAfterPt:0.#}pt");
        TxtParaSummary.Text = string.Join(", ", pParts);
    }

    // ── 프리셋 선택 ─────────────────────────────────────────────

    private void OnPresetChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_suppress) return;
        var presets = OutlineStyleSet.BuiltInPresets;
        int idx = CboPreset.SelectedIndex;
        if (idx < 0 || idx >= presets.Count) return;
        _styleSet = DeepClone(presets[idx]);
        LoadLevelUI();
    }

    // ── 글자 모양 편집 ──────────────────────────────────────────

    private void OnEditChar(object sender, RoutedEventArgs e)
    {
        var ls = CurrentLevelStyle;
        var dlg = new CharFormatWindow(CloneRunStyle(ls.Char)) { Owner = this };
        if (dlg.ShowDialog() == true && dlg.ResultStyle is { } result)
        {
            ls.Char = result;
            RefreshSummaries(ls);
            UpdatePreview();
        }
    }

    // ── 문단 모양 편집 ──────────────────────────────────────────

    private void OnEditPara(object sender, RoutedEventArgs e)
    {
        var ls = CurrentLevelStyle;
        var dlg = new ParaFormatWindow(CloneParagraphStyle(ls.Para)) { Owner = this };
        if (dlg.ShowDialog() == true && dlg.ResultStyle is { } result)
        {
            ls.Para = result;
            RefreshSummaries(ls);
            UpdatePreview();
        }
    }

    // ── 번호 스타일 / 접두접미 ──────────────────────────────────

    private void OnNumberStyleChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_suppress) return;
        if (CboNumberStyle.SelectedItem is ComboBoxItem ci && ci.Tag is NumberingStyle ns)
        {
            CurrentLevelStyle.Numbering.Style = ns;
            PanelPrefixSuffix.IsEnabled = ns != NumberingStyle.None;
        }
        UpdatePreview();
    }

    private void OnNumberingChanged(object sender, TextChangedEventArgs e)
    {
        if (_suppress) return;
        CurrentLevelStyle.Numbering.Prefix = TxtPrefix.Text;
        CurrentLevelStyle.Numbering.Suffix = TxtSuffix.Text;
        UpdatePreview();
    }

    // ── 테두리 ──────────────────────────────────────────────────

    private void OnBorderChanged(object sender, RoutedEventArgs e)
    {
        if (_suppress) return;
        CurrentLevelStyle.Border.ShowTop    = ChkBorderTop.IsChecked == true;
        CurrentLevelStyle.Border.ShowBottom = ChkBorderBottom.IsChecked == true;
        UpdatePreview();
    }

    private void OnBorderSwatchClick(object sender, MouseButtonEventArgs e)
        => PickColor(TxtBorderColor);

    private void OnBorderColorChanged(object sender, TextChangedEventArgs e)
    {
        if (_suppress) return;
        var hex = TxtBorderColor.Text.Trim();
        if (TryParseColor(hex, out var c))
        {
            BorderColorSwatch.Background = new WpfMedia.SolidColorBrush(c);
            CurrentLevelStyle.Border.Color = NormalizeHex(hex);
        }
        else
        {
            BorderColorSwatch.Background = null;
            CurrentLevelStyle.Border.Color = null;
        }
        UpdatePreview();
    }

    // ── 배경색 ──────────────────────────────────────────────────

    private void OnBgSwatchClick(object sender, MouseButtonEventArgs e)
        => PickColor(TxtBgColor);

    private void OnBgNoneChanged(object sender, RoutedEventArgs e)
    {
        if (_suppress) return;
        bool none = ChkBgNone.IsChecked == true;
        PanelBgColor.IsEnabled = !none;
        if (none)
        {
            CurrentLevelStyle.BackgroundColor = null;
            BgColorSwatch.Background = null;
        }
        else if (TryParseColor(TxtBgColor.Text, out var c))
        {
            CurrentLevelStyle.BackgroundColor = NormalizeHex(TxtBgColor.Text.Trim());
            BgColorSwatch.Background = new WpfMedia.SolidColorBrush(c);
        }
        UpdatePreview();
    }

    private void OnBgColorChanged(object sender, TextChangedEventArgs e)
    {
        if (_suppress) return;
        var hex = TxtBgColor.Text.Trim();
        if (TryParseColor(hex, out var c))
        {
            BgColorSwatch.Background = new WpfMedia.SolidColorBrush(c);
            CurrentLevelStyle.BackgroundColor = NormalizeHex(hex);
        }
        else
        {
            BgColorSwatch.Background = null;
            CurrentLevelStyle.BackgroundColor = null;
        }
        UpdatePreview();
    }

    // ── 미리보기 ────────────────────────────────────────────────

    private void UpdatePreview()
    {
        var fd = new FlowDocument
        {
            FontFamily = new WpfMedia.FontFamily("맑은 고딕, Malgun Gothic, Segoe UI"),
            FontSize   = FlowDocumentBuilder.PtToDip(11),
            PagePadding = new Thickness(8),
        };

        // 현재 수준 + 본문 예시 한 단락
        var curLevel = CurrentLevel;
        var previewLevels = curLevel == OutlineLevel.Body
            ? new[] { OutlineLevel.Body }
            : new[] { curLevel, OutlineLevel.Body };

        foreach (var level in previewLevels)
        {
            var ls = _styleSet.GetLevel(level);
            var wpfPara = BuildPreviewParagraph(ls, level, _styleSet);
            fd.Blocks.Add(wpfPara);
        }

        PreviewViewer.Document = fd;
    }

    private static Wpf.Paragraph BuildPreviewParagraph(OutlineLevelStyle ls, OutlineLevel level, OutlineStyleSet set)
    {
        var c = ls.Char;
        var p = ls.Para;
        var fontSize = FlowDocumentBuilder.PtToDip(c.FontSizePt > 0 ? c.FontSizePt : 11);

        var sample = level == OutlineLevel.Body
            ? "가나다라마바사 ABCDEFG 1234567. 본문 텍스트 미리보기입니다."
            : $"H{(int)level} — {(string.IsNullOrEmpty(c.FontFamily) ? "맑은 고딕" : c.FontFamily)} {c.FontSizePt:0.#}pt{(c.Bold ? " 굵게" : "")}{(c.Italic ? " 기울임" : "")}";

        // 번호 접두어 추가
        if (ls.Numbering.Style != NumberingStyle.None)
        {
            var num = ls.Numbering.Prefix + NumberSample(ls.Numbering.Style, level) + ls.Numbering.Suffix;
            sample = num + " " + sample;
        }

        var run = new Wpf.Run(sample)
        {
            FontSize   = fontSize,
            FontWeight = c.Bold   ? FontWeights.Bold    : FontWeights.Normal,
            FontStyle  = c.Italic ? FontStyles.Italic   : FontStyles.Normal,
        };
        if (!string.IsNullOrEmpty(c.FontFamily))
            run.FontFamily = new WpfMedia.FontFamily(c.FontFamily);
        if (c.Foreground is { } fg)
            run.Foreground = new WpfMedia.SolidColorBrush(WpfMedia.Color.FromArgb(fg.A, fg.R, fg.G, fg.B));

        var para = new Wpf.Paragraph(run)
        {
            TextAlignment = p.Alignment switch
            {
                Alignment.Center      => TextAlignment.Center,
                Alignment.Right       => TextAlignment.Right,
                Alignment.Justify or Alignment.Distributed => TextAlignment.Justify,
                _                     => TextAlignment.Left,
            },
            Margin = new Thickness(
                p.IndentLeftMm > 0 ? FlowDocumentBuilder.MmToDip(p.IndentLeftMm) : 0,
                p.SpaceBeforePt > 0 ? FlowDocumentBuilder.PtToDip(p.SpaceBeforePt) : 0,
                p.IndentRightMm > 0 ? FlowDocumentBuilder.MmToDip(p.IndentRightMm) : 0,
                p.SpaceAfterPt > 0  ? FlowDocumentBuilder.PtToDip(p.SpaceAfterPt)  : 0),
            LineHeight = p.LineHeightFactor > 0.1 ? fontSize * p.LineHeightFactor : double.NaN,
            TextIndent = p.IndentFirstLineMm != 0 ? FlowDocumentBuilder.MmToDip(p.IndentFirstLineMm) : 0,
        };

        // 배경색
        if (!string.IsNullOrEmpty(ls.BackgroundColor) && TryParseColor(ls.BackgroundColor, out var bgc))
            para.Background = new WpfMedia.SolidColorBrush(bgc);

        // 테두리 — WPF Paragraph 의 Border* 속성으로 표현
        if (ls.Border.ShowTop || ls.Border.ShowBottom)
        {
            WpfMedia.SolidColorBrush borderBrush = !string.IsNullOrEmpty(ls.Border.Color) && TryParseColor(ls.Border.Color, out var brc)
                ? new WpfMedia.SolidColorBrush(brc)
                : new WpfMedia.SolidColorBrush(WpfMedia.Colors.DimGray);
            para.BorderBrush     = borderBrush;
            para.BorderThickness = new Thickness(0,
                ls.Border.ShowTop    ? 1 : 0,
                0,
                ls.Border.ShowBottom ? 1 : 0);
        }

        return para;
    }

    private static string NumberSample(NumberingStyle style, OutlineLevel level)
    {
        var n = (int)level;
        return style switch
        {
            NumberingStyle.Decimal        => n.ToString(),
            NumberingStyle.AlphaLower     => ((char)('a' + n - 1)).ToString(),
            NumberingStyle.AlphaUpper     => ((char)('A' + n - 1)).ToString(),
            NumberingStyle.RomanLower     => ToRoman(n).ToLower(),
            NumberingStyle.RomanUpper     => ToRoman(n),
            NumberingStyle.HangulSyllable => ((char)(0xAC00 + (n - 1) * 588)).ToString(),
            _                             => n.ToString(),
        };
    }

    private static string ToRoman(int n)
    {
        var vals  = new[] { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
        var syms  = new[] { "M","CM","D","CD","C","XC","L","XL","X","IX","V","IV","I" };
        var sb = new System.Text.StringBuilder();
        for (int i = 0; i < vals.Length && n > 0; i++)
            while (n >= vals[i]) { sb.Append(syms[i]); n -= vals[i]; }
        return sb.Length > 0 ? sb.ToString() : "I";
    }

    // ── 버튼 ────────────────────────────────────────────────────

    private void OnResetLevel(object sender, RoutedEventArgs e)
    {
        var level = CurrentLevel;
        _styleSet.SetLevel(level, OutlineStyleSet.DefaultForLevel(level));
        LoadLevelUI();
    }

    private void OnApply(object sender, RoutedEventArgs e)
    {
        StyleApplied?.Invoke(this, DeepClone(_styleSet));
        DialogResult = true;
    }

    private void OnClose(object sender, RoutedEventArgs e) => DialogResult = false;

    // ── 유틸 ────────────────────────────────────────────────────

    private void PickColor(System.Windows.Controls.TextBox target)
    {
        using var dlg = new System.Windows.Forms.ColorDialog { FullOpen = true, AnyColor = true };
        if (TryParseColor(target.Text, out var current))
            dlg.Color = System.Drawing.Color.FromArgb(current.A, current.R, current.G, current.B);
        if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            var p = dlg.Color;
            target.Text = $"#{p.R:X2}{p.G:X2}{p.B:X2}";
        }
    }

    private static bool TryParseColor(string? hex, out WpfMedia.Color color)
    {
        color = default;
        if (string.IsNullOrWhiteSpace(hex)) return false;
        var t = hex.Trim();
        if (!t.StartsWith('#')) t = '#' + t;
        try { color = (WpfMedia.Color)WpfMedia.ColorConverter.ConvertFromString(t); return true; }
        catch { return false; }
    }

    private static string NormalizeHex(string hex)
    {
        var t = hex.Trim();
        return t.StartsWith('#') ? t : '#' + t;
    }

    private static PolyDonky.Core.RunStyle CloneRunStyle(PolyDonky.Core.RunStyle s) => s.Clone();

    private static PolyDonky.Core.ParagraphStyle CloneParagraphStyle(PolyDonky.Core.ParagraphStyle s) => s.Clone();

    private static OutlineStyleSet DeepClone(OutlineStyleSet src)
    {
        var dst = new OutlineStyleSet { Name = src.Name };
        foreach (var (key, ls) in src.Levels)
        {
            dst.Levels[key] = new OutlineLevelStyle
            {
                Char  = CloneRunStyle(ls.Char),
                Para  = CloneParagraphStyle(ls.Para),
                Numbering = new OutlineNumbering
                {
                    Style            = ls.Numbering.Style,
                    Prefix           = ls.Numbering.Prefix,
                    Suffix           = ls.Numbering.Suffix,
                    RestartFromHigher = ls.Numbering.RestartFromHigher,
                },
                Border = new OutlineBorder
                {
                    ShowTop    = ls.Border.ShowTop,
                    ShowBottom = ls.Border.ShowBottom,
                    Color      = ls.Border.Color,
                },
                BackgroundColor = ls.BackgroundColor,
            };
        }
        return dst;
    }
}

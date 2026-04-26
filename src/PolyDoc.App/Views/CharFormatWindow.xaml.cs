using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using PolyDoc.App.Services;

namespace PolyDoc.App.Views;

/// <summary>
/// 글자 서식 다이얼로그. RichTextBox 현재 선택 영역의 서식을 읽어 초기화하고,
/// OK 시 선택 영역에 서식을 일괄 적용한다.
/// 선택이 비어 있으면 캐럿 위치의 서식을 보여주고, 적용 시 이후 입력 문자에 반영된다.
/// </summary>
public partial class CharFormatWindow : Window
{
    private readonly RichTextBox _editor;
    private bool _suppressPreview;

    public CharFormatWindow(RichTextBox editor)
    {
        _editor = editor;
        InitializeComponent();
        PopulateFontFamilies();
        Loaded += (_, _) => LoadCurrentFormatting();
    }

    // ── 초기화 ──────────────────────────────────────────────────

    private void PopulateFontFamilies()
    {
        var fonts = Fonts.SystemFontFamilies
            .OrderBy(f => f.Source, StringComparer.OrdinalIgnoreCase)
            .Select(f => f.Source)
            .ToList();
        CboFont.ItemsSource = fonts;
    }

    private void LoadCurrentFormatting()
    {
        _suppressPreview = true;
        try
        {
            var sel = _editor.Selection;

            // ── 글꼴 이름 ──
            var fontFamilyVal = sel.GetPropertyValue(TextElement.FontFamilyProperty);
            if (fontFamilyVal is FontFamily ff)
                CboFont.Text = ff.Source;

            // ── 글자 크기 ──
            var fontSizeVal = sel.GetPropertyValue(TextElement.FontSizeProperty);
            if (fontSizeVal is double dips)
                TxtSize.Text = FlowDocumentBuilder.DipToPt(dips).ToString("0.#");

            // ── 굵게 ──
            var weightVal = sel.GetPropertyValue(TextElement.FontWeightProperty);
            ChkBold.IsChecked = weightVal == DependencyProperty.UnsetValue
                ? null
                : (bool?)(weightVal is FontWeight fw && fw.ToOpenTypeWeight() >= FontWeights.Bold.ToOpenTypeWeight());

            // ── 기울임꼴 ──
            var styleVal = sel.GetPropertyValue(TextElement.FontStyleProperty);
            ChkItalic.IsChecked = styleVal == DependencyProperty.UnsetValue
                ? null
                : (bool?)(styleVal is FontStyle fs && fs == FontStyles.Italic);

            // ── 텍스트 장식 (밑줄 / 취소선 / 위선) ──
            var decoVal = sel.GetPropertyValue(Inline.TextDecorationsProperty);
            if (decoVal == DependencyProperty.UnsetValue)
            {
                ChkUnderline.IsChecked    = null;
                ChkStrikethrough.IsChecked = null;
                ChkOverline.IsChecked     = null;
            }
            else if (decoVal is TextDecorationCollection tdc)
            {
                ChkUnderline.IsChecked    = tdc.Any(d => d.Location == TextDecorationLocation.Underline);
                ChkStrikethrough.IsChecked = tdc.Any(d => d.Location == TextDecorationLocation.Strikethrough);
                ChkOverline.IsChecked     = tdc.Any(d => d.Location == TextDecorationLocation.OverLine);
            }
            else
            {
                ChkUnderline.IsChecked    = false;
                ChkStrikethrough.IsChecked = false;
                ChkOverline.IsChecked     = false;
            }

            // ── 위/아래첨자 ──
            var baselineVal = sel.GetPropertyValue(Inline.BaselineAlignmentProperty);
            if (baselineVal == DependencyProperty.UnsetValue)
            {
                ChkSuperscript.IsChecked = null;
                ChkSubscript.IsChecked   = null;
            }
            else if (baselineVal is BaselineAlignment ba)
            {
                ChkSuperscript.IsChecked = ba == BaselineAlignment.Superscript;
                ChkSubscript.IsChecked   = ba == BaselineAlignment.Subscript;
            }

            // ── 글자색 ──
            var fgVal = sel.GetPropertyValue(TextElement.ForegroundProperty);
            if (fgVal is SolidColorBrush fgBrush)
            {
                TxtFgColor.Text   = ToHex(fgBrush.Color);
                FgSwatch.Background = new SolidColorBrush(fgBrush.Color);
            }

            // ── 배경색 ──
            var bgVal = sel.GetPropertyValue(TextElement.BackgroundProperty);
            if (bgVal is SolidColorBrush bgBrush)
            {
                TxtBgColor.Text   = ToHex(bgBrush.Color);
                BgSwatch.Background = new SolidColorBrush(bgBrush.Color);
            }

            // ── 글자폭 / 자간: 선택 영역 첫 인라인에서 읽음 ──
            TxtWidthPercent.Text  = "100";
            TxtLetterSpacing.Text = "0";
            var firstInline = GetFirstInlineInSelection();
            if (firstInline is Run wr)
            {
                if (wr.Tag is PolyDoc.Core.Run pr && Math.Abs(pr.Style.WidthPercent - 100) > 0.1)
                    TxtWidthPercent.Text = pr.Style.WidthPercent.ToString("0.#");
                var cs = Typography.GetCharacterSpacing(wr);
                if (cs != 0)
                {
                    var dip = wr.FontSize > 0 ? wr.FontSize : FlowDocumentBuilder.PtToDip(11);
                    TxtLetterSpacing.Text = (cs / 1000.0 * dip).ToString("0.##");
                }
            }
            else if (firstInline is InlineUIContainer { Tag: PolyDoc.Core.Run scaledRun })
            {
                TxtWidthPercent.Text  = scaledRun.Style.WidthPercent.ToString("0.#");
                TxtLetterSpacing.Text = scaledRun.Style.LetterSpacingPx.ToString("0.##");
            }
        }
        finally
        {
            _suppressPreview = false;
        }

        UpdatePreview();
        CboFont.Focus();
    }

    // ── 이벤트 핸들러 ────────────────────────────────────────────

    private void OnPreviewUpdate(object sender, EventArgs e) => UpdatePreview();

    private void OnStyleClick(object sender, RoutedEventArgs e) => UpdatePreview();

    private void OnSuperscriptClick(object sender, RoutedEventArgs e)
    {
        if (ChkSuperscript.IsChecked == true) ChkSubscript.IsChecked = false;
        UpdatePreview();
    }

    private void OnSubscriptClick(object sender, RoutedEventArgs e)
    {
        if (ChkSubscript.IsChecked == true) ChkSuperscript.IsChecked = false;
        UpdatePreview();
    }

    private void OnFgSwatchClick(object sender, MouseButtonEventArgs e)
        => TxtFgColor.Focus();

    private void OnBgSwatchClick(object sender, MouseButtonEventArgs e)
        => TxtBgColor.Focus();

    private void OnFgColorChanged(object sender, TextChangedEventArgs e)
    {
        if (TryParseColor(TxtFgColor.Text, out var color))
            FgSwatch.Background = new SolidColorBrush(color);
        UpdatePreview();
    }

    private void OnBgColorChanged(object sender, TextChangedEventArgs e)
    {
        if (TryParseColor(TxtBgColor.Text, out var color))
            BgSwatch.Background = new SolidColorBrush(color);
        else if (string.IsNullOrWhiteSpace(TxtBgColor.Text))
            BgSwatch.Background = null;
        UpdatePreview();
    }

    // ── 미리보기 갱신 ────────────────────────────────────────────

    private void UpdatePreview()
    {
        if (_suppressPreview) return;

        // 글꼴
        if (!string.IsNullOrWhiteSpace(CboFont.Text))
        {
            try { PreviewText.FontFamily = new FontFamily(CboFont.Text.Trim()); }
            catch { /* 잘못된 폰트 이름 무시 */ }
        }

        // 크기 (최소 4, 최대 144pt)
        if (double.TryParse(TxtSize.Text, out var sizePt) && sizePt >= 1)
            PreviewText.FontSize = FlowDocumentBuilder.PtToDip(Math.Clamp(sizePt, 4, 144));

        // 굵기/기울임꼴
        PreviewText.FontWeight = ChkBold.IsChecked == true   ? FontWeights.Bold   : FontWeights.Normal;
        PreviewText.FontStyle  = ChkItalic.IsChecked == true ? FontStyles.Italic  : FontStyles.Normal;

        // 텍스트 장식
        var decos = new TextDecorationCollection();
        if (ChkUnderline.IsChecked == true)
            foreach (var d in TextDecorations.Underline) decos.Add(d);
        if (ChkStrikethrough.IsChecked == true)
            foreach (var d in TextDecorations.Strikethrough) decos.Add(d);
        if (ChkOverline.IsChecked == true)
            foreach (var d in TextDecorations.OverLine) decos.Add(d);
        PreviewText.TextDecorations = decos;

        // 글자색
        if (TryParseColor(TxtFgColor.Text, out var fgColor))
            PreviewText.Foreground = new SolidColorBrush(fgColor);

        // 배경색
        if (TryParseColor(TxtBgColor.Text, out var bgColor))
            PreviewText.Background = new SolidColorBrush(bgColor);
        else if (string.IsNullOrWhiteSpace(TxtBgColor.Text))
            PreviewText.Background = null;

        // 글자폭: LayoutTransform ScaleTransform
        if (double.TryParse(TxtWidthPercent.Text, out var wp) && wp >= 1)
            PreviewText.LayoutTransform = new ScaleTransform(wp / 100.0, 1.0);
        else
            PreviewText.LayoutTransform = Transform.Identity;

        // 자간: Typography.CharacterSpacing
        if (double.TryParse(TxtLetterSpacing.Text, out var lsPx) && PreviewText.FontSize > 0)
        {
            var emUnits = (int)Math.Round(lsPx / PreviewText.FontSize * 1000);
            Typography.SetCharacterSpacing(PreviewText, emUnits);
        }
    }

    // ── OK / Cancel ──────────────────────────────────────────────

    private void OnOk(object sender, RoutedEventArgs e)
    {
        ApplyToSelection();
        DialogResult = true;
    }

    private void OnCancel(object sender, RoutedEventArgs e) => DialogResult = false;

    // ── 선택 영역 서식 적용 ──────────────────────────────────────

    private void ApplyToSelection()
    {
        var sel = _editor.Selection;

        // 글꼴 이름
        var fontName = CboFont.Text?.Trim();
        if (!string.IsNullOrEmpty(fontName))
            sel.ApplyPropertyValue(TextElement.FontFamilyProperty, new FontFamily(fontName));

        // 글자 크기
        if (double.TryParse(TxtSize.Text, out var sizePt) && sizePt >= 1)
            sel.ApplyPropertyValue(TextElement.FontSizeProperty, FlowDocumentBuilder.PtToDip(sizePt));

        // 굵게
        if (ChkBold.IsChecked is bool bold)
            sel.ApplyPropertyValue(TextElement.FontWeightProperty,
                bold ? FontWeights.Bold : FontWeights.Normal);

        // 기울임꼴
        if (ChkItalic.IsChecked is bool italic)
            sel.ApplyPropertyValue(TextElement.FontStyleProperty,
                italic ? FontStyles.Italic : FontStyles.Normal);

        // 텍스트 장식
        if (ChkUnderline.IsChecked != null || ChkStrikethrough.IsChecked != null || ChkOverline.IsChecked != null)
        {
            var decos = new TextDecorationCollection();
            if (ChkUnderline.IsChecked == true)
                foreach (var d in TextDecorations.Underline) decos.Add(d);
            if (ChkStrikethrough.IsChecked == true)
                foreach (var d in TextDecorations.Strikethrough) decos.Add(d);
            if (ChkOverline.IsChecked == true)
                foreach (var d in TextDecorations.OverLine) decos.Add(d);
            sel.ApplyPropertyValue(Inline.TextDecorationsProperty, decos);
        }

        // 위/아래첨자
        if (ChkSuperscript.IsChecked == true)
            sel.ApplyPropertyValue(Inline.BaselineAlignmentProperty, BaselineAlignment.Superscript);
        else if (ChkSubscript.IsChecked == true)
            sel.ApplyPropertyValue(Inline.BaselineAlignmentProperty, BaselineAlignment.Subscript);
        else if (ChkSuperscript.IsChecked == false && ChkSubscript.IsChecked == false)
            sel.ApplyPropertyValue(Inline.BaselineAlignmentProperty, BaselineAlignment.Baseline);

        // 글자색
        if (TryParseColor(TxtFgColor.Text, out var fgColor))
            sel.ApplyPropertyValue(TextElement.ForegroundProperty, new SolidColorBrush(fgColor));

        // 배경색 — 빈 문자열이면 배경 제거
        if (string.IsNullOrWhiteSpace(TxtBgColor.Text))
            sel.ApplyPropertyValue(TextElement.BackgroundProperty, null);
        else if (TryParseColor(TxtBgColor.Text, out var bgColor))
            sel.ApplyPropertyValue(TextElement.BackgroundProperty, new SolidColorBrush(bgColor));

        // 자간: Typography.CharacterSpacing (1/1000 em 단위로 변환)
        if (double.TryParse(TxtLetterSpacing.Text, out var lsPx))
        {
            var fontSizePt = double.TryParse(TxtSize.Text, out var sp) ? sp : 11.0;
            var fontSizeDip = FlowDocumentBuilder.PtToDip(fontSizePt);
            var emUnits = (int)Math.Round(lsPx / fontSizeDip * 1000);
            sel.ApplyPropertyValue(Typography.CharacterSpacingProperty, emUnits);
        }

        // 글자폭: 선택 영역의 Run/InlineUIContainer 를 직접 재구성
        if (double.TryParse(TxtWidthPercent.Text, out var widthPct) && widthPct >= 1)
        {
            ApplyWidthPercent(widthPct);
        }
    }

    // ── 글자폭 적용 (Run ↔ InlineUIContainer 재구성) ─────────────

    private void ApplyWidthPercent(double widthPercent)
    {
        bool scaled = Math.Abs(widthPercent - 100) > 0.5;
        var sel = _editor.Selection;

        // TextPointer 로 선택 영역 내 모든 Inline 을 수집
        var inlines = new System.Collections.Generic.List<Inline>();
        var ptr = sel.Start;
        while (ptr != null && ptr.CompareTo(sel.End) < 0)
        {
            if (ptr.GetPointerContext(LogicalDirection.Forward) == TextPointerContext.ElementStart)
            {
                var elem = ptr.GetAdjacentElement(LogicalDirection.Forward);
                if (elem is Run or InlineUIContainer)
                    inlines.Add((Inline)elem);
            }
            ptr = ptr.GetNextContextPosition(LogicalDirection.Forward);
        }

        foreach (var inline in inlines)
        {
            if (scaled)
            {
                if (inline is Run run)
                {
                    // PolyDoc Run 을 Tag 에서 가져오거나 WPF 속성에서 추출
                    var polyRun = run.Tag as PolyDoc.Core.Run ?? ExtractPolyRun(run);
                    polyRun.Style.WidthPercent = widthPercent;
                    ReplaceInline(run, FlowDocumentBuilder.BuildScaledContainer(polyRun));
                }
                else if (inline is InlineUIContainer iuc && iuc.Child is TextBlock tb)
                {
                    // 기존 컨테이너의 ScaleTransform 만 갱신
                    tb.LayoutTransform = new ScaleTransform(widthPercent / 100.0, 1.0);
                    if (iuc.Tag is PolyDoc.Core.Run pr) pr.Style.WidthPercent = widthPercent;
                }
            }
            else
            {
                // 100%: InlineUIContainer → 일반 Run 으로 복원
                if (inline is InlineUIContainer iuc)
                {
                    PolyDoc.Core.Run? polyRun = null;
                    Inline newInline;
                    if (iuc.Tag is PolyDoc.Core.Run pr)
                    {
                        pr.Style.WidthPercent = 100;
                        polyRun = pr;
                        newInline = FlowDocumentBuilder.BuildInline(pr);
                    }
                    else if (iuc.Child is TextBlock tb)
                    {
                        var extracted = ExtractPolyRunFromTextBlock(tb);
                        extracted.Style.WidthPercent = 100;
                        newInline = FlowDocumentBuilder.BuildInline(extracted);
                    }
                    else
                    {
                        newInline = new Run(string.Empty);
                    }
                    ReplaceInline(iuc, newInline);
                }
                else if (inline is Run r && r.Tag is PolyDoc.Core.Run prun)
                {
                    prun.Style.WidthPercent = 100;
                }
            }
        }
    }

    private static void ReplaceInline(Inline old, Inline replacement)
    {
        if (old.Parent is Paragraph para)
        {
            para.Inlines.InsertBefore(old, replacement);
            para.Inlines.Remove(old);
        }
        else if (old.Parent is Span span)
        {
            span.Inlines.InsertBefore(old, replacement);
            span.Inlines.Remove(old);
        }
    }

    private static PolyDoc.Core.Run ExtractPolyRun(Run wpfRun)
    {
        var s = new PolyDoc.Core.RunStyle
        {
            FontFamily    = wpfRun.FontFamily?.Source,
            FontSizePt    = FlowDocumentBuilder.DipToPt(wpfRun.FontSize),
            Bold          = wpfRun.FontWeight.ToOpenTypeWeight() >= FontWeights.Bold.ToOpenTypeWeight(),
            Italic        = wpfRun.FontStyle == FontStyles.Italic,
            WidthPercent  = 100,
        };
        if (wpfRun.TextDecorations is { Count: > 0 } decos)
        {
            foreach (var d in decos)
            {
                if (d.Location == TextDecorationLocation.Underline) s.Underline = true;
                else if (d.Location == TextDecorationLocation.Strikethrough) s.Strikethrough = true;
                else if (d.Location == TextDecorationLocation.OverLine) s.Overline = true;
            }
        }
        if (wpfRun.Foreground is SolidColorBrush fg)
            s.Foreground = new PolyDoc.Core.Color(fg.Color.R, fg.Color.G, fg.Color.B, fg.Color.A);
        if (wpfRun.Background is SolidColorBrush bg)
            s.Background = new PolyDoc.Core.Color(bg.Color.R, bg.Color.G, bg.Color.B, bg.Color.A);
        var cs = Typography.GetCharacterSpacing(wpfRun);
        if (cs != 0) s.LetterSpacingPx = cs / 1000.0 * wpfRun.FontSize;

        return new PolyDoc.Core.Run { Text = wpfRun.Text, Style = s };
    }

    private static PolyDoc.Core.Run ExtractPolyRunFromTextBlock(TextBlock tb)
    {
        var s = new PolyDoc.Core.RunStyle
        {
            FontFamily   = tb.FontFamily?.Source,
            FontSizePt   = FlowDocumentBuilder.DipToPt(tb.FontSize),
            Bold         = tb.FontWeight.ToOpenTypeWeight() >= FontWeights.Bold.ToOpenTypeWeight(),
            Italic       = tb.FontStyle == FontStyles.Italic,
            WidthPercent = tb.LayoutTransform is ScaleTransform st ? st.ScaleX * 100.0 : 100,
        };
        if (tb.TextDecorations is { Count: > 0 } decos)
        {
            foreach (var d in decos)
            {
                if (d.Location == TextDecorationLocation.Underline) s.Underline = true;
                else if (d.Location == TextDecorationLocation.Strikethrough) s.Strikethrough = true;
                else if (d.Location == TextDecorationLocation.OverLine) s.Overline = true;
            }
        }
        if (tb.Foreground is SolidColorBrush fg)
            s.Foreground = new PolyDoc.Core.Color(fg.Color.R, fg.Color.G, fg.Color.B, fg.Color.A);
        if (tb.Background is SolidColorBrush bg)
            s.Background = new PolyDoc.Core.Color(bg.Color.R, bg.Color.G, bg.Color.B, bg.Color.A);
        var cs = Typography.GetCharacterSpacing(tb);
        if (cs != 0) s.LetterSpacingPx = cs / 1000.0 * tb.FontSize;

        return new PolyDoc.Core.Run { Text = tb.Text, Style = s };
    }

    // ── 선택 영역 첫 인라인 조회 ────────────────────────────────

    private Inline? GetFirstInlineInSelection()
    {
        var sel = _editor.Selection;
        if (sel.IsEmpty)
            return sel.Start.Parent as Inline;

        var ptr = sel.Start;
        while (ptr != null && ptr.CompareTo(sel.End) < 0)
        {
            if (ptr.GetPointerContext(LogicalDirection.Forward) == TextPointerContext.ElementStart)
            {
                var elem = ptr.GetAdjacentElement(LogicalDirection.Forward);
                if (elem is Run or InlineUIContainer)
                    return (Inline)elem;
            }
            ptr = ptr.GetNextContextPosition(LogicalDirection.Forward);
        }
        return null;
    }

    // ── 색상 유틸 ────────────────────────────────────────────────

    private static bool TryParseColor(string? text, out Color color)
    {
        color = default;
        if (string.IsNullOrWhiteSpace(text)) return false;
        var t = text.Trim();
        if (!t.StartsWith('#')) t = '#' + t;
        try
        {
            color = (Color)ColorConverter.ConvertFromString(t);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static string ToHex(Color c) => $"#{c.R:X2}{c.G:X2}{c.B:X2}";
}

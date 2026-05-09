using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using PolyDonky.App.Services;

namespace PolyDonky.App.Views;

/// <summary>
/// 글자 서식 다이얼로그.
/// <list type="bullet">
/// <item>RichTextBox 모드: 선택 영역의 서식을 읽어 초기화, OK 시 선택 영역에 일괄 적용.</item>
/// <item>단독 모드(RunStyle 생성자): 에디터 없이 스타일 객체만 편집. ResultStyle 로 결과 반환.</item>
/// </list>
/// </summary>
public partial class CharFormatWindow : Window
{
    private readonly RichTextBox? _editor;
    private readonly PolyDonky.Core.RunStyle? _standaloneStyle;
    private bool _suppressPreview;

    /// <summary>RichTextBox 선택 영역 편집 모드.</summary>
    public CharFormatWindow(RichTextBox editor)
    {
        _editor = editor;
        InitializeComponent();
        PopulateFontFamilies();
        Loaded += (_, _) => LoadCurrentFormatting();
    }

    /// <summary>단독 모드 — 에디터 없이 RunStyle 직접 편집. OK 후 ResultStyle 로 결과 읽기.</summary>
    public CharFormatWindow(PolyDonky.Core.RunStyle initial)
    {
        _standaloneStyle = CloneRunStyle(initial);
        InitializeComponent();
        PopulateFontFamilies();
        Loaded += (_, _) => LoadCurrentFormatting();
    }

    /// <summary>단독 모드에서 OK 후 편집 결과를 담는다.</summary>
    public PolyDonky.Core.RunStyle? ResultStyle { get; private set; }

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
            if (_standaloneStyle is { } s)
            {
                // 단독 모드: RunStyle 에서 직접 읽음
                if (!string.IsNullOrEmpty(s.FontFamily))  CboFont.Text    = s.FontFamily;
                TxtSize.Text          = s.FontSizePt.ToString("0.#");
                ChkBold.IsChecked     = s.Bold;
                ChkItalic.IsChecked   = s.Italic;
                ChkUnderline.IsChecked    = s.Underline;
                ChkStrikethrough.IsChecked = s.Strikethrough;
                ChkOverline.IsChecked     = s.Overline;
                ChkSuperscript.IsChecked  = s.Superscript;
                ChkSubscript.IsChecked    = s.Subscript;
                if (s.Foreground is { } fg)
                {
                    TxtFgColor.Text     = $"#{fg.R:X2}{fg.G:X2}{fg.B:X2}";
                    FgSwatch.Background = new SolidColorBrush(Color.FromArgb(fg.A, fg.R, fg.G, fg.B));
                }
                if (s.Background is { } bg)
                {
                    TxtBgColor.Text     = $"#{bg.R:X2}{bg.G:X2}{bg.B:X2}";
                    BgSwatch.Background = new SolidColorBrush(Color.FromArgb(bg.A, bg.R, bg.G, bg.B));
                }
                TxtWidthPercent.Text  = s.WidthPercent.ToString("0.#");
                TxtLetterSpacing.Text = s.LetterSpacingPx.ToString("0.##");
            }
            else
            {
                // RichTextBox 모드: 선택 영역에서 읽음
                var sel = _editor!.Selection;

                var fontFamilyVal = sel.GetPropertyValue(TextElement.FontFamilyProperty);
                if (fontFamilyVal is FontFamily ff)
                    CboFont.Text = ff.Source;

                var fontSizeVal = sel.GetPropertyValue(TextElement.FontSizeProperty);
                if (fontSizeVal is double dips)
                    TxtSize.Text = FlowDocumentBuilder.DipToPt(dips).ToString("0.#");

                var weightVal = sel.GetPropertyValue(TextElement.FontWeightProperty);
                ChkBold.IsChecked = weightVal == DependencyProperty.UnsetValue
                    ? null
                    : (bool?)(weightVal is FontWeight fw && fw.ToOpenTypeWeight() >= FontWeights.Bold.ToOpenTypeWeight());

                var styleVal = sel.GetPropertyValue(TextElement.FontStyleProperty);
                ChkItalic.IsChecked = styleVal == DependencyProperty.UnsetValue
                    ? null
                    : (bool?)(styleVal is FontStyle fs && fs == FontStyles.Italic);

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

                var fgVal = sel.GetPropertyValue(TextElement.ForegroundProperty);
                if (fgVal is SolidColorBrush fgBrush)
                {
                    TxtFgColor.Text   = ToHex(fgBrush.Color);
                    FgSwatch.Background = new SolidColorBrush(fgBrush.Color);
                }

                var bgVal = sel.GetPropertyValue(TextElement.BackgroundProperty);
                if (bgVal is SolidColorBrush bgBrush)
                {
                    TxtBgColor.Text   = ToHex(bgBrush.Color);
                    BgSwatch.Background = new SolidColorBrush(bgBrush.Color);
                }

                TxtWidthPercent.Text  = "100";
                TxtLetterSpacing.Text = "0";
                var firstInline = GetFirstInlineInSelection();
                var taggedRun = firstInline switch
                {
                    Run r             => r.Tag as PolyDonky.Core.Run,
                    Span sp           => sp.Tag as PolyDonky.Core.Run,
                    InlineUIContainer iuc => iuc.Tag as PolyDonky.Core.Run,
                    _                 => null,
                };
                if (taggedRun != null)
                {
                    if (Math.Abs(taggedRun.Style.WidthPercent - 100) > 0.1)
                        TxtWidthPercent.Text = taggedRun.Style.WidthPercent.ToString("0.#");
                    if (Math.Abs(taggedRun.Style.LetterSpacingPx) > 0.01)
                        TxtLetterSpacing.Text = taggedRun.Style.LetterSpacingPx.ToString("0.##");
                }
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
        => PickColor(TxtFgColor);

    private void OnBgSwatchClick(object sender, MouseButtonEventArgs e)
        => PickColor(TxtBgColor);

    private void OnFgPickerClick(object sender, RoutedEventArgs e)
        => PickColor(TxtFgColor);

    private void OnBgPickerClick(object sender, RoutedEventArgs e)
        => PickColor(TxtBgColor);

    private void PickColor(TextBox target)
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

        // 크기
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

        // 글자폭: LayoutTransform
        if (double.TryParse(TxtWidthPercent.Text, out var wp) && wp >= 1)
            PreviewText.LayoutTransform = new ScaleTransform(wp / 100.0, 1.0);
        else
            PreviewText.LayoutTransform = Transform.Identity;

        // 자간: 미리보기 TextBlock 에 직접 표현할 수단이 없으므로 ToolTip 으로 표시
        if (double.TryParse(TxtLetterSpacing.Text, out var ls) && Math.Abs(ls) > 0.01)
            PreviewText.ToolTip = $"자간 {ls:0.##}px";
        else
            PreviewText.ToolTip = null;
    }

    // ── OK / Cancel ──────────────────────────────────────────────

    private void OnOk(object sender, RoutedEventArgs e)
    {
        if (_standaloneStyle != null)
        {
            // 단독 모드: UI 값을 _standaloneStyle 에 기록 후 복사본을 ResultStyle 로 노출.
            WriteToRunStyle(_standaloneStyle);
            ResultStyle = CloneRunStyle(_standaloneStyle);
        }
        else
        {
            ApplyToSelection();
            try
            {
                var end = _editor!.Selection.End;
                if (end != null) _editor.Selection.Select(end, end);
            }
            catch { }
        }
        DialogResult = true;
    }

    private void OnCancel(object sender, RoutedEventArgs e) => DialogResult = false;

    // ── UI → RunStyle 쓰기 (단독 모드용) ────────────────────────

    private void WriteToRunStyle(PolyDonky.Core.RunStyle s)
    {
        s.FontFamily   = CboFont.Text?.Trim();
        if (double.TryParse(TxtSize.Text, out var pt) && pt >= 1)  s.FontSizePt = pt;
        s.Bold         = ChkBold.IsChecked == true;
        s.Italic       = ChkItalic.IsChecked == true;
        s.Underline    = ChkUnderline.IsChecked == true;
        s.Strikethrough = ChkStrikethrough.IsChecked == true;
        s.Overline     = ChkOverline.IsChecked == true;
        s.Superscript  = ChkSuperscript.IsChecked == true;
        s.Subscript    = ChkSubscript.IsChecked == true;
        s.Foreground   = TryParseColor(TxtFgColor.Text, out var fg) ? new PolyDonky.Core.Color(fg.R, fg.G, fg.B, fg.A) : null;
        s.Background   = TryParseColor(TxtBgColor.Text, out var bg) ? new PolyDonky.Core.Color(bg.R, bg.G, bg.B, bg.A) : null;
        if (double.TryParse(TxtWidthPercent.Text, out var wp) && wp >= 1) s.WidthPercent = wp;
        if (double.TryParse(TxtLetterSpacing.Text, out var ls)) s.LetterSpacingPx = ls;
    }

    private static PolyDonky.Core.RunStyle CloneRunStyle(PolyDonky.Core.RunStyle s) => s.Clone();

    // ── 선택 영역 서식 적용 (RichTextBox 모드) ───────────────────

    private void ApplyToSelection()
    {
        var sel = _editor!.Selection;

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

        // 배경색
        if (string.IsNullOrWhiteSpace(TxtBgColor.Text))
            sel.ApplyPropertyValue(TextElement.BackgroundProperty, null);
        else if (TryParseColor(TxtBgColor.Text, out var bgColor))
            sel.ApplyPropertyValue(TextElement.BackgroundProperty, new SolidColorBrush(bgColor));

        // 글자폭 / 자간: InlineUIContainer 재구성
        double widthPct = double.TryParse(TxtWidthPercent.Text, out var wp) && wp >= 1 ? wp : 100;
        double lsPx     = double.TryParse(TxtLetterSpacing.Text, out var ls) ? ls : 0;
        ApplyTypographicProps(widthPct, lsPx);
    }

    // ── 글자폭·자간 적용 (Run ↔ InlineUIContainer 재구성) ────────

    private void ApplyTypographicProps(double widthPercent, double letterSpacingPx)
    {
        var sel = _editor!.Selection;
        if (sel.IsEmpty) return;

        // 글자폭/자간이 모두 기본값(100%, 0px) 이면 ReplaceInline 으로 인라인을 교체할
        // 이유가 없다. 직전 ApplyPropertyValue 가 Wpf.Run 에 적용한 색·폰트·볼드 결과를
        // Tag 기반 stale RunStyle 로 다시 빌드해 덮어쓰는 회귀를 방지 (글상자 로드/복사
        // 케이스에서 모든 글자속성 변경이 시각적으로 반영되지 않던 원인).
        bool widthChanged   = Math.Abs(widthPercent - 100) > 0.5;
        bool spacingChanged = Math.Abs(letterSpacingPx)    > 0.01;
        if (!widthChanged && !spacingChanged) return;

        var inlines = CollectLeafInlines(sel);

        foreach (var inline in inlines)
        {
            // 수식·이모지 IUC 는 글자폭/자간 재구성 대상에서 제외 (Image 로 렌더링된 객체는 변형 불필요)
            if (inline is InlineUIContainer { Tag: PolyDonky.Core.Run { LatexSource: { Length: > 0 } } })
                continue;
            if (inline is InlineUIContainer { Tag: PolyDonky.Core.Run { EmojiKey: { Length: > 0 } } })
                continue;

            // PolyDonky Run 추출. Wpf.Run 의 경우 Tag 가 있어도 항상 현재 Wpf 속성에서
            // 새로 추출한다 — 직전 ApplyPropertyValue 결과가 stale Tag.Style 로 덮이지 않도록.
            // Span / IUC 는 per-char 구조라 Tag.pr 을 그대로 사용 (자체 형식이 의미 있음).
            PolyDonky.Core.Run polyRun = inline switch
            {
                Run r                                   => ExtractPolyRun(r),
                Span sp when sp.Tag is PolyDonky.Core.Run pr => pr,
                InlineUIContainer { Tag: PolyDonky.Core.Run pr } => pr,
                InlineUIContainer iuc                   => ExtractPolyRunFromContainer(iuc),
                _                                       => null!,
            };
            if (polyRun is null) continue;

            polyRun.Style.WidthPercent     = widthPercent;
            polyRun.Style.LetterSpacingPx  = letterSpacingPx;

            var newInline = FlowDocumentBuilder.BuildInline(polyRun);
            ReplaceInline(inline, newInline);
        }
    }

    /// <summary>
    /// 선택 영역의 모든 Run / Span / InlineUIContainer 를 중복 없이 수집.
    /// per-char IUC 가 Span(Tag=Run) 안에 들어 있는 경우는 IUC 들을 묶어 Span 단위로 한 번만 추가
    /// — 그래야 ApplyTypographicProps 가 Span 한 덩어리를 새로운 Span 으로 통째 교체할 수 있다.
    /// sel.Start 가 Run/IUC 내부(Text 컨텍스트)에 있을 때 ElementStart 가 안 보이므로
    /// 루프 전에 sel.Start.Parent 를 먼저 확인한다.
    /// </summary>
    private static System.Collections.Generic.List<Inline> CollectLeafInlines(TextSelection sel)
    {
        var result = new System.Collections.Generic.List<Inline>();

        void TryAdd(object? obj)
        {
            // per-char IUC 가 PolyDonky.Run-tagged Span 안에 있으면 Span 자체를 단위로 잡는다.
            if (obj is InlineUIContainer iuc
                && iuc.Parent is Span parentSpan
                && parentSpan.Tag is PolyDonky.Core.Run
                && ReferenceEquals(parentSpan.Tag, iuc.Tag))
            {
                if (!result.Contains(parentSpan)) result.Add(parentSpan);
                return;
            }
            if (obj is Run r && !result.Contains(r)) result.Add(r);
            else if (obj is InlineUIContainer i && !result.Contains(i)) result.Add(i);
            else if (obj is Span s && s.Tag is PolyDonky.Core.Run && !result.Contains(s)) result.Add(s);
        }

        TryAdd(sel.Start.Parent);

        var ptr = sel.Start;
        while (ptr != null && ptr.CompareTo(sel.End) < 0)
        {
            if (ptr.GetPointerContext(LogicalDirection.Forward) == TextPointerContext.ElementStart)
                TryAdd(ptr.GetAdjacentElement(LogicalDirection.Forward));
            ptr = ptr.GetNextContextPosition(LogicalDirection.Forward);
        }

        return result;
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

    private static PolyDonky.Core.Run ExtractPolyRun(Run wpfRun)
    {
        var s = new PolyDonky.Core.RunStyle
        {
            FontFamily   = wpfRun.FontFamily?.Source,
            FontSizePt   = FlowDocumentBuilder.DipToPt(wpfRun.FontSize),
            Bold         = wpfRun.FontWeight.ToOpenTypeWeight() >= FontWeights.Bold.ToOpenTypeWeight(),
            Italic       = wpfRun.FontStyle == FontStyles.Italic,
            WidthPercent = 100,
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
            s.Foreground = new PolyDonky.Core.Color(fg.Color.R, fg.Color.G, fg.Color.B, fg.Color.A);
        if (wpfRun.Background is SolidColorBrush bg)
            s.Background = new PolyDonky.Core.Color(bg.Color.R, bg.Color.G, bg.Color.B, bg.Color.A);

        return new PolyDonky.Core.Run { Text = wpfRun.Text, Style = s };
    }

    private static PolyDonky.Core.Run ExtractPolyRunFromContainer(InlineUIContainer iuc)
    {
        if (iuc.Tag is PolyDonky.Core.Run pr) return pr;

        var s = new PolyDonky.Core.RunStyle { WidthPercent = 100 };
        string text = string.Empty;

        if (iuc.Child is StackPanel panel)
        {
            var sb = new System.Text.StringBuilder();
            TextBlock? first = null;
            foreach (var child in panel.Children)
            {
                if (child is TextBlock ctb) { sb.Append(ctb.Text); first ??= ctb; }
            }
            text = sb.ToString();
            if (first != null)
            {
                s.LetterSpacingPx = first.Margin.Right;
                if (first.LayoutTransform is ScaleTransform st) s.WidthPercent = st.ScaleX * 100.0;
                CopyTextBlockStyle(first, s);
            }
        }
        else if (iuc.Child is TextBlock tb)
        {
            text = tb.Text;
            if (tb.LayoutTransform is ScaleTransform st) s.WidthPercent = st.ScaleX * 100.0;
            CopyTextBlockStyle(tb, s);
        }

        return new PolyDonky.Core.Run { Text = text, Style = s };
    }

    private static void CopyTextBlockStyle(TextBlock tb, PolyDonky.Core.RunStyle s)
    {
        s.FontFamily = tb.FontFamily?.Source;
        s.FontSizePt = FlowDocumentBuilder.DipToPt(tb.FontSize);
        s.Bold       = tb.FontWeight.ToOpenTypeWeight() >= FontWeights.Bold.ToOpenTypeWeight();
        s.Italic     = tb.FontStyle == FontStyles.Italic;
        if (tb.Foreground is SolidColorBrush fg)
            s.Foreground = new PolyDonky.Core.Color(fg.Color.R, fg.Color.G, fg.Color.B, fg.Color.A);
        if (tb.Background is SolidColorBrush bg)
            s.Background = new PolyDonky.Core.Color(bg.Color.R, bg.Color.G, bg.Color.B, bg.Color.A);
    }

    // ── 선택 영역 첫 인라인 조회 ────────────────────────────────

    private Inline? GetFirstInlineInSelection()
    {
        var sel = _editor!.Selection;

        // per-char IUC 안에 캐럿이 있으면 부모 Span 을 우선 반환 — 자간/글자폭 값을 그쪽 Tag 에서 읽기.
        Inline? Lift(object? elem)
        {
            if (elem is InlineUIContainer iuc
                && iuc.Parent is Span parent
                && parent.Tag is PolyDonky.Core.Run
                && ReferenceEquals(parent.Tag, iuc.Tag))
                return parent;
            return elem as Inline;
        }

        if (sel.Start.Parent is Run or InlineUIContainer or Span)
            return Lift(sel.Start.Parent);

        var ptr = sel.Start;
        while (ptr != null && ptr.CompareTo(sel.End) < 0)
        {
            if (ptr.GetPointerContext(LogicalDirection.Forward) == TextPointerContext.ElementStart)
            {
                var elem = ptr.GetAdjacentElement(LogicalDirection.Forward);
                if (elem is Run or InlineUIContainer or Span)
                    return Lift(elem);
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

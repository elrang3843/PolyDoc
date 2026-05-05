using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using WpfDoc = System.Windows.Documents;
using System.Windows.Shapes;
using PolyDonky.Core;
using WpfMedia = System.Windows.Media;

namespace PolyDonky.App.Views;

/// <summary>
/// 페이지(용지·여백·레이아웃·머리말/꼬리말) 설정 다이얼로그.
///</summary>
public partial class PageFormatWindow : Window
{
    private PageSettings _settings;
    private bool _suppress;
    private readonly int _initialTab;

    public PageSettings ResultSettings { get; private set; } = new();

    public PageFormatWindow(PageSettings current, int initialTab = 0)
    {
        _settings   = Clone(current);
        _initialTab = initialTab;
        InitializeComponent();
        Loaded += OnLoaded;
    }

    // ── 초기화 ──────────────────────────────────────────────────

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        PopulateSizeCombo();
        InitHFToolbar();
        LoadUI();
        if (_initialTab > 0 && _initialTab < MainTabControl.Items.Count)
            MainTabControl.SelectedIndex = _initialTab;
    }

    private void PopulateSizeCombo()
    {
        Add(PaperSizeKind.Custom, "사용자 정의");
        AddHeader("— ISO A 시리즈 —");
        Add(PaperSizeKind.A0, "A0  (841×1189 mm)");
        Add(PaperSizeKind.A1, "A1  (594×841 mm)");
        Add(PaperSizeKind.A2, "A2  (420×594 mm)");
        Add(PaperSizeKind.A3, "A3  (297×420 mm)");
        Add(PaperSizeKind.A4, "A4  (210×297 mm)");
        Add(PaperSizeKind.A5, "A5  (148×210 mm)");
        Add(PaperSizeKind.A6, "A6  (105×148 mm)");
        Add(PaperSizeKind.A7, "A7  (74×105 mm)");
        AddHeader("— ISO B 시리즈 —");
        Add(PaperSizeKind.B4_ISO, "B4 ISO  (250×353 mm)");
        Add(PaperSizeKind.B5_ISO, "B5 ISO  (176×250 mm)");
        Add(PaperSizeKind.B6_ISO, "B6 ISO  (125×176 mm)");
        AddHeader("— JIS/KS B 시리즈 —");
        Add(PaperSizeKind.B4_JIS, "B4 JIS/KS  (257×364 mm)");
        Add(PaperSizeKind.B5_JIS, "B5 JIS/KS  (182×257 mm)");
        Add(PaperSizeKind.B6_JIS, "B6 JIS/KS  (128×182 mm)");
        AddHeader("— 미국·국제 표준 —");
        Add(PaperSizeKind.Letter,    "Letter  (215.9×279.4 mm, 8.5\"×11\")");
        Add(PaperSizeKind.Legal,     "Legal  (215.9×355.6 mm, 8.5\"×14\")");
        Add(PaperSizeKind.Ledger,    "Ledger  (431.8×279.4 mm, 17\"×11\")");
        Add(PaperSizeKind.Tabloid,   "Tabloid  (279.4×431.8 mm, 11\"×17\")");
        Add(PaperSizeKind.Statement, "Statement  (139.7×215.9 mm, 5.5\"×8.5\")");
        Add(PaperSizeKind.Executive, "Executive  (184.2×266.7 mm, 7.25\"×10.5\")");
        Add(PaperSizeKind.Folio,     "Folio  (215.9×330.2 mm, 8.5\"×13\")");
        AddHeader("— 신문 판형 —");
        Add(PaperSizeKind.Broadsheet,       "대판 Broadsheet  (381×585 mm)");
        Add(PaperSizeKind.Berliner,         "베를리너  (315×470 mm)");
        Add(PaperSizeKind.Compact,          "콤팩트  (282×380 mm)");
        Add(PaperSizeKind.KoreanBroadsheet, "한국 신문 대판  (393×546 mm)");
        AddHeader("— 한국 도서 판형 —");
        Add(PaperSizeKind.SinGukPan,  "신국판  (152×225 mm)");
        Add(PaperSizeKind.GukPan,     "국판  (148×210 mm)");
        Add(PaperSizeKind.Pan46Bae,   "4×6배판  (188×257 mm)");
        Add(PaperSizeKind.CrownPan,   "크라운판  (176×248 mm)");
        Add(PaperSizeKind.Pan46,      "46판  (127×188 mm)");
        Add(PaperSizeKind.MassMarket, "문고판  (106×175 mm)");
        AddHeader("— 국제 도서 판형 —");
        Add(PaperSizeKind.BookRoyal,       "Royal  (156×234 mm)");
        Add(PaperSizeKind.BookDemy,        "Demy  (138×216 mm)");
        Add(PaperSizeKind.BookTrade,       "Trade 6\"×9\"  (152×229 mm)");
        Add(PaperSizeKind.BookQuarto,      "Quarto  (189×246 mm)");
        Add(PaperSizeKind.BookLargeFormat, "Large Format  (216×280 mm)");

        // 현재 SizeKind 선택
        SelectSizeKind(_settings.SizeKind);
    }

    private void Add(PaperSizeKind kind, string label)
        => CboSize.Items.Add(new ComboBoxItem { Content = label, Tag = kind });

    private void AddHeader(string text)
        => CboSize.Items.Add(new ComboBoxItem
        {
            Content    = text,
            IsEnabled  = false,
            Foreground = SystemColors.GrayTextBrush,
            FontStyle  = FontStyles.Italic,
        });

    private void SelectSizeKind(PaperSizeKind kind)
    {
        for (int i = 0; i < CboSize.Items.Count; i++)
        {
            if (CboSize.Items[i] is ComboBoxItem ci && ci.Tag is PaperSizeKind k && k == kind)
            {
                CboSize.SelectedIndex = i;
                return;
            }
        }
        CboSize.SelectedIndex = 0; // Custom fallback
    }

    // ── UI 로드 ──────────────────────────────────────────────────

    private void LoadUI()
    {
        _suppress = true;
        try
        {
            // 용지 치수
            TxtWidth.Text  = _settings.WidthMm.ToString("0.##");
            TxtHeight.Text = _settings.HeightMm.ToString("0.##");

            // 방향
            RbPortrait.IsChecked  = _settings.Orientation == PageOrientation.Portrait;
            RbLandscape.IsChecked = _settings.Orientation == PageOrientation.Landscape;

            // 글자 방향
            CboTextOrientation.SelectedIndex = (int)_settings.TextOrientation;
            CboTextProgression.SelectedIndex = (int)_settings.TextProgression;

            // 용지 색
            bool isDefault = string.IsNullOrEmpty(_settings.PaperColor);
            ChkDefaultColor.IsChecked = isDefault;
            TxtPaperColor.IsEnabled   = !isDefault;
            PaperColorSwatch.IsEnabled = !isDefault;
            TxtPaperColor.Text        = _settings.PaperColor ?? "";
            UpdateColorSwatch();

            // 여백
            TxtMarginTop.Text    = _settings.MarginTopMm.ToString("0.##");
            TxtMarginBottom.Text = _settings.MarginBottomMm.ToString("0.##");
            TxtMarginLeft.Text   = _settings.MarginLeftMm.ToString("0.##");
            TxtMarginRight.Text  = _settings.MarginRightMm.ToString("0.##");
            TxtMarginHeader.Text = _settings.MarginHeaderMm.ToString("0.##");
            TxtMarginFooter.Text = _settings.MarginFooterMm.ToString("0.##");

            // 머리말/꼬리말 내용 — RichTextBox 로드 (Loaded 이후이므로 컨트롤이 존재)
            LoadSlotToRtb(_settings.Header.Left,   RtbHeaderLeft);
            LoadSlotToRtb(_settings.Header.Center, RtbHeaderCenter);
            LoadSlotToRtb(_settings.Header.Right,  RtbHeaderRight);
            LoadSlotToRtb(_settings.Footer.Left,   RtbFooterLeft);
            LoadSlotToRtb(_settings.Footer.Center, RtbFooterCenter);
            LoadSlotToRtb(_settings.Footer.Right,  RtbFooterRight);

            // 여백 안내선
            ChkShowMarginGuides.IsChecked = _settings.ShowMarginGuides;

            // 레이아웃
            TxtColumns.Text         = _settings.ColumnCount.ToString();
            TxtColumnGap.Text       = _settings.ColumnGapMm.ToString("0.##");
            TxtPageNumberStart.Text = _settings.PageNumberStart.ToString();

            // 단 구분선
            ChkColumnDividerVisible.IsChecked   = _settings.ColumnDividerVisible;
            TxtColumnDividerColor.Text          = _settings.ColumnDividerColor ?? "#888888";
            TxtColumnDividerThickness.Text      = _settings.ColumnDividerThicknessPt.ToString("0.##");
            CboColumnDividerStyle.SelectedIndex = (int)_settings.ColumnDividerStyle;
            UpdateDividerColorSwatch();

            // Custom 패널 가시성 — 선택 연동은 OnSizeChanged 에서
            PanelCustomSize.IsEnabled = _settings.SizeKind == PaperSizeKind.Custom;
        }
        finally
        {
            _suppress = false;
        }
        UpdatePreview();
    }

    // ── 용지 크기 변경 ───────────────────────────────────────────

    private void OnSizeChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_suppress) return;
        if (CboSize.SelectedItem is not ComboBoxItem ci || ci.Tag is not PaperSizeKind kind) return;

        _settings.ApplySizeKind(kind);
        bool isCustom = kind == PaperSizeKind.Custom;
        PanelCustomSize.IsEnabled = isCustom;

        _suppress = true;
        TxtWidth.Text  = _settings.WidthMm.ToString("0.##");
        TxtHeight.Text = _settings.HeightMm.ToString("0.##");
        _suppress = false;
        UpdatePreview();
    }

    private void OnDimensionChanged(object sender, TextChangedEventArgs e)
    {
        if (_suppress) return;
        if (double.TryParse(TxtWidth.Text,  out double w) && w > 0) _settings.WidthMm  = w;
        if (double.TryParse(TxtHeight.Text, out double h) && h > 0) _settings.HeightMm = h;
        _settings.SizeKind = PaperSizeKind.Custom;
        SelectSizeKind(PaperSizeKind.Custom);
        UpdatePreview();
    }

    private void OnOrientationChanged(object sender, RoutedEventArgs e)
    {
        if (_suppress) return;
        if (RbLandscape is null) return; // InitializeComponent 중 조기 발화 방어
        _settings.Orientation = RbLandscape.IsChecked == true
            ? PageOrientation.Landscape
            : PageOrientation.Portrait;
        UpdatePreview();
    }

    private void OnTextDirectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_suppress) return;
        if (CboTextOrientation is null || CboTextProgression is null) return;
        _settings.TextOrientation = (TextOrientation)System.Math.Clamp(CboTextOrientation.SelectedIndex, 0, 1);
        _settings.TextProgression = (TextProgression)System.Math.Clamp(CboTextProgression.SelectedIndex, 0, 1);
    }

    // ── 용지 색상 ────────────────────────────────────────────────

    private void OnDefaultColorChanged(object sender, RoutedEventArgs e)
    {
        if (_suppress) return;
        bool isDefault = ChkDefaultColor.IsChecked == true;
        TxtPaperColor.IsEnabled    = !isDefault;
        PaperColorSwatch.IsEnabled = !isDefault;
        if (isDefault)
        {
            _settings.PaperColor = null;
            PaperColorSwatch.Background = null;
        }
        else if (TryParseWpfColor(TxtPaperColor.Text, out var c))
        {
            _settings.PaperColor = NormalizeHex(TxtPaperColor.Text.Trim());
            PaperColorSwatch.Background = new WpfMedia.SolidColorBrush(c);
        }
        UpdatePreview();
    }

    private void OnPaperColorSwatchClick(object sender, MouseButtonEventArgs e)
    {
        using var dlg = new System.Windows.Forms.ColorDialog { FullOpen = true, AnyColor = true };
        if (TryParseWpfColor(TxtPaperColor.Text, out var cur))
            dlg.Color = System.Drawing.Color.FromArgb(cur.A, cur.R, cur.G, cur.B);
        if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            var p = dlg.Color;
            _suppress = true;
            TxtPaperColor.Text = $"#{p.R:X2}{p.G:X2}{p.B:X2}";
            _suppress = false;
            _settings.PaperColor = TxtPaperColor.Text;
            UpdateColorSwatch();
            UpdatePreview();
        }
    }

    private void OnPaperColorChanged(object sender, TextChangedEventArgs e)
    {
        if (_suppress) return;
        var hex = TxtPaperColor.Text.Trim();
        if (TryParseWpfColor(hex, out _))
        {
            _settings.PaperColor = NormalizeHex(hex);
            UpdateColorSwatch();
        }
        else
        {
            _settings.PaperColor = null;
        }
        UpdatePreview();
    }

    private void UpdateColorSwatch()
    {
        if (TryParseWpfColor(_settings.PaperColor, out var c))
            PaperColorSwatch.Background = new WpfMedia.SolidColorBrush(c);
        else
            PaperColorSwatch.Background = new WpfMedia.SolidColorBrush(WpfMedia.Colors.White);
    }

    // ── 여백 ─────────────────────────────────────────────────────

    private void OnMarginChanged(object sender, TextChangedEventArgs e)
    {
        if (_suppress) return;
        TrySetMm(TxtMarginTop,    v => _settings.MarginTopMm    = v);
        TrySetMm(TxtMarginBottom, v => _settings.MarginBottomMm = v);
        TrySetMm(TxtMarginLeft,   v => _settings.MarginLeftMm   = v);
        TrySetMm(TxtMarginRight,  v => _settings.MarginRightMm  = v);
        TrySetMm(TxtMarginHeader, v => _settings.MarginHeaderMm = v);
        TrySetMm(TxtMarginFooter, v => _settings.MarginFooterMm = v);
        UpdatePreview();
    }

    private void OnShowMarginGuidesChanged(object sender, RoutedEventArgs e)
    {
        if (_suppress) return;
        _settings.ShowMarginGuides = ChkShowMarginGuides.IsChecked == true;
    }

    // ── 레이아웃 ─────────────────────────────────────────────────

    private void OnLayoutChanged(object sender, TextChangedEventArgs e)
        => RecomputeLayout();

    private void OnDividerVisibleChanged(object sender, RoutedEventArgs e)
        => RecomputeLayout();

    private void RecomputeLayout()
    {
        if (_suppress) return;
        if (int.TryParse(TxtColumns.Text, out int cols) && cols >= 1)
            _settings.ColumnCount = cols;
        TrySetMm(TxtColumnGap, v => _settings.ColumnGapMm = v);
        if (int.TryParse(TxtPageNumberStart.Text, out int pg) && pg >= 0)
            _settings.PageNumberStart = pg;

        _settings.ColumnDividerVisible = ChkColumnDividerVisible.IsChecked == true;
        if (double.TryParse(TxtColumnDividerThickness.Text, out double thk) && thk > 0)
            _settings.ColumnDividerThicknessPt = thk;

        UpdatePreview();
    }

    private void OnDividerColorChanged(object sender, TextChangedEventArgs e)
    {
        if (_suppress) return;
        var hex = TxtColumnDividerColor.Text.Trim();
        if (TryParseWpfColor(hex, out _))
            _settings.ColumnDividerColor = NormalizeHex(hex);
        UpdateDividerColorSwatch();
        UpdatePreview();
    }

    private void OnDividerColorPickClick(object sender, RoutedEventArgs e)
    {
        using var dlg = new System.Windows.Forms.ColorDialog { FullOpen = true, AnyColor = true };
        if (TryParseWpfColor(TxtColumnDividerColor.Text, out var cur))
            dlg.Color = System.Drawing.Color.FromArgb(cur.A, cur.R, cur.G, cur.B);
        if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            var p = dlg.Color;
            _suppress = true;
            TxtColumnDividerColor.Text = $"#{p.R:X2}{p.G:X2}{p.B:X2}";
            _suppress = false;
            _settings.ColumnDividerColor = TxtColumnDividerColor.Text;
            UpdateDividerColorSwatch();
            UpdatePreview();
        }
    }

    private void OnDividerStyleChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_suppress) return;
        _settings.ColumnDividerStyle = (ColumnDividerStyle)System.Math.Clamp(
            CboColumnDividerStyle.SelectedIndex, 0, 3);
        UpdatePreview();
    }

    private void UpdateDividerColorSwatch()
    {
        if (TryParseWpfColor(_settings.ColumnDividerColor, out var c))
            BtnColumnDividerColorPick.Background = new WpfMedia.SolidColorBrush(c);
        else
            BtnColumnDividerColorPick.Background = new WpfMedia.SolidColorBrush(WpfMedia.Colors.Gray);
    }

    // ── 머리말/꼬리말 툴바 ──────────────────────────────────────

    private RichTextBox? _focusedRtb;
    private WpfMedia.Color _hfFontColor = WpfMedia.Colors.Black;
    private bool _suppressFontSizeChange;
    private bool _suppressFontFamilyChange;

    private void InitHFToolbar()
    {
        foreach (var ff in Fonts.SystemFontFamilies
                                .Select(f => f.Source)
                                .OrderBy(n => n, StringComparer.CurrentCultureIgnoreCase))
            CboHFFontFamily.Items.Add(ff);
        CboHFFontFamily.Text = "맑은 고딕";

        var sizes = new[] { "6", "7", "8", "9", "10", "10.5", "11", "12",
                            "14", "16", "18", "20", "22", "24", "28", "36", "48", "72" };
        foreach (var s in sizes)
            CboHFFontSize.Items.Add(s);
        CboHFFontSize.Text = "11";

        var tokens = new (string display, string token)[]
        {
            ("{PAGE} — 페이지 번호",     "{PAGE}"),
            ("{NUMPAGES} — 전체 페이지", "{NUMPAGES}"),
            ("{DATE} — 날짜",           "{DATE}"),
            ("{TIME} — 시간",           "{TIME}"),
            ("{TITLE} — 제목",          "{TITLE}"),
            ("{AUTHOR} — 작성자",       "{AUTHOR}"),
            ("{FILENAME} — 파일명",     "{FILENAME}"),
        };
        foreach (var (display, token) in tokens)
            CboHFToken.Items.Add(new ComboBoxItem { Content = display, Tag = token });
    }

    private void OnRtbGotFocus(object sender, RoutedEventArgs e)
    {
        if (sender is RichTextBox rtb)
            _focusedRtb = rtb;
    }

    private void OnRtbSelectionChanged(object sender, RoutedEventArgs e)
    {
        if (sender is not RichTextBox rtb) return;
        var sel = rtb.Selection;

        var fw = sel.GetPropertyValue(WpfDoc.TextElement.FontWeightProperty);
        BtnHFBold.IsChecked = fw is FontWeight w && w == FontWeights.Bold;

        var fs = sel.GetPropertyValue(WpfDoc.TextElement.FontStyleProperty);
        BtnHFItalic.IsChecked = fs is FontStyle st && st == FontStyles.Italic;

        var td = sel.GetPropertyValue(WpfDoc.Inline.TextDecorationsProperty) as TextDecorationCollection;
        BtnHFUnderline.IsChecked  = td?.Any(d => d.Location == TextDecorationLocation.Underline)     == true;
        BtnHFStrike.IsChecked     = td?.Any(d => d.Location == TextDecorationLocation.Strikethrough) == true;

        var sz = sel.GetPropertyValue(WpfDoc.TextElement.FontSizeProperty);
        if (sz is double dip && !double.IsNaN(dip))
        {
            _suppressFontSizeChange = true;
            CboHFFontSize.Text = (dip * 72.0 / 96.0).ToString("0.##");
            _suppressFontSizeChange = false;
        }

        var ff = sel.GetPropertyValue(WpfDoc.TextElement.FontFamilyProperty) as FontFamily;
        if (ff is not null)
        {
            _suppressFontFamilyChange = true;
            CboHFFontFamily.Text = ff.Source;
            _suppressFontFamilyChange = false;
        }

        var align = sel.GetPropertyValue(WpfDoc.Block.TextAlignmentProperty);
        var ta = align is TextAlignment t ? t : TextAlignment.Left;
        BtnHFAlignLeft.IsChecked    = ta == TextAlignment.Left;
        BtnHFAlignCenter.IsChecked  = ta == TextAlignment.Center;
        BtnHFAlignRight.IsChecked   = ta == TextAlignment.Right;
        BtnHFAlignJustify.IsChecked = ta == TextAlignment.Justify;

        var fg = sel.GetPropertyValue(WpfDoc.TextElement.ForegroundProperty) as SolidColorBrush;
        if (fg is not null)
        {
            _hfFontColor = fg.Color;
            HFColorBar.Fill = new SolidColorBrush(_hfFontColor);
        }
    }

    private void OnHFFontFamilyChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_suppressFontFamilyChange || _focusedRtb is null) return;
        if (CboHFFontFamily.SelectedItem is string name && !string.IsNullOrWhiteSpace(name))
            ApplyFontFamily(name);
    }

    private void OnHFFontFamilyLostFocus(object sender, RoutedEventArgs e)
    {
        if (_focusedRtb is null) return;
        var name = CboHFFontFamily.Text;
        if (!string.IsNullOrWhiteSpace(name))
            ApplyFontFamily(name);
    }

    private void ApplyFontFamily(string name)
    {
        try
        {
            var ff = new FontFamily(name);
            _focusedRtb?.Selection.ApplyPropertyValue(
                WpfDoc.TextElement.FontFamilyProperty, ff);
            _focusedRtb?.Focus();
        }
        catch (ArgumentException) { }
    }

    private void OnHFAlignLeft(object sender, RoutedEventArgs e)
    {
        if (_focusedRtb is null) return;
        WpfDoc.EditingCommands.AlignLeft.Execute(null, _focusedRtb);
        _focusedRtb.Focus();
    }

    private void OnHFAlignCenter(object sender, RoutedEventArgs e)
    {
        if (_focusedRtb is null) return;
        WpfDoc.EditingCommands.AlignCenter.Execute(null, _focusedRtb);
        _focusedRtb.Focus();
    }

    private void OnHFAlignRight(object sender, RoutedEventArgs e)
    {
        if (_focusedRtb is null) return;
        WpfDoc.EditingCommands.AlignRight.Execute(null, _focusedRtb);
        _focusedRtb.Focus();
    }

    private void OnHFAlignJustify(object sender, RoutedEventArgs e)
    {
        if (_focusedRtb is null) return;
        WpfDoc.EditingCommands.AlignJustify.Execute(null, _focusedRtb);
        _focusedRtb.Focus();
    }

    private void OnHFBold(object sender, RoutedEventArgs e)
    {
        if (_focusedRtb is null) return;
        WpfDoc.EditingCommands.ToggleBold.Execute(null, _focusedRtb);
        _focusedRtb.Focus();
    }

    private void OnHFItalic(object sender, RoutedEventArgs e)
    {
        if (_focusedRtb is null) return;
        WpfDoc.EditingCommands.ToggleItalic.Execute(null, _focusedRtb);
        _focusedRtb.Focus();
    }

    private void OnHFUnderline(object sender, RoutedEventArgs e)
    {
        if (_focusedRtb is null) return;
        WpfDoc.EditingCommands.ToggleUnderline.Execute(null, _focusedRtb);
        _focusedRtb.Focus();
    }

    private void OnHFStrikethrough(object sender, RoutedEventArgs e)
    {
        if (_focusedRtb is null) return;
        var sel = _focusedRtb.Selection;
        var td = sel.GetPropertyValue(WpfDoc.Inline.TextDecorationsProperty) as TextDecorationCollection;
        bool hasStrike = td?.Any(d => d.Location == TextDecorationLocation.Strikethrough) == true;
        var newDec = new TextDecorationCollection(
            hasStrike ? (td!.Where(d => d.Location != TextDecorationLocation.Strikethrough))
                      : (td ?? Enumerable.Empty<TextDecoration>()).Append(TextDecorations.Strikethrough[0]));
        sel.ApplyPropertyValue(WpfDoc.Inline.TextDecorationsProperty, newDec);
        _focusedRtb.Focus();
    }

    private void OnHFFontSizeChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_suppressFontSizeChange || _focusedRtb is null) return;
        if (CboHFFontSize.SelectedItem is string s && double.TryParse(s, out double pt) && pt > 0)
            ApplyFontSize(pt);
    }

    private void OnHFFontSizeLostFocus(object sender, RoutedEventArgs e)
    {
        if (_focusedRtb is null) return;
        if (double.TryParse(CboHFFontSize.Text, out double pt) && pt > 0)
            ApplyFontSize(pt);
    }

    private void ApplyFontSize(double pt)
    {
        _focusedRtb?.Selection.ApplyPropertyValue(
            WpfDoc.TextElement.FontSizeProperty, pt * 96.0 / 72.0);
        _focusedRtb?.Focus();
    }

    private void OnHFFontColor(object sender, RoutedEventArgs e)
    {
        if (_focusedRtb is null) return;
        var dlg = new System.Windows.Forms.ColorDialog
        {
            Color            = System.Drawing.Color.FromArgb(_hfFontColor.A, _hfFontColor.R, _hfFontColor.G, _hfFontColor.B),
            FullOpen         = true,
            AllowFullOpen    = true,
        };
        if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            var c = dlg.Color;
            _hfFontColor    = WpfMedia.Color.FromArgb(c.A, c.R, c.G, c.B);
            HFColorBar.Fill = new SolidColorBrush(_hfFontColor);
            _focusedRtb.Selection.ApplyPropertyValue(
                WpfDoc.TextElement.ForegroundProperty, new SolidColorBrush(_hfFontColor));
            _focusedRtb.Focus();
        }
    }

    private void OnInsertToken(object sender, RoutedEventArgs e)
    {
        if (_focusedRtb is null) return;
        if (CboHFToken.SelectedItem is ComboBoxItem ci && ci.Tag is string token)
        {
            _focusedRtb.Selection.Text = token;
            _focusedRtb.CaretPosition  = _focusedRtb.Selection.End;
            _focusedRtb.Focus();
        }
    }

    // ── 슬롯 ↔ FlowDocument 변환 ─────────────────────────────────

    private static void LoadSlotToRtb(HeaderFooterSlot slot, RichTextBox rtb)
    {
        var fd = new WpfDoc.FlowDocument { PagePadding = new Thickness(0) };
        if (slot.IsEmpty)
        {
            fd.Blocks.Add(new WpfDoc.Paragraph());
        }
        else
        {
            foreach (var para in slot.Paragraphs)
            {
                var wpfPara = new WpfDoc.Paragraph
                {
                    Margin        = new Thickness(0),
                    TextAlignment = para.Style.Alignment switch
                    {
                        PolyDonky.Core.Alignment.Center      => TextAlignment.Center,
                        PolyDonky.Core.Alignment.Right       => TextAlignment.Right,
                        PolyDonky.Core.Alignment.Justify     => TextAlignment.Justify,
                        PolyDonky.Core.Alignment.Distributed => TextAlignment.Justify,
                        _                                    => TextAlignment.Left,
                    },
                };
                foreach (var run in para.Runs)
                {
                    var wpfRun = new WpfDoc.Run(run.Text);
                    ApplyRunStyleToRun(wpfRun, run.Style);
                    wpfPara.Inlines.Add(wpfRun);
                }
                if (!wpfPara.Inlines.Any())
                    wpfPara.Inlines.Add(new WpfDoc.Run(string.Empty));
                fd.Blocks.Add(wpfPara);
            }
        }
        rtb.Document = fd;
    }

    private static HeaderFooterSlot SaveRtbToSlot(RichTextBox rtb)
    {
        var slot = new HeaderFooterSlot();
        foreach (var block in rtb.Document.Blocks)
        {
            if (block is not WpfDoc.Paragraph para) continue;
            var corePara = new PolyDonky.Core.Paragraph();
            corePara.Style.Alignment = para.TextAlignment switch
            {
                TextAlignment.Center  => PolyDonky.Core.Alignment.Center,
                TextAlignment.Right   => PolyDonky.Core.Alignment.Right,
                TextAlignment.Justify => PolyDonky.Core.Alignment.Justify,
                _                    => PolyDonky.Core.Alignment.Left,
            };
            ExtractInlines(para.Inlines, corePara);
            if (corePara.Runs.Any(r => !string.IsNullOrEmpty(r.Text)))
                slot.Paragraphs.Add(corePara);
        }
        return slot;
    }

    private static void ExtractInlines(WpfDoc.InlineCollection inlines, PolyDonky.Core.Paragraph para)
    {
        foreach (var inline in inlines)
        {
            switch (inline)
            {
                case WpfDoc.Run run when !string.IsNullOrEmpty(run.Text):
                    para.Runs.Add(new PolyDonky.Core.Run
                    {
                        Text  = run.Text,
                        Style = ExtractRunStyle(run),
                    });
                    break;
                case WpfDoc.Span span:
                    ExtractInlines(span.Inlines, para);
                    break;
            }
        }
    }

    private static PolyDonky.Core.RunStyle ExtractRunStyle(WpfDoc.TextElement elem)
    {
        double dip   = (double)elem.GetValue(WpfDoc.TextElement.FontSizeProperty);
        double pt    = double.IsNaN(dip) ? 11.0 : dip * 72.0 / 96.0;
        var fw       = (FontWeight)elem.GetValue(WpfDoc.TextElement.FontWeightProperty);
        var fst      = (FontStyle) elem.GetValue(WpfDoc.TextElement.FontStyleProperty);
        var ff       = elem.GetValue(WpfDoc.TextElement.FontFamilyProperty) as FontFamily;
        var fg       = elem.GetValue(WpfDoc.TextElement.ForegroundProperty) as SolidColorBrush;
        var td       = elem.GetValue(WpfDoc.Inline.TextDecorationsProperty) as TextDecorationCollection;
        return new PolyDonky.Core.RunStyle
        {
            FontSizePt    = pt,
            FontFamily    = ff?.Source,
            Bold          = fw == FontWeights.Bold,
            Italic        = fst == FontStyles.Italic,
            Underline     = td?.Any(d => d.Location == TextDecorationLocation.Underline)     == true,
            Strikethrough = td?.Any(d => d.Location == TextDecorationLocation.Strikethrough) == true,
            Foreground    = fg is not null
                ? new PolyDonky.Core.Color(fg.Color.R, fg.Color.G, fg.Color.B, fg.Color.A)
                : null,
        };
    }

    private static void ApplyRunStyleToRun(WpfDoc.Run wpfRun, PolyDonky.Core.RunStyle s)
    {
        if (s.FontSizePt > 0) wpfRun.FontSize = s.FontSizePt * 96.0 / 72.0;
        wpfRun.FontWeight = s.Bold   ? FontWeights.Bold  : FontWeights.Normal;
        wpfRun.FontStyle  = s.Italic ? FontStyles.Italic : FontStyles.Normal;
        if (!string.IsNullOrEmpty(s.FontFamily))
            wpfRun.FontFamily = new FontFamily(s.FontFamily);
        if (s.Foreground is { } fg)
            wpfRun.Foreground = new SolidColorBrush(
                WpfMedia.Color.FromArgb(fg.A, fg.R, fg.G, fg.B));
        var decs = new TextDecorationCollection();
        if (s.Underline)     decs.Add(TextDecorations.Underline[0]);
        if (s.Strikethrough) decs.Add(TextDecorations.Strikethrough[0]);
        if (decs.Count > 0)  wpfRun.TextDecorations = decs;
    }

    // ── 미리보기 ─────────────────────────────────────────────────

    private void UpdatePreview()
    {
        PreviewCanvas.Children.Clear();

        const double cw = 148;  // canvas 너비
        const double ch = 200;  // canvas 높이

        double pw = _settings.EffectiveWidthMm;
        double ph = _settings.EffectiveHeightMm;
        if (pw <= 0 || ph <= 0) return;

        double scale = Math.Min(cw / pw, ch / ph) * 0.92;
        double rw    = pw * scale;
        double rh    = ph * scale;
        double ox    = (cw - rw) / 2;
        double oy    = (ch - rh) / 2;

        // 용지 외곽 (그림자)
        var shadow = new Rectangle
        {
            Width  = rw, Height = rh,
            Fill   = new WpfMedia.SolidColorBrush(WpfMedia.Color.FromArgb(40, 0, 0, 0)),
        };
        Canvas.SetLeft(shadow, ox + 3);
        Canvas.SetTop(shadow, oy + 3);
        PreviewCanvas.Children.Add(shadow);

        // 용지 본체
        var paper = new Rectangle
        {
            Width       = rw, Height = rh,
            Fill        = GetPaperBrush(),
            Stroke      = new WpfMedia.SolidColorBrush(WpfMedia.Color.FromRgb(0xAA, 0xAA, 0xAA)),
            StrokeThickness = 0.5,
        };
        Canvas.SetLeft(paper, ox);
        Canvas.SetTop(paper, oy);
        PreviewCanvas.Children.Add(paper);

        // 여백 경계선 (파란색 점선)
        double ml = _settings.MarginLeftMm   * scale;
        double mr = _settings.MarginRightMm  * scale;
        double mt = _settings.MarginTopMm    * scale;
        double mb = _settings.MarginBottomMm * scale;
        double mh = _settings.MarginHeaderMm * scale;
        double mf = _settings.MarginFooterMm * scale;

        double cx = ox + ml;
        double cy = oy + mt;
        double crw = rw - ml - mr;
        double crh = rh - mt - mb;

        if (crw > 0 && crh > 0)
        {
            var margin = new Rectangle
            {
                Width  = crw, Height = crh,
                Fill   = WpfMedia.Brushes.Transparent,
                Stroke = new WpfMedia.SolidColorBrush(WpfMedia.Color.FromArgb(180, 0x44, 0x88, 0xCC)),
                StrokeThickness = 0.7,
                StrokeDashArray = new WpfMedia.DoubleCollection { 3, 2 },
            };
            Canvas.SetLeft(margin, cx);
            Canvas.SetTop(margin, cy);
            PreviewCanvas.Children.Add(margin);

            // 머리말 선
            if (mh > 0 && mh < mt)
            {
                double hy = oy + mh;
                var hLine = MakeHLine(ox, hy, rw, WpfMedia.Color.FromArgb(120, 0x44, 0x88, 0xCC));
                PreviewCanvas.Children.Add(hLine);
            }
            // 꼬리말 선
            if (mf > 0 && mf < mb)
            {
                double fy = oy + rh - mf;
                var fLine = MakeHLine(ox, fy, rw, WpfMedia.Color.FromArgb(120, 0x44, 0x88, 0xCC));
                PreviewCanvas.Children.Add(fLine);
            }

            // 단 구분선
            if (_settings.ColumnCount > 1 && crw > 0)
            {
                double gapPx = _settings.ColumnGapMm * scale;
                double colW  = (crw - gapPx * (_settings.ColumnCount - 1)) / _settings.ColumnCount;
                for (int i = 1; i < _settings.ColumnCount; i++)
                {
                    double lx = cx + i * (colW + gapPx) - gapPx / 2;
                    var vLine = new Line
                    {
                        X1 = lx, Y1 = cy, X2 = lx, Y2 = cy + crh,
                        Stroke          = new WpfMedia.SolidColorBrush(WpfMedia.Color.FromArgb(100, 0xFF, 0x80, 0x00)),
                        StrokeThickness = 0.6,
                        StrokeDashArray = new WpfMedia.DoubleCollection { 3, 2 },
                    };
                    PreviewCanvas.Children.Add(vLine);
                }
            }
        }

        // 치수 레이블
        var label = new System.Windows.Controls.TextBlock
        {
            Text     = $"{pw:0.#}×{ph:0.#} mm",
            FontSize = 9,
            Foreground = SystemColors.GrayTextBrush,
        };
        Canvas.SetLeft(label, ox);
        Canvas.SetTop(label, oy + rh + 2);
        PreviewCanvas.Children.Add(label);
    }

    private WpfMedia.Brush GetPaperBrush()
    {
        if (!string.IsNullOrEmpty(_settings.PaperColor) &&
            TryParseWpfColor(_settings.PaperColor, out var c))
            return new WpfMedia.SolidColorBrush(c);
        return WpfMedia.Brushes.White;
    }

    private static Line MakeHLine(double x, double y, double w, WpfMedia.Color color) => new()
    {
        X1 = x, Y1 = y, X2 = x + w, Y2 = y,
        Stroke          = new WpfMedia.SolidColorBrush(color),
        StrokeThickness = 0.6,
        StrokeDashArray = new WpfMedia.DoubleCollection { 4, 3 },
    };

    // ── 버튼 ─────────────────────────────────────────────────────

    private void OnOk(object sender, RoutedEventArgs e)
    {
        _settings.Header = new HeaderFooterContent
        {
            Left   = SaveRtbToSlot(RtbHeaderLeft),
            Center = SaveRtbToSlot(RtbHeaderCenter),
            Right  = SaveRtbToSlot(RtbHeaderRight),
        };
        _settings.Footer = new HeaderFooterContent
        {
            Left   = SaveRtbToSlot(RtbFooterLeft),
            Center = SaveRtbToSlot(RtbFooterCenter),
            Right  = SaveRtbToSlot(RtbFooterRight),
        };
        ResultSettings = Clone(_settings);
        DialogResult   = true;
    }

    private void OnCancel(object sender, RoutedEventArgs e) => DialogResult = false;

    // ── 유틸 ─────────────────────────────────────────────────────

    private static void TrySetMm(TextBox tb, Action<double> setter)
    {
        if (double.TryParse(tb.Text, out double v) && v >= 0)
            setter(v);
    }

    private static bool TryParseWpfColor(string? hex, out WpfMedia.Color color)
    {
        color = default;
        if (string.IsNullOrWhiteSpace(hex)) return false;
        var t = hex.Trim();
        if (!t.StartsWith('#')) t = '#' + t;
        try { color = (WpfMedia.Color)WpfMedia.ColorConverter.ConvertFromString(t)!; return true; }
        catch { return false; }
    }

    private static string NormalizeHex(string hex)
    {
        var t = hex.Trim();
        return t.StartsWith('#') ? t : '#' + t;
    }

    private static PageSettings Clone(PageSettings s) => new()
    {
        SizeKind        = s.SizeKind,
        WidthMm         = s.WidthMm,
        HeightMm        = s.HeightMm,
        Orientation     = s.Orientation,
        PaperColor      = s.PaperColor,
        MarginTopMm     = s.MarginTopMm,
        MarginBottomMm  = s.MarginBottomMm,
        MarginLeftMm    = s.MarginLeftMm,
        MarginRightMm   = s.MarginRightMm,
        MarginHeaderMm  = s.MarginHeaderMm,
        MarginFooterMm  = s.MarginFooterMm,
        ColumnCount     = s.ColumnCount,
        ColumnGapMm     = s.ColumnGapMm,
        ColumnDividerVisible      = s.ColumnDividerVisible,
        ColumnDividerColor        = s.ColumnDividerColor,
        ColumnDividerThicknessPt  = s.ColumnDividerThicknessPt,
        ColumnDividerStyle        = s.ColumnDividerStyle,
        PageNumberStart = s.PageNumberStart,
        Header          = s.Header.Clone(),
        Footer          = s.Footer.Clone(),
        DifferentFirstPage = s.DifferentFirstPage,
        DifferentOddEven   = s.DifferentOddEven,
        ShowMarginGuides   = s.ShowMarginGuides,
        TextOrientation    = s.TextOrientation,
        TextProgression    = s.TextProgression,
    };
}

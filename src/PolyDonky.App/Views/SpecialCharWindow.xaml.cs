using System;
using System.Collections.Generic;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;

namespace PolyDonky.App.Views;

public partial class SpecialCharWindow : Window
{
    private readonly RichTextBox _editor;
    private string? _selectedChar;

    public SpecialCharWindow(RichTextBox editor)
    {
        InitializeComponent();
        _editor = editor;
        Loaded += OnLoaded;
    }

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        foreach (var category in Categories)
            CategoryList.Items.Add(category.Name);
        CategoryList.SelectedIndex = 0;
        UnicodeInput.Focus();
    }

    private void OnCategoryChanged(object sender, SelectionChangedEventArgs e)
    {
        if (CategoryList.SelectedIndex < 0 || CategoryList.SelectedIndex >= Categories.Count)
            return;
        var chars = Categories[CategoryList.SelectedIndex].Characters;
        CharGrid.ItemsSource = chars;
    }

    private void OnCharButtonClick(object sender, RoutedEventArgs e)
    {
        if (sender is Button b && b.Content is string s)
            SelectChar(s);
    }

    private void OnCharButtonDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (sender is Button b && b.Content is string s)
        {
            SelectChar(s);
            InsertSelectedAndClose();
        }
    }

    private void SelectChar(string s)
    {
        _selectedChar = s;
        PreviewText.Text = s;
        var cp = char.ConvertToUtf32(s, 0);
        CodePointText.Text = $"U+{cp:X4}  ({s})";
        CharNameText.Text = LookupName(cp) ?? "(이름 정보 없음)";
    }

    private void OnUnicodeInputKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            OnFindUnicode(sender, e);
            e.Handled = true;
        }
    }

    private void OnFindUnicode(object sender, RoutedEventArgs e)
    {
        var text = (UnicodeInput.Text ?? string.Empty).Trim();
        if (text.StartsWith("U+", StringComparison.OrdinalIgnoreCase))
            text = text[2..];
        if (!int.TryParse(text, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var cp)
            || cp < 0 || cp > 0x10FFFF)
        {
            MessageBox.Show(this, "유효하지 않은 유니코드 코드포인트입니다.", "특수문자",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        try
        {
            SelectChar(char.ConvertFromUtf32(cp));
        }
        catch (ArgumentException)
        {
            MessageBox.Show(this, "표시할 수 없는 코드포인트입니다.", "특수문자",
                MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private void OnInsertClick(object sender, RoutedEventArgs e)
    {
        if (_selectedChar is null) return;
        InsertSelectedAndClose();
    }

    private void OnCloseClick(object sender, RoutedEventArgs e) => Close();

    private void InsertSelectedAndClose()
    {
        if (_selectedChar is null) return;
        InsertAtCaret(_editor, _selectedChar);
        DialogResult = true;
        Close();
    }

    /// <summary>현재 캐럿 위치에 일반 텍스트를 삽입한다. 선택 영역이 있으면 대체.</summary>
    public static void InsertAtCaret(RichTextBox editor, string text)
    {
        if (string.IsNullOrEmpty(text)) return;

        if (!editor.Selection.IsEmpty)
        {
            editor.Selection.Text = string.Empty;
        }

        var caret = editor.CaretPosition;
        // CaretPosition 이 텍스트 컨텍스트가 아닐 수 있으므로 가장 가까운 삽입 가능한 지점으로 보정.
        var insertPos = caret.GetInsertionPosition(LogicalDirection.Forward) ?? caret;
        insertPos.InsertTextInRun(text);
        editor.CaretPosition = insertPos.GetPositionAtOffset(text.Length) ?? insertPos;
        editor.Focus();
    }

    // ── 카테고리 데이터 ────────────────────────────────────────────────────
    private sealed record CharCategory(string Name, List<string> Characters);

    private static readonly List<CharCategory> Categories = new()
    {
        new("자주 쓰는", Expand("'“”‘’«»•·…–—™©®°§¶†‡№※¿¡")),
        new("라틴 보충", Expand("ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÑÒÓÔÕÖØÙÚÛÜÝßàáâãäåæçèéêëìíîïñòóôõöøùúûüýÿ")),
        new("그리스", Expand("ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩαβγδεζηθικλμνξοπρςστυφχψω")),
        new("키릴", Expand("АБВГДЕЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯабвгдежзийклмнопрстуфхцчшщъыьэюя")),
        new("화살표", Expand("←↑→↓↔↕↖↗↘↙⇐⇑⇒⇓⇔⇕↩↪↰↱↲↳↺↻↵↶↷⟵⟶⟷⟸⟹⟺")),
        new("수학기호", Expand("±×÷√∛∜∞≈≠≤≥∑∏∫∂∇∈∉⊂⊃⊄⊆⊇∪∩∀∃∄∅∧∨¬≡≢∥∦⊥⊤∠⦠′″‴")),
        new("통화", Expand("$¢£¤¥₩€₠₡₢₣₤₥₦₧₨₪₫₭₮₯₰₱₲₳₴₵₹₽₿")),
        new("숫자/위첨자", Expand("¹²³⁰⁴⁵⁶⁷⁸⁹₀₁₂₃₄₅₆₇₈₉¼½¾⅓⅔⅕⅖⅗⅘⅙⅚⅛⅜⅝⅞ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩⅪⅫ")),
        new("도형", Expand("■□▣▤▥▦▧▨▩▪▫▲△▴▵▶▷▼▽▾▿◀◁◆◇○●◐◑◘◙◦★☆♠♡♢♣♤♥♦♧♨♩♪♫")),
        new("표시", Expand("☀☁☂☃☄☎☑☒☓☔☕☖☗☞☠☯☺☻♀♂♈♉♊♋♌♍♎♏♐♑♒♓♔♕♖♗♘♙♚♛♜♝♞♟⚀⚁⚂⚃⚄⚅✓✔✗✘✰✱✲✳✴❁❍❖❤")),
        new("한글 자모", Expand("ㄱㄲㄳㄴㄵㄶㄷㄸㄹㄺㄻㄼㄽㄾㄿㅀㅁㅂㅃㅄㅅㅆㅇㅈㅉㅊㅋㅌㅍㅎㅏㅐㅑㅒㅓㅔㅕㅖㅗㅘㅙㅚㅛㅜㅝㅞㅟㅠㅡㅢㅣ")),
        new("괄호/구두점", Expand("〈〉⟨⟩⟪⟫〈〉《》「」『』【】〔〕〖〗〘〙〚〛、。・ー，．：；？！")),
        new("박스 그리기", Expand("─│┌┐└┘├┤┬┴┼═║╔╗╚╝╠╣╦╩╬▀▄█▌▐░▒▓")),
    };

    private static List<string> Expand(string s)
    {
        var list = new List<string>();
        var en = StringInfo.GetTextElementEnumerator(s);
        while (en.MoveNext())
            list.Add((string)en.Current!);
        return list;
    }

    /// <summary>일부 자주 쓰이는 코드포인트에 대해 사람이 읽을 수 있는 이름을 돌려준다.
    /// 전수 Unicode 이름 데이터베이스는 무거우므로 카테고리 단위 라벨링만 제공.</summary>
    private static string? LookupName(int cp)
    {
        return cp switch
        {
            >= 0x0020 and <= 0x007E   => "Basic Latin",
            >= 0x00A0 and <= 0x00FF   => "Latin-1 Supplement",
            >= 0x0370 and <= 0x03FF   => "Greek and Coptic",
            >= 0x0400 and <= 0x04FF   => "Cyrillic",
            >= 0x2000 and <= 0x206F   => "General Punctuation",
            >= 0x2070 and <= 0x209F   => "Superscripts and Subscripts",
            >= 0x20A0 and <= 0x20CF   => "Currency Symbols",
            >= 0x2100 and <= 0x214F   => "Letterlike Symbols",
            >= 0x2150 and <= 0x218F   => "Number Forms",
            >= 0x2190 and <= 0x21FF   => "Arrows",
            >= 0x2200 and <= 0x22FF   => "Mathematical Operators",
            >= 0x2300 and <= 0x23FF   => "Miscellaneous Technical",
            >= 0x2500 and <= 0x257F   => "Box Drawing",
            >= 0x2580 and <= 0x259F   => "Block Elements",
            >= 0x25A0 and <= 0x25FF   => "Geometric Shapes",
            >= 0x2600 and <= 0x26FF   => "Miscellaneous Symbols",
            >= 0x2700 and <= 0x27BF   => "Dingbats",
            >= 0x27C0 and <= 0x27EF   => "Miscellaneous Mathematical Symbols-A",
            >= 0x27F0 and <= 0x27FF   => "Supplemental Arrows-A",
            >= 0x3000 and <= 0x303F   => "CJK Symbols and Punctuation",
            >= 0x3040 and <= 0x309F   => "Hiragana",
            >= 0x30A0 and <= 0x30FF   => "Katakana",
            >= 0x3130 and <= 0x318F   => "Hangul Compatibility Jamo",
            >= 0xAC00 and <= 0xD7AF   => "Hangul Syllables",
            >= 0x4E00 and <= 0x9FFF   => "CJK Unified Ideographs",
            >= 0xFF00 and <= 0xFFEF   => "Halfwidth and Fullwidth Forms",
            _                          => null,
        };
    }
}

using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;

namespace PolyDoc.App.Views;

public partial class EquationWindow : Window
{
    private readonly RichTextBox _editor;

    public EquationWindow(RichTextBox editor)
    {
        InitializeComponent();
        _editor = editor;
        BuildQuickPalette();
        SourceText.Text = @"E = mc^2";
        UpdatePreview();
        Loaded += (_, _) => SourceText.Focus();
    }

    private void BuildQuickPalette()
    {
        // 자주 쓰이는 LaTeX 토큰. 클릭 시 커서 위치에 삽입.
        var entries = new (string label, string snippet, string tip)[]
        {
            ("α",  @"\alpha",   "alpha"),
            ("β",  @"\beta",    "beta"),
            ("γ",  @"\gamma",   "gamma"),
            ("π",  @"\pi",      "pi"),
            ("Σ",  @"\sum",     "sum"),
            ("∏",  @"\prod",    "product"),
            ("∫",  @"\int",     "integral"),
            ("√",  @"\sqrt{}",  "sqrt"),
            ("a/b",@"\frac{}{}","fraction"),
            ("xⁿ", "^{}",        "superscript"),
            ("xₙ", "_{}",        "subscript"),
            ("≤",  @"\le",      "less-or-equal"),
            ("≥",  @"\ge",      "greater-or-equal"),
            ("≠",  @"\ne",      "not equal"),
            ("∞",  @"\infty",   "infinity"),
            ("→",  @"\to",      "to / arrow"),
        };

        foreach (var (label, snippet, tip) in entries)
        {
            var btn = new Button
            {
                Content = label,
                Width = 36,
                Height = 30,
                Margin = new Thickness(2, 0, 2, 0),
                ToolTip = $"{snippet} — {tip}",
                Tag = snippet,
            };
            btn.Click += OnQuickInsert;
            QuickPanel.Children.Add(btn);
        }
    }

    private void OnQuickInsert(object sender, RoutedEventArgs e)
    {
        if (sender is not Button b || b.Tag is not string snippet) return;
        var caret = SourceText.CaretIndex;
        SourceText.Text = SourceText.Text.Insert(caret, snippet);
        // {} 가 포함된 경우 첫 번째 빈 그룹 안에 캐럿 위치.
        var braceIdx = snippet.IndexOf("{}");
        SourceText.CaretIndex = braceIdx >= 0 ? caret + braceIdx + 1 : caret + snippet.Length;
        SourceText.Focus();
    }

    private void OnSourceChanged(object sender, TextChangedEventArgs e) => UpdatePreview();

    private void UpdatePreview()
    {
        var src = SourceText.Text ?? string.Empty;
        var (open, close) = RbDisplay.IsChecked == true ? (@"\[", @"\]") : (@"\(", @"\)");
        PreviewText.Text = $"{open} {src} {close}";
    }

    private void OnInsertClick(object sender, RoutedEventArgs e)
    {
        var src = (SourceText.Text ?? string.Empty).Trim();
        if (src.Length == 0)
        {
            MessageBox.Show(this, "수식을 입력하세요.", "수식",
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var isDisplay = RbDisplay.IsChecked == true;
        var (open, close) = isDisplay ? (@"\[", @"\]") : (@"\(", @"\)");
        var fullSource = $"{open}{src}{close}";

        // 현 단계: 본문 캐럿 위치에 monospace 폰트로 텍스트 삽입.
        // 추후 사이클에서 EquationRun (모델) + 실제 렌더러로 승격 예정.
        InsertEquationText(_editor, fullSource, isDisplay);

        DialogResult = true;
        Close();
    }

    private void OnCloseClick(object sender, RoutedEventArgs e) => Close();

    /// <summary>편집기 캐럿 위치에 수식 텍스트를 삽입한 뒤 monospace 스타일을 적용한다.</summary>
    private static void InsertEquationText(RichTextBox editor, string text, bool isDisplay)
    {
        if (string.IsNullOrEmpty(text)) return;

        if (!editor.Selection.IsEmpty)
            editor.Selection.Text = string.Empty;

        var caret = editor.CaretPosition;
        var insertPos = caret.GetInsertionPosition(LogicalDirection.Forward) ?? caret;
        insertPos.InsertTextInRun(text);
        var endPos = insertPos.GetPositionAtOffset(text.Length) ?? insertPos;

        // 삽입한 범위에 수식 스타일을 적용.
        var range = new TextRange(insertPos, endPos);
        range.ApplyPropertyValue(TextElement.FontFamilyProperty,
            new FontFamily("Cambria Math, Cambria, Consolas, monospace"));
        range.ApplyPropertyValue(TextElement.FontSizeProperty,
            isDisplay ? 16.0 : 14.0);

        editor.CaretPosition = endPos;
        editor.Focus();
    }
}

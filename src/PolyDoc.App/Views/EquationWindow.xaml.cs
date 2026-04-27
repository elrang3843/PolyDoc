using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using PolyDoc.App.Services;
using WpfMath.Controls;
using CoreRun = PolyDoc.Core.Run;

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
            ("→",  @"\to",      "arrow"),
        };

        foreach (var (label, snippet, tip) in entries)
        {
            var btn = new Button
            {
                Content = label,
                Width   = 36,
                Height  = 30,
                Margin  = new Thickness(2, 0, 2, 0),
                ToolTip = $"{snippet} — {tip}",
                Tag     = snippet,
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
        var braceIdx = snippet.IndexOf("{}");
        SourceText.CaretIndex = braceIdx >= 0 ? caret + braceIdx + 1 : caret + snippet.Length;
        SourceText.Focus();
    }

    private void OnSourceChanged(object sender, TextChangedEventArgs e) => UpdatePreview();
    private void OnModeChanged(object sender, RoutedEventArgs e) => UpdatePreview();

    private void UpdatePreview()
    {
        var src = SourceText?.Text ?? string.Empty;
        if (FormulaPreview is null) return;

        try
        {
            FormulaPreview.Formula = src;
            FormulaPreview.Visibility = Visibility.Visible;
            PreviewErrorText.Visibility = Visibility.Collapsed;
        }
        catch (Exception ex)
        {
            FormulaPreview.Visibility = Visibility.Collapsed;
            PreviewErrorText.Text = $"수식 파싱 오류: {ex.Message}";
            PreviewErrorText.Visibility = Visibility.Visible;
        }
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

        InsertEquationInline(_editor, src, isDisplay: RbDisplay.IsChecked == true);
        DialogResult = true;
        Close();
    }

    private void OnCloseClick(object sender, RoutedEventArgs e) => Close();

    /// <summary>RichTextBox 캐럿 위치에 수식 이미지를 InlineUIContainer 로 삽입.
    /// FormulaControl 은 XamlWriter.Save() 와 충돌하므로 오프스크린 렌더링 후 Image 로 변환한다.
    /// Tag 에 Core Run(LatexSource 포함)을 달아 FlowDocumentParser 가 라운드트립 가능하게 한다.</summary>
    private static void InsertEquationInline(RichTextBox editor, string latexSource, bool isDisplay)
    {
        if (string.IsNullOrEmpty(latexSource)) return;

        if (!editor.Selection.IsEmpty)
            editor.Selection.Text = string.Empty;

        var caret     = editor.CaretPosition;
        var insertPos = caret.GetInsertionPosition(LogicalDirection.Forward) ?? caret;

        var (open, close) = isDisplay ? (@"\[", @"\]") : (@"\(", @"\)");
        var modelRun = new CoreRun
        {
            Text              = $"{open}{latexSource}{close}",
            LatexSource       = latexSource,
            IsDisplayEquation = isDisplay,
        };

        var img = FlowDocumentBuilder.RenderEquationToImage(latexSource, isDisplay ? 18.0 : 14.0);
        if (img is null)
        {
            // 파싱 실패: plain-text 폴백
            SpecialCharWindow.InsertAtCaret(editor, modelRun.Text);
            return;
        }

        // InlineUIContainer(UIElement, TextPointer) 생성자가 지정 위치에 즉시 삽입한다.
        var iuc = new InlineUIContainer(img, insertPos)
        {
            Tag               = modelRun,
            BaselineAlignment = BaselineAlignment.Center,
        };

        editor.CaretPosition = iuc.ElementEnd;
        editor.Focus();
    }
}

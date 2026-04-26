using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using PolyDoc.App.Services;

namespace PolyDoc.App.Views;

public partial class FindReplaceWindow : Window
{
    private readonly RichTextBox _editor;
    private TextPointer? _lastMatchEnd;

    public FindReplaceWindow(RichTextBox editor)
    {
        InitializeComponent();
        _editor = editor;
    }

    private void OnKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Return) FindNext();
    }

    private void OnFindNext(object sender, RoutedEventArgs e) => FindNext();

    private void OnReplace(object sender, RoutedEventArgs e)
    {
        var query = FindBox.Text;
        if (string.IsNullOrEmpty(query)) { SetStatus("찾을 내용을 입력하세요."); return; }

        var caseSensitive = CaseSensitiveBox.IsChecked == true;
        var comparison = caseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

        // 현재 선택이 검색어와 일치하면 교체
        var sel = _editor.Selection;
        if (!sel.IsEmpty && string.Equals(sel.Text, query, comparison))
        {
            sel.Text = ReplaceBox.Text;
            _lastMatchEnd = _editor.Selection.End;
        }

        // 다음 찾기
        FindNext();
    }

    private void OnReplaceAll(object sender, RoutedEventArgs e)
    {
        var query = FindBox.Text;
        if (string.IsNullOrEmpty(query)) { SetStatus("찾을 내용을 입력하세요."); return; }

        var count = FlowDocumentSearch.ReplaceAll(
            _editor.Document, query, ReplaceBox.Text, CaseSensitiveBox.IsChecked == true);

        _lastMatchEnd = null;
        SetStatus(count == 0 ? "찾을 수 없습니다." : $"{count}개 바꿨습니다.");
    }

    private void OnClose(object sender, RoutedEventArgs e) => Close();

    private void FindNext()
    {
        var query = FindBox.Text;
        if (string.IsNullOrEmpty(query)) { SetStatus("찾을 내용을 입력하세요."); return; }

        var doc = _editor.Document;
        var start = _lastMatchEnd ?? doc.ContentStart;
        var found = FlowDocumentSearch.FindNext(doc, query, start, CaseSensitiveBox.IsChecked == true);

        if (found is null)
        {
            found = FlowDocumentSearch.FindNext(doc, query, doc.ContentStart, CaseSensitiveBox.IsChecked == true);
            if (found is null) { _lastMatchEnd = null; SetStatus("찾을 수 없습니다."); return; }
            SetStatus("문서 처음부터 다시 검색했습니다.");
        }
        else
        {
            SetStatus(string.Empty);
        }

        _editor.Selection.Select(found.Start, found.End);
        _editor.Focus();
        _lastMatchEnd = found.End;
    }

    private void SetStatus(string msg) => StatusText.Text = msg;
}

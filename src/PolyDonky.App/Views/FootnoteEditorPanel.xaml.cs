using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using PolyDonky.Core;

namespace PolyDonky.App.Views;

/// <summary>
/// 본문 하단에 도킹되는 각주/미주 편집 패널.
/// <list type="bullet">
/// <item>탭으로 각주/미주 모드 전환.</item>
/// <item>각 항목: 번호 + 내용 TextBox + 본문이동 + 삭제 버튼.</item>
/// <item>TextBox 편집 시 즉시 <see cref="FootnoteEntry"/> 의 Blocks 갱신
///       (단일 Paragraph 안에 줄바꿈 보존된 Run 들로 변환).</item>
/// </list>
/// 본문 측 참조 런 갱신·삭제는 <see cref="EntryDeleted"/>·<see cref="EntryFocusRequested"/>
/// 이벤트로 외부(MainWindow) 에 위임한다.
/// </summary>
public partial class FootnoteEditorPanel : UserControl
{
    /// <summary>현재 표시 중인 모드.</summary>
    public enum NoteKind { Footnote, Endnote }

    private PolyDonkyument? _document;
    private NoteKind _currentKind = NoteKind.Footnote;
    private readonly ObservableCollection<EntryViewModel> _items = new();

    /// <summary>사용자가 항목 삭제를 요청. 인자: (Kind, EntryId).</summary>
    public event EventHandler<(NoteKind Kind, string Id)>? EntryDeleted;

    /// <summary>사용자가 본문 참조 위치로 이동을 요청. 인자: (Kind, EntryId).</summary>
    public event EventHandler<(NoteKind Kind, string Id)>? EntryFocusRequested;

    /// <summary>패널 닫기 버튼이 클릭됨.</summary>
    public event EventHandler? CloseRequested;

    public FootnoteEditorPanel()
    {
        InitializeComponent();
        EntriesList.ItemsSource = _items;
    }

    /// <summary>문서를 패널에 바인딩. 이후 <see cref="Refresh"/> 로 표시 갱신.</summary>
    public void BindToDocument(PolyDonkyument document)
    {
        _document = document;
        Refresh();
    }

    /// <summary>현재 모드(각주/미주) 의 항목 목록을 문서에서 다시 읽어 표시.</summary>
    public void Refresh()
    {
        _items.Clear();

        var source = GetSourceList();
        if (source is null)
        {
            UpdateEmptyHint();
            return;
        }

        for (int i = 0; i < source.Count; i++)
        {
            _items.Add(new EntryViewModel(source[i], i + 1));
        }
        UpdateEmptyHint();
    }

    /// <summary>특정 항목 ID 의 TextBox 에 포커스. 없으면 무시.</summary>
    public void FocusEntry(NoteKind kind, string id)
    {
        // 모드 전환이 필요하면 먼저 탭 변경
        if (kind != _currentKind)
        {
            if (kind == NoteKind.Footnote) TabFn.IsChecked = true;
            else                            TabEn.IsChecked = true;
        }

        var idx = _items.ToList().FindIndex(e => e.Id == id);
        if (idx < 0) return;

        // TextBox 에 포커스 — Dispatcher 로 Layout 후 실행
        Dispatcher.BeginInvoke(new Action(() =>
        {
            if (EntriesList.ItemContainerGenerator.ContainerFromIndex(idx)
                is ContentPresenter cp)
            {
                cp.ApplyTemplate();
                var tb = FindDescendant<TextBox>(cp);
                tb?.Focus();
                tb?.SelectAll();
            }
        }), System.Windows.Threading.DispatcherPriority.Loaded);
    }

    // ── 이벤트 핸들러 ──────────────────────────────────────────────────────

    private void OnTabChanged(object sender, RoutedEventArgs e)
    {
        if (!IsLoaded) return;
        _currentKind = TabFn.IsChecked == true ? NoteKind.Footnote : NoteKind.Endnote;
        Refresh();
    }

    private void OnCloseClick(object sender, RoutedEventArgs e)
        => CloseRequested?.Invoke(this, EventArgs.Empty);

    private void OnDeleteClick(object sender, RoutedEventArgs e)
    {
        if (sender is FrameworkElement fe && fe.Tag is string id)
            EntryDeleted?.Invoke(this, (_currentKind, id));
    }

    private void OnGotoRefClick(object sender, RoutedEventArgs e)
    {
        if (sender is FrameworkElement fe && fe.Tag is string id)
            EntryFocusRequested?.Invoke(this, (_currentKind, id));
    }

    // ── 내부 헬퍼 ───────────────────────────────────────────────────────────

    private IList<FootnoteEntry>? GetSourceList()
    {
        if (_document is null) return null;
        return _currentKind == NoteKind.Footnote ? _document.Footnotes : _document.Endnotes;
    }

    private void UpdateEmptyHint()
        => EmptyHint.Visibility = _items.Count == 0 ? Visibility.Visible : Visibility.Collapsed;

    private static T? FindDescendant<T>(DependencyObject? root) where T : DependencyObject
    {
        if (root is null) return null;
        int count = System.Windows.Media.VisualTreeHelper.GetChildrenCount(root);
        for (int i = 0; i < count; i++)
        {
            var child = System.Windows.Media.VisualTreeHelper.GetChild(root, i);
            if (child is T match) return match;
            if (FindDescendant<T>(child) is { } nested) return nested;
        }
        return null;
    }

    // ── ViewModel: 단일 FootnoteEntry 의 표시·편집 상태 ───────────────────

    /// <summary>FootnoteEntry 의 wrapper — 단순 텍스트 편집을 위해
    /// Blocks ↔ 평문 텍스트 변환을 양방향으로 처리한다.</summary>
    private sealed class EntryViewModel : INotifyPropertyChanged
    {
        private readonly FootnoteEntry _entry;
        private string _content;

        public EntryViewModel(FootnoteEntry entry, int displayNumber)
        {
            _entry        = entry;
            _content      = BlocksToText(entry.Blocks);
            NumberLabel   = displayNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".";
        }

        public string Id => _entry.Id;
        public string NumberLabel { get; }

        public string Content
        {
            get => _content;
            set
            {
                if (_content == value) return;
                _content = value;
                _entry.Blocks = TextToBlocks(value);
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string? n = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(n));

        // ── Blocks ↔ Text 변환 ─────────────────────────────────────────
        // 단순 평문 모델: 단락 = 줄. 각 단락의 모든 Run.Text 를 합쳐 한 줄로,
        // 단락 사이는 '\n' 으로 결합. 입력 시 역변환.

        private static string BlocksToText(IList<Block> blocks)
        {
            var sb = new StringBuilder();
            bool first = true;
            foreach (var b in blocks)
            {
                if (b is Paragraph para)
                {
                    if (!first) sb.Append('\n');
                    sb.Append(para.GetPlainText());
                    first = false;
                }
            }
            return sb.ToString();
        }

        private static IList<Block> TextToBlocks(string text)
        {
            var list = new List<Block>();
            // 빈 입력도 빈 단락 1 개로 보존 — 라운드트립 안정성.
            var lines = (text ?? string.Empty).Split('\n');
            foreach (var line in lines)
                list.Add(Paragraph.Of(line));
            return list;
        }
    }
}

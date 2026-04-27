using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media.Imaging;
using PolyDonky.App.Services;
using CoreRun = PolyDonky.Core.Run;

namespace PolyDonky.App.Views;

/// <summary>
/// 이모지 선택 다이얼로그. 8개 섹션 × 10개 = 80개의 PNG 이모지를 카테고리별 그리드로 표시.
/// 선택 시 활성 RichTextBox 의 캐럿 위치에 <see cref="InlineUIContainer"/> 로 삽입.
/// 라운드트립용으로 IUC.Tag 에 <see cref="CoreRun"/> (EmojiKey 포함) 을 심어둔다.
/// </summary>
public partial class EmojiWindow : Window
{
    private readonly RichTextBox _editor;
    private EmojiItem? _selected;

    public EmojiWindow(RichTextBox editor)
    {
        InitializeComponent();
        _editor = editor;
        Loaded += OnLoaded;
    }

    internal static readonly (double Pt, string Label)[] SizeOptions =
    {
        (12, "12pt — 소형"),
        (16, "16pt — 중형"),
        (20, "20pt — 보통"),
        (24, "24pt — 대형"),
        (32, "32pt — 특대"),
        (48, "48pt — 초대"),
    };

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        foreach (var c in Categories)
            CategoryList.Items.Add(c.DisplayName);
        CategoryList.SelectedIndex = 0;

        foreach (var (_, label) in SizeOptions)
            SizeCombo.Items.Add(label);
        SizeCombo.SelectedIndex = 1; // 16pt 기본

        SearchInput.Focus();
    }

    private double SelectedSizePt =>
        SizeCombo.SelectedIndex >= 0 && SizeCombo.SelectedIndex < SizeOptions.Length
            ? SizeOptions[SizeCombo.SelectedIndex].Pt
            : 16.0;

    private void OnCategoryChanged(object sender, SelectionChangedEventArgs e)
    {
        if (CategoryList.SelectedIndex < 0) return;
        // 카테고리 변경 시 검색어 초기화 — 카테고리 의도가 검색을 덮는다.
        if (!string.IsNullOrEmpty(SearchInput.Text))
            SearchInput.Text = string.Empty;
        EmojiGrid.ItemsSource = Categories[CategoryList.SelectedIndex].Items;
    }

    private void OnSearchChanged(object sender, TextChangedEventArgs e)
    {
        var q = (SearchInput.Text ?? string.Empty).Trim();
        if (q.Length == 0)
        {
            // 빈 검색어 — 현재 카테고리만 표시
            if (CategoryList.SelectedIndex >= 0)
                EmojiGrid.ItemsSource = Categories[CategoryList.SelectedIndex].Items;
            return;
        }

        // 전체 카테고리에서 키 또는 라벨에 q 가 포함된 항목 (대소문자 무시)
        var hits = Categories
            .SelectMany(c => c.Items)
            .Where(it => it.Name.Contains(q, StringComparison.OrdinalIgnoreCase)
                      || it.Section.Contains(q, StringComparison.OrdinalIgnoreCase)
                      || it.Label.Contains(q, StringComparison.OrdinalIgnoreCase))
            .ToList();
        EmojiGrid.ItemsSource = hits;
    }

    private void OnEmojiButtonClick(object sender, RoutedEventArgs e)
    {
        if (sender is Button b && b.Tag is string key)
            SelectByKey(key);
    }

    private void OnEmojiButtonDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (sender is Button b && b.Tag is string key)
        {
            SelectByKey(key);
            InsertSelectedAndClose();
        }
    }

    private void SelectByKey(string key)
    {
        var item = Categories.SelectMany(c => c.Items).FirstOrDefault(i => i.Key == key);
        if (item is null) return;
        _selected = item;
        PreviewImage.Source = item.ImageSource;
        NameText.Text       = item.Label;
        KeyText.Text        = $"{item.Section} / {item.Name}  ·  키: {item.Key}";
    }

    private void OnInsertClick(object sender, RoutedEventArgs e)
    {
        if (_selected is null)
        {
            MessageBox.Show(this, "이모지를 먼저 선택하세요.", "이모지",
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }
        InsertSelectedAndClose();
    }

    private void OnCloseClick(object sender, RoutedEventArgs e) => Close();

    private void InsertSelectedAndClose()
    {
        if (_selected is null) return;
        InsertEmojiInline(_editor, _selected.Key, SelectedSizePt);
        DialogResult = true;
        Close();
    }

    /// <summary>
    /// 활성 편집기의 캐럿 위치에 이모지를 InlineUIContainer 로 삽입한다.
    /// Tag 에 EmojiKey 와 FontSizePt 가 들어간 PolyDonky.Core.Run 을 심어 라운드트립 보장.
    /// </summary>
    /// <param name="sizePt">이모지 표시 크기(포인트). 0 이하면 16pt 로 대체.</param>
    public static void InsertEmojiInline(RichTextBox editor, string emojiKey, double sizePt = 16.0)
    {
        if (string.IsNullOrEmpty(emojiKey)) return;
        if (sizePt <= 0) sizePt = 16.0;

        if (!editor.Selection.IsEmpty)
            editor.Selection.Text = string.Empty;

        var caret     = editor.CaretPosition;
        var insertPos = caret.GetInsertionPosition(LogicalDirection.Forward) ?? caret;

        var modelRun = new CoreRun
        {
            Text     = $"[{emojiKey}]",
            EmojiKey = emojiKey,
            Style    = new PolyDonky.Core.RunStyle { FontSizePt = sizePt },
        };

        double sizeDip = FlowDocumentBuilder.PtToDip(sizePt);
        var img = FlowDocumentBuilder.LoadEmojiImage(emojiKey, sizeDip);
        if (img is null)
        {
            SpecialCharWindow.InsertAtCaret(editor, modelRun.Text);
            return;
        }

        // img.Tag 에 iuc 를 저장하지 말 것 — iuc.Child = img 와 함께 순환 참조가 되어
        // XamlWriter.Save() 가 StackOverflow 로 폭주한다 (글자 속성 적용 시 undo 스냅샷에서 발생).
        var iuc = new InlineUIContainer(img, insertPos)
        {
            Tag               = modelRun,
            BaselineAlignment = BaselineAlignment.Center,
        };
        editor.CaretPosition = iuc.ElementEnd;
        editor.Focus();
    }

    // ── 카탈로그 ────────────────────────────────────────────────────────
    //
    // emoji-manifest.json 과 일치 — 8개 섹션 × 10개 = 80개. 라벨은 README 의 한글 의미.
    // 새 이모지가 추가되면 manifest 와 함께 이 표도 갱신해야 한다.

    private sealed record EmojiItem(string Section, string Name, string Label)
    {
        public string Key     => $"{Section}_{Name}";
        public string Tooltip => $"{Label}  ({Key})";
        public BitmapImage? ImageSource => LoadBitmap(Section, Name);

        private static BitmapImage? LoadBitmap(string section, string name)
        {
            try
            {
                var bmp = new BitmapImage();
                bmp.BeginInit();
                bmp.UriSource   = new Uri($"pack://application:,,,/Resources/Emojis/{section}/{name}.png", UriKind.Absolute);
                bmp.CacheOption = BitmapCacheOption.OnLoad;
                bmp.DecodePixelWidth = 48;
                bmp.EndInit();
                bmp.Freeze();
                return bmp;
            }
            catch { return null; }
        }
    }

    private sealed record EmojiCategory(string DisplayName, IReadOnlyList<EmojiItem> Items);

    private static readonly IReadOnlyList<EmojiCategory> Categories = new EmojiCategory[]
    {
        new("상태 (Status)", new EmojiItem[]
        {
            new("Status", "todo",        "할 일"),
            new("Status", "in_progress", "진행 중"),
            new("Status", "done",        "완료"),
            new("Status", "blocked",     "막힘"),
            new("Status", "review",      "검토"),
            new("Status", "draft",       "초안"),
            new("Status", "on_hold",     "보류"),
            new("Status", "waiting",     "대기"),
            new("Status", "archive",     "보관"),
            new("Status", "new",         "신규"),
        }),
        new("반응 (Reactions)", new EmojiItem[]
        {
            new("Reactions", "thumbs_up",   "추천"),
            new("Reactions", "thumbs_down", "비추천"),
            new("Reactions", "heart",       "하트"),
            new("Reactions", "laugh",       "웃음"),
            new("Reactions", "mind_blown",  "놀람"),
            new("Reactions", "thinking",    "생각 중"),
            new("Reactions", "celebrate",   "축하"),
            new("Reactions", "eyes",        "주목"),
            new("Reactions", "raised_hand", "손들기"),
            new("Reactions", "comment",     "댓글"),
        }),
        new("동작 (Actions)", new EmojiItem[]
        {
            new("Actions", "edit",     "수정"),
            new("Actions", "save",     "저장"),
            new("Actions", "share",    "공유"),
            new("Actions", "copy",     "복사"),
            new("Actions", "delete",   "삭제"),
            new("Actions", "download", "다운로드"),
            new("Actions", "upload",   "업로드"),
            new("Actions", "search",   "검색"),
            new("Actions", "link",     "링크"),
            new("Actions", "print",    "인쇄"),
        }),
        new("우선순위 (Priority)", new EmojiItem[]
        {
            new("Priority", "urgent",   "긴급"),
            new("Priority", "high",     "높음"),
            new("Priority", "medium",   "중간"),
            new("Priority", "low",      "낮음"),
            new("Priority", "pinned",   "고정"),
            new("Priority", "flagged",  "플래그"),
            new("Priority", "bookmark", "북마크"),
            new("Priority", "star",     "별표"),
            new("Priority", "warning",  "경고"),
            new("Priority", "locked",   "잠김"),
        }),
        new("사람 (People)", new EmojiItem[]
        {
            new("People", "person",        "사람"),
            new("People", "team",          "팀"),
            new("People", "mention",       "멘션"),
            new("People", "assign",        "지정"),
            new("People", "invite",        "초대"),
            new("People", "handshake",     "악수"),
            new("People", "group_chat",    "그룹 채팅"),
            new("People", "handover",      "인계"),
            new("People", "collaborators", "협업자"),
            new("People", "request",       "요청"),
        }),
        new("도구 (Tools)", new EmojiItem[]
        {
            new("Tools", "pen",        "펜"),
            new("Tools", "pencil",     "연필"),
            new("Tools", "clipboard",  "클립보드"),
            new("Tools", "calendar",   "달력"),
            new("Tools", "folder",     "폴더"),
            new("Tools", "document",   "문서"),
            new("Tools", "chart",      "차트"),
            new("Tools", "table",      "표"),
            new("Tools", "image",      "이미지"),
            new("Tools", "attachment", "첨부"),
        }),
        new("화살표 (Pointers)", new EmojiItem[]
        {
            new("Pointers", "up",      "위"),
            new("Pointers", "down",    "아래"),
            new("Pointers", "left",    "왼쪽"),
            new("Pointers", "right",   "오른쪽"),
            new("Pointers", "point",   "지시"),
            new("Pointers", "redo",    "다시"),
            new("Pointers", "undo",    "되돌리기"),
            new("Pointers", "expand",  "펼치기"),
            new("Pointers", "refresh", "새로고침"),
            new("Pointers", "callout", "콜아웃"),
        }),
        new("장식 (Decoration)", new EmojiItem[]
        {
            new("Decoration", "sparkle",  "반짝"),
            new("Decoration", "fire",     "불"),
            new("Decoration", "rocket",   "로켓"),
            new("Decoration", "crown",    "왕관"),
            new("Decoration", "trophy",   "트로피"),
            new("Decoration", "award",    "상장"),
            new("Decoration", "idea",     "아이디어"),
            new("Decoration", "target",   "타깃"),
            new("Decoration", "gem",      "보석"),
            new("Decoration", "confetti", "축포"),
        }),
    };
}

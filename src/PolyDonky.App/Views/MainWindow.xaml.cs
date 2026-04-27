using System;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
using PolyDonky.App.ViewModels;
using PolyDonky.Core;
using SR = PolyDonky.App.Properties.Resources;
using WpfMedia = System.Windows.Media;

namespace PolyDonky.App.Views;

public partial class MainWindow : Window
{
    private MainViewModel? _viewModel;
    private bool _suppressTextChanged;
    private DispatcherTimer? _statusTimer;

    // ── 글상자 드래그 생성 / 선택 상태 ────────────────────────────
    private bool _drawingTextBox;
    private bool _drawingInProgress;
    private Point _drawStart;
    private TextBoxShape _drawingShape = TextBoxShape.Rectangle;
    private TextBoxOverlay? _selectedOverlay;

    public MainWindow()
    {
        InitializeComponent();

        Loaded += OnLoaded;
        Unloaded += OnUnloaded;
        BodyEditor.TextChanged += OnEditorTextChanged;
        // 상태 표시줄 Insert/CapsLock/NumLock 갱신을 위해 윈도우 레벨 키 입력 가로채기.
        PreviewKeyDown += OnPreviewKeyDown;
        // 마지막으로 키보드 포커스를 가졌던 텍스트 편집기를 추적 — 메뉴 클릭 시 포커스가
        // 메뉴로 옮겨가도 직전 편집 컨텍스트(BodyEditor 또는 글상자 InnerEditor) 를 잃지 않게.
        BodyEditor.GotKeyboardFocus += (_, _) => _lastTextEditor = BodyEditor;
    }

    /// <summary>가장 최근에 키보드 포커스를 가졌던 RichTextBox.</summary>
    private RichTextBox? _lastTextEditor;

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        if (DataContext is MainViewModel vm)
        {
            _viewModel = vm;
            ApplyFlowDocument(vm.FlowDocument);
            vm.PropertyChanged += OnViewModelPropertyChanged;
            vm.FindReplaceRequested  += OnFindReplaceRequested;
            vm.SettingsRequested     += OnSettingsRequested;
            vm.OutlineStyleRequested += OnOutlineStyleRequested;
            vm.InsertTextBoxRequested += OnInsertTextBoxRequested;
            vm.RefreshSystemKeys();
            vm.RefreshMemoryUsage();
        }

        // RichTextBox 클릭 = 본문 편집 의도. 드래그 생성 모드가 아니면 글상자 선택 해제.
        BodyEditor.PreviewMouseLeftButtonDown += (_, _) =>
        {
            if (!_drawingTextBox) DeselectAllOverlays();
        };

        // 이모지·이미지 우클릭 → 속성 컨텍스트 메뉴, 더블클릭 → 속성 다이얼로그
        BodyEditor.ContextMenuOpening      += OnEmbeddedObjectContextMenuOpening;
        BodyEditor.PreviewMouseDoubleClick += OnEmbeddedObjectDoubleClick;

        _statusTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(1),
        };
        _statusTimer.Tick += OnStatusTimerTick;
        _statusTimer.Start();
    }

    private void OnUnloaded(object sender, RoutedEventArgs e)
    {
        if (_statusTimer is not null)
        {
            _statusTimer.Stop();
            _statusTimer.Tick -= OnStatusTimerTick;
            _statusTimer = null;
        }
    }

    private void OnStatusTimerTick(object? sender, EventArgs e)
    {
        _viewModel?.RefreshSystemKeys();
        _viewModel?.RefreshMemoryUsage();
    }

    private void OnPreviewKeyDown(object sender, KeyEventArgs e)
    {
        switch (e.Key)
        {
            case Key.Insert:
                _viewModel?.ToggleInsertMode();
                break;
            case Key.CapsLock:
            case Key.NumLock:
                // 토글 상태는 KeyDown 직후엔 Keyboard.IsKeyToggled 가 갱신 전이라
                // Dispatcher.BeginInvoke 로 다음 사이클에 읽는다.
                Dispatcher.BeginInvoke(new Action(() => _viewModel?.RefreshSystemKeys()),
                    DispatcherPriority.Input);
                break;
            case Key.Escape:
                if (_drawingTextBox)
                {
                    EndDrawingMode();
                    e.Handled = true;
                }
                else if (_selectedOverlay is not null)
                {
                    // 1단계: 안쪽 본문 편집 중이면 chrome 만 선택 상태로 전환 (포커스를 overlay 로 이동).
                    //        이후 Ctrl+C 가 글상자 자체를 복사하도록 해주는 진입점.
                    // 2단계: chrome 만 선택된 상태에서 다시 누르면 완전 해제.
                    if (_selectedOverlay.InnerEditor.IsKeyboardFocusWithin)
                    {
                        _selectedOverlay.Focus();
                        Keyboard.Focus(_selectedOverlay);
                    }
                    else
                    {
                        DeselectAllOverlays();
                    }
                    e.Handled = true;
                }
                break;

            // 글상자(부유 객체) 자체의 복사/잘라내기/붙여넣기.
            // 안쪽 본문(InnerEditor)이 포커스 중이면 가로채지 않고 RichTextBox 의
            // 일반 텍스트 클립보드 동작으로 넘긴다.
            case Key.C when (Keyboard.Modifiers & ModifierKeys.Control) != 0:
                if (TryCopySelectedFloatingObject()) e.Handled = true;
                break;
            case Key.X when (Keyboard.Modifiers & ModifierKeys.Control) != 0:
                if (TryCutSelectedFloatingObject()) e.Handled = true;
                break;
            case Key.V when (Keyboard.Modifiers & ModifierKeys.Control) != 0:
                if (TryPasteFloatingObject()) e.Handled = true;
                break;
        }
    }

    private void OnViewModelPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (_viewModel is null) return;
        if (e.PropertyName == nameof(MainViewModel.FlowDocument))
        {
            ApplyFlowDocument(_viewModel.FlowDocument);
            _viewModel.RefreshMemoryUsage();
        }
    }

    private void ApplyFlowDocument(System.Windows.Documents.FlowDocument fd)
    {
        // FlowDocument 는 RichTextBox 의 자식이지만 자체 시각 트리 루트라
        // RichTextBox.Foreground 가 Run 까지 전파되지 않는다.
        // SetResourceReference 로 테마 사전을 동적 바인딩 — 테마 교체 시 자동 갱신.
        fd.SetResourceReference(System.Windows.Documents.FlowDocument.ForegroundProperty, "OnSurface");
        // Background 는 PaperBorder 가 담당하므로 FlowDocument 는 투명으로 둔다.
        fd.Background = Brushes.Transparent;

        _suppressTextChanged = true;
        try
        {
            // 글자 방향은 추후 지원 예정 — 현재 항상 LTR.
            BodyEditor.Document = fd;
            BodyEditor.FlowDirection = FlowDirection.LeftToRight;
        }
        finally
        {
            _suppressTextChanged = false;
        }

        // 용지 크기·색상을 PaperBorder 에 반영
        var page = _viewModel?.Document.Sections.FirstOrDefault()?.Page;
        ApplyPageSettings(page);

        // 부유 객체 (글상자 등) 오버레이 재구축
        RebuildFloatingObjects();

        // 그림 오버레이 재구축 (InFrontOfText / BehindText 모드)
        RebuildOverlayImages();
    }

    /// <summary>
    /// Document 모델을 순회해 InFrontOfText / BehindText 모드 ImageBlock 들을
    /// OverlayImageCanvas / UnderlayImageCanvas 에 절대 위치로 배치한다.
    /// </summary>
    private void RebuildOverlayImages()
    {
        OverlayImageCanvas.Children.Clear();
        UnderlayImageCanvas.Children.Clear();

        if (_viewModel?.Document is null) return;

        foreach (var section in _viewModel.Document.Sections)
        foreach (var block in section.Blocks)
        {
            if (block is not PolyDonky.Core.ImageBlock img) continue;
            if (img.WrapMode is not (PolyDonky.Core.ImageWrapMode.InFrontOfText
                                  or PolyDonky.Core.ImageWrapMode.BehindText)) continue;

            var ctrl = Services.FlowDocumentBuilder.BuildOverlayImageControl(img);
            if (ctrl is null) continue;

            ctrl.Tag = img;   // 우클릭 → 속성 라우팅용
            System.Windows.Controls.Canvas.SetLeft(ctrl, Services.FlowDocumentBuilder.MmToDip(img.OverlayXMm));
            System.Windows.Controls.Canvas.SetTop(ctrl,  Services.FlowDocumentBuilder.MmToDip(img.OverlayYMm));
            ctrl.Cursor = System.Windows.Input.Cursors.SizeAll;

            // 우클릭 → 속성 다이얼로그
            ctrl.MouseRightButtonDown += OnOverlayImageRightClick;
            // 좌클릭 → 드래그 시작 또는 더블클릭 시 속성 다이얼로그
            ctrl.MouseLeftButtonDown += OnOverlayImageMouseDown;

            var canvas = img.WrapMode == PolyDonky.Core.ImageWrapMode.BehindText
                ? UnderlayImageCanvas
                : OverlayImageCanvas;
            canvas.Children.Add(ctrl);
        }
    }

    // ── 오버레이 그림 드래그 이동 상태 ────────────────────────────────
    private System.Windows.FrameworkElement? _draggingOverlayImage;
    private Point  _overlayDragStart;
    private double _overlayDragStartLeft;
    private double _overlayDragStartTop;
    private bool   _overlayDragMoved;

    private void OnOverlayImageRightClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (sender is not System.Windows.FrameworkElement fe ||
            fe.Tag is not PolyDonky.Core.ImageBlock img) return;

        var menu  = new System.Windows.Controls.ContextMenu();
        var item  = new System.Windows.Controls.MenuItem { Header = "속성(_P)..." };
        item.Click += (_, _) => OpenOverlayImageProperties(img);
        menu.Items.Add(item);
        menu.PlacementTarget = fe;
        menu.Placement = System.Windows.Controls.Primitives.PlacementMode.MousePoint;
        menu.IsOpen = true;
        e.Handled = true;
    }

    private void OnOverlayImageMouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (sender is not System.Windows.FrameworkElement fe) return;

        // 더블클릭 → 속성 다이얼로그
        if (e.ClickCount == 2 && fe.Tag is PolyDonky.Core.ImageBlock dblImg)
        {
            OpenOverlayImageProperties(dblImg);
            e.Handled = true;
            return;
        }

        // 단일 클릭 → 드래그 시작 (마우스 캡처)
        if (fe.Parent is not System.Windows.Controls.Canvas canvas) return;
        _draggingOverlayImage   = fe;
        _overlayDragStart       = e.GetPosition(canvas);
        _overlayDragStartLeft   = System.Windows.Controls.Canvas.GetLeft(fe);
        _overlayDragStartTop    = System.Windows.Controls.Canvas.GetTop(fe);
        if (double.IsNaN(_overlayDragStartLeft)) _overlayDragStartLeft = 0;
        if (double.IsNaN(_overlayDragStartTop))  _overlayDragStartTop  = 0;
        _overlayDragMoved = false;
        fe.CaptureMouse();
        fe.MouseMove         += OnOverlayImageDragMove;
        fe.MouseLeftButtonUp += OnOverlayImageDragUp;
        e.Handled = true;
    }

    private void OnOverlayImageDragMove(object sender, System.Windows.Input.MouseEventArgs e)
    {
        if (sender is not System.Windows.FrameworkElement fe ||
            !ReferenceEquals(_draggingOverlayImage, fe)) return;
        if (fe.Parent is not System.Windows.Controls.Canvas canvas) return;

        var pos = e.GetPosition(canvas);
        double dx = pos.X - _overlayDragStart.X;
        double dy = pos.Y - _overlayDragStart.Y;
        if (Math.Abs(dx) > 0.5 || Math.Abs(dy) > 0.5) _overlayDragMoved = true;

        // 페이지 좌상단 기준 음수 좌표는 허용하지 않는다.
        System.Windows.Controls.Canvas.SetLeft(fe, Math.Max(0, _overlayDragStartLeft + dx));
        System.Windows.Controls.Canvas.SetTop (fe, Math.Max(0, _overlayDragStartTop  + dy));
    }

    private void OnOverlayImageDragUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (sender is not System.Windows.FrameworkElement fe ||
            !ReferenceEquals(_draggingOverlayImage, fe)) return;

        fe.ReleaseMouseCapture();
        fe.MouseMove         -= OnOverlayImageDragMove;
        fe.MouseLeftButtonUp -= OnOverlayImageDragUp;
        _draggingOverlayImage = null;

        if (_overlayDragMoved && fe.Tag is PolyDonky.Core.ImageBlock img)
        {
            double left = System.Windows.Controls.Canvas.GetLeft(fe);
            double top  = System.Windows.Controls.Canvas.GetTop(fe);
            if (double.IsNaN(left)) left = 0;
            if (double.IsNaN(top))  top  = 0;
            img.OverlayXMm = Math.Max(0, Services.FlowDocumentBuilder.DipToMm(left));
            img.OverlayYMm = Math.Max(0, Services.FlowDocumentBuilder.DipToMm(top));
            _viewModel?.MarkDirty();
        }
        e.Handled = true;
    }

    private void OpenOverlayImageProperties(PolyDonky.Core.ImageBlock img)
    {
        var dlg = new ImagePropertiesWindow(img) { Owner = this };
        if (dlg.ShowDialog() == true)
        {
            // 속성 변경 — 본문 흐름의 placeholder 도, 캔버스 오버레이도 모두 갱신 필요.
            // 가장 안전한 방법: FlowDocument 전체 재빌드.
            _viewModel?.RebuildFlowDocument();
            _viewModel?.MarkDirty();
        }
    }

    private void ApplyPageSettings(PolyDonky.Core.PageSettings? page)
    {
        if (page is null) return;

        double padL = PolyDonky.App.Services.FlowDocumentBuilder.MmToDip(page.MarginLeftMm);
        double padT = PolyDonky.App.Services.FlowDocumentBuilder.MmToDip(page.MarginTopMm);
        double padR = PolyDonky.App.Services.FlowDocumentBuilder.MmToDip(page.MarginRightMm);
        double padB = PolyDonky.App.Services.FlowDocumentBuilder.MmToDip(page.MarginBottomMm);

        // 용지 너비·높이 (세로·가로 방향 보정).
        // Height 는 MinHeight 로 지정해 빈 문서도 한 페이지 분량으로 보이고, 본문이 길어지면 늘어난다.
        PaperBorder.Width     = PolyDonky.App.Services.FlowDocumentBuilder.MmToDip(page.EffectiveWidthMm);
        PaperBorder.MinHeight = PolyDonky.App.Services.FlowDocumentBuilder.MmToDip(page.EffectiveHeightMm);

        // 여백을 RichTextBox Padding 으로 반영 — FlowDocument.PagePadding 은 RichTextBox 컨텍스트에서 무시됨
        BodyEditor.Padding = new Thickness(padL, padT, padR, padB);

        // 여백 안내선 위치 갱신
        MarginGuideRect.Margin     = new Thickness(padL, padT, padR, padB);
        MarginGuideRect.Visibility = page.ShowMarginGuides
            ? System.Windows.Visibility.Visible
            : System.Windows.Visibility.Collapsed;

        // 용지 배경색 — 지정된 경우 SolidColorBrush, 없으면 테마 Surface 동적 리소스
        if (!string.IsNullOrEmpty(page.PaperColor))
        {
            try
            {
                // PolyDonky.Core.Color 와 충돌하므로 WpfMedia alias 로 명시.
                var c = (WpfMedia.Color)WpfMedia.ColorConverter.ConvertFromString(page.PaperColor)!;
                PaperBorder.Background = new SolidColorBrush(c);
            }
            catch
            {
                PaperBorder.SetResourceReference(System.Windows.Controls.Border.BackgroundProperty, "Surface");
            }
        }
        else
        {
            PaperBorder.SetResourceReference(System.Windows.Controls.Border.BackgroundProperty, "Surface");
        }
    }

    private void OnEditorTextChanged(object sender, TextChangedEventArgs e)
    {
        if (_suppressTextChanged) return;
        _viewModel?.MarkDirty();
    }

    private void OnFindReplaceRequested(object? sender, EventArgs e)
    {
        var dlg = new FindReplaceWindow(BodyEditor) { Owner = this };
        dlg.Show();
    }

    private void OnSettingsRequested(object? sender, EventArgs e)
    {
        var dlg = new SettingsWindow { Owner = this };
        dlg.ShowDialog();
    }

    private void OnOutlineStyleRequested(object? sender, EventArgs e)
    {
        var current = (_viewModel is { } vm2 && vm2.Document?.OutlineStyles != null)
            ? vm2.Document.OutlineStyles
            : PolyDonky.Core.OutlineStyleSet.CreateDefault();
        var dlg = new OutlineStyleWindow(current) { Owner = this };
        dlg.StyleApplied += (_, styleSet) => _viewModel?.ApplyOutlineStyles(styleSet);
        dlg.ShowDialog();
    }

    /// <summary>
    /// 글자/문단 속성 다이얼로그가 적용해야 할 활성 RichTextBox 를 결정한다.
    /// 우선순위:
    /// 1) 글상자가 선택된 상태(`_selectedOverlay != null`) — 사용자가 chrome 만 선택했든
    ///    안쪽 본문을 편집 중이었든 의도는 그 글상자에 작업하는 것. 해당 InnerEditor 사용.
    /// 2) 마지막으로 키보드 포커스를 가졌던 RichTextBox — 메뉴 클릭으로 포커스가 일시
    ///    이동한 경우에도 직전 편집 대상을 보존.
    /// 3) BodyEditor — 기본 폴백.
    /// </summary>
    private RichTextBox GetActiveTextEditor()
        => _selectedOverlay?.InnerEditor ?? _lastTextEditor ?? BodyEditor;

    /// <summary>
    /// 글상자 InnerEditor 를 다이얼로그 대상으로 잡았는데 안쪽 selection 이 비어 있으면
    /// 전체 선택해서 적용 — chrome 만 선택한 사용자도 눈에 보이는 결과를 얻을 수 있게.
    /// 본문(BodyEditor) 에는 적용하지 않는다 — 본문에서 빈 selection 은 "다음 입력에
    /// 적용" 의미가 명확.
    /// </summary>
    private static void EnsureInnerSelectionForDialog(RichTextBox editor, TextBoxOverlay? overlay)
    {
        if (overlay is null) return;
        if (!ReferenceEquals(editor, overlay.InnerEditor)) return;
        if (!editor.Selection.IsEmpty) return;
        editor.SelectAll();
    }

    private void OnFormatChar(object sender, RoutedEventArgs e)
    {
        var editor = GetActiveTextEditor();
        EnsureInnerSelectionForDialog(editor, _selectedOverlay);
        var dlg = new CharFormatWindow(editor) { Owner = this };
        if (dlg.ShowDialog() == true)
        {
            _viewModel?.MarkDirty();
            editor.Focus();
        }
    }

    private void OnFormatPara(object sender, RoutedEventArgs e)
    {
        var editor = GetActiveTextEditor();
        EnsureInnerSelectionForDialog(editor, _selectedOverlay);
        var dlg = new ParaFormatWindow(editor) { Owner = this };
        if (dlg.ShowDialog() == true)
        {
            _viewModel?.MarkDirty();
            editor.Focus();
        }
    }

    private void OnFormatPage(object sender, RoutedEventArgs e)
    {
        var current = _viewModel?.Document.Sections.FirstOrDefault()?.Page
                      ?? new PolyDonky.Core.PageSettings();
        var dlg = new PageFormatWindow(current) { Owner = this };
        if (dlg.ShowDialog() == true)
        {
            if (_viewModel?.Document.Sections.FirstOrDefault() is { } section)
                section.Page = dlg.ResultSettings;
            _viewModel?.RebuildFlowDocument();
            _viewModel?.MarkDirty();
        }
    }

    private void OnInsertSpecialChar(object sender, RoutedEventArgs e)
    {
        var dlg = new SpecialCharWindow(BodyEditor) { Owner = this };
        if (dlg.ShowDialog() == true)
        {
            _viewModel?.MarkDirty();
            BodyEditor.Focus();
        }
    }

    private void OnInsertEquation(object sender, RoutedEventArgs e)
    {
        var dlg = new EquationWindow(BodyEditor) { Owner = this };
        if (dlg.ShowDialog() == true)
        {
            _viewModel?.MarkDirty();
            BodyEditor.Focus();
        }
    }

    private void OnInsertEmoji(object sender, RoutedEventArgs e)
    {
        // 글상자가 선택되어 있으면 안쪽 InnerEditor 로, 아니면 본문으로 라우팅.
        // 글상자 안쪽 selection 비어 있을 때 SelectAll 강제는 하지 않는다 — 이모지는 캐럿
        // 위치에 삽입되는 객체이므로 의도치 않게 본문 전체를 대체하면 안 된다.
        var editor = GetActiveTextEditor();
        var dlg = new EmojiWindow(editor) { Owner = this };
        if (dlg.ShowDialog() == true)
        {
            _viewModel?.MarkDirty();
            editor.Focus();
        }
    }

    private void OnInsertImage(object sender, RoutedEventArgs e)
    {
        // 그림은 블록 단위 삽입 — 본문 편집기에만 삽입하고 글상자 안은 지원하지 않는다.
        var dlg = new ImageWindow(BodyEditor) { Owner = this };
        if (dlg.ShowDialog() == true)
        {
            _viewModel?.MarkDirty();
            BodyEditor.Focus();
        }
    }

    // ── 이모지·이미지 속성 편집 ─────────────────────────────────────────

    private void OnEmbeddedObjectContextMenuOpening(object sender, System.Windows.Controls.ContextMenuEventArgs e)
    {
        // 기본 컨텍스트 메뉴(잘라내기/복사/붙여넣기)를 그대로 두고,
        // 이모지·이미지 위에서 우클릭한 경우에만 구분선 + 속성 항목을 동적으로 추가.
        var pt = System.Windows.Input.Mouse.GetPosition(BodyEditor);
        if (FindEmbeddedObjectAt(e.OriginalSource, pt) is not { } found) return;

        var menu = BodyEditor.ContextMenu;
        if (menu is null) return;

        var sep  = new System.Windows.Controls.Separator();
        var item = new System.Windows.Controls.MenuItem { Header = "속성(_P)..." };
        item.Click += (_, _) => OpenEmbeddedObjectProperties(found.img, found.container);

        menu.Items.Add(sep);
        menu.Items.Add(item);

        // 메뉴 닫힘 시 동적으로 추가한 항목을 제거 — 다음 일반 우클릭에 속성이 남지 않도록.
        void Cleanup(object? s, System.Windows.RoutedEventArgs _ev)
        {
            menu.Closed -= Cleanup;
            menu.Items.Remove(sep);
            menu.Items.Remove(item);
        }
        menu.Closed += Cleanup;
    }

    private void OnEmbeddedObjectDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        var pt = e.GetPosition(BodyEditor);
        if (FindEmbeddedObjectAt(e.OriginalSource, pt) is not { } found) return;

        OpenEmbeddedObjectProperties(found.img, found.container);
        e.Handled = true;
    }

    /// <summary>
    /// OriginalSource 비주얼 트리 + 마우스 위치 InputHitTest 양쪽으로 이모지·이미지 Image 컨트롤을 찾는다.
    /// RichTextBox 내부 hosting 으로 OriginalSource 가 Image 까지 닿지 않는 경우의 폴백.
    /// </summary>
    private (System.Windows.Controls.Image img, object container)? FindEmbeddedObjectAt(
        object? originalSource, System.Windows.Point pt)
    {
        if (TryWalkUpToImage(originalSource as System.Windows.DependencyObject) is { } a) return a;
        if (TryWalkUpToImage(BodyEditor.InputHitTest(pt) as System.Windows.DependencyObject) is { } b) return b;
        return null;
    }

    private static (System.Windows.Controls.Image img, object container)? TryWalkUpToImage(
        System.Windows.DependencyObject? dep)
    {
        // 1단계: 비주얼 트리에서 Image 찾기.
        // VisualTreeHelper.GetParent 는 Visual/Visual3D 만 허용. InputHitTest 는
        // ContentElement(Run, Paragraph, FlowDocument 등)를 반환할 수 있으므로
        // Visual 인지 확인 후 진행한다.
        while (dep is not null)
        {
            if (dep is System.Windows.Controls.Image img)
            {
                // 2단계: Image 의 logical tree 부모를 따라 올라가며 BUC/IUC 컨테이너 찾기.
                // (Image.Tag 에 컨테이너를 저장하면 container.Child = image 와 순환 참조가 되어
                //  WPF undo 의 XamlWriter.Save() 가 StackOverflow 로 폭주한다.)
                System.Windows.DependencyObject? logical = img;
                while (logical is not null)
                {
                    // BlockUIContainer (인라인 모드 그림)
                    if (logical is System.Windows.Documents.BlockUIContainer buc &&
                        buc.Tag is PolyDonky.Core.ImageBlock)
                        return (img, buc);
                    // Paragraph (래핑 모드 그림 — Floater 가 든 단락의 Tag 에 ImageBlock 보존)
                    if (logical is System.Windows.Documents.Paragraph wrappedPara &&
                        wrappedPara.Tag is PolyDonky.Core.ImageBlock)
                        return (img, wrappedPara);
                    // InlineUIContainer (이모지)
                    if (logical is System.Windows.Documents.InlineUIContainer iuc &&
                        iuc.Tag is PolyDonky.Core.Run { EmojiKey: { Length: > 0 } })
                        return (img, iuc);
                    logical = System.Windows.LogicalTreeHelper.GetParent(logical);
                }
                return null;
            }
            // ContentElement (FlowDocument, Paragraph, Run …) 는 Visual 이 아니므로
            // VisualTreeHelper.GetParent 를 호출하면 InvalidOperationException 이 발생한다.
            if (dep is not Visual) return null;
            dep = VisualTreeHelper.GetParent(dep);
        }
        return null;
    }

    private void OpenEmbeddedObjectProperties(System.Windows.Controls.Image imgControl, object container)
    {
        if (container is System.Windows.Documents.InlineUIContainer iuc &&
            iuc.Tag is PolyDonky.Core.Run emojiRun)
        {
            var dlg = new EmojiPropertiesWindow(imgControl, iuc, emojiRun) { Owner = this };
            if (dlg.ShowDialog() == true)
                _viewModel?.MarkDirty();
        }
        else if (container is System.Windows.Documents.Block oldBlock &&
                 GetImageBlockFromBlock(oldBlock) is { } imageBlock)
        {
            var dlg = new ImagePropertiesWindow(imageBlock) { Owner = this };
            if (dlg.ShowDialog() == true)
            {
                // WrapMode 변경 시 컨테이너 종류가 달라질 수 있음
                // (BlockUIContainer ↔ Paragraph+Floater ↔ 빈 placeholder for overlay).
                // BuildImage 가 적절한 Block 타입을 반환하므로 그걸 그대로 교체.
                var newBlock = Services.FlowDocumentBuilder.BuildImage(imageBlock);
                var doc      = BodyEditor.Document;
                doc.Blocks.InsertBefore(oldBlock, newBlock);
                doc.Blocks.Remove(oldBlock);
                // 캔버스 오버레이도 갱신 — 모드 전환(예: Inline → BehindText) 반영.
                RebuildOverlayImages();
                _viewModel?.MarkDirty();
            }
        }
    }

    private static PolyDonky.Core.ImageBlock? GetImageBlockFromBlock(System.Windows.Documents.Block block) =>
        block switch
        {
            System.Windows.Documents.BlockUIContainer buc when buc.Tag is PolyDonky.Core.ImageBlock ib => ib,
            System.Windows.Documents.Paragraph p          when p.Tag   is PolyDonky.Core.ImageBlock ib => ib,
            _ => null,
        };

    // 편집 > 지우기: RichTextBox 는 ApplicationCommands.Delete 를 자체 바인딩하지 않으므로
    // 메뉴에서 직접 호출 시 동작하도록 선택 영역을 지운다. 선택이 비어 있으면 캐럿 직후
    // 한 글자(EditingCommands.Delete) 를 지우는 일반 워드프로세서 동작을 따른다.
    private void OnEditDelete(object sender, RoutedEventArgs e)
    {
        if (!BodyEditor.Selection.IsEmpty)
        {
            BodyEditor.Selection.Text = string.Empty;
        }
        else
        {
            System.Windows.Documents.EditingCommands.Delete.Execute(null, BodyEditor);
        }
        BodyEditor.Focus();
    }

    // 쓰기 보호: IsReadOnly 가 켜져 있으면 RichTextBox 가 자체적으로 입력을 막지만,
    // 사용자에게 "왜 안 되지?" 의 침묵 대신 비밀번호 프롬프트를 띄워 즉시 잠금 해제 흐름으로 안내한다.
    // 검증 성공 시 ViewModel 이 IsWriteProtected=false 로 풀고, 다음 키부터 바로 편집된다.
    private void OnEditorPreviewTextInput(object sender, TextCompositionEventArgs e)
    {
        if (_viewModel?.IsWriteProtected != true) return;
        e.Handled = true;
        _viewModel.TryUnlockForEditing();
    }

    private void OnEditorPreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (_viewModel?.IsWriteProtected != true) return;
        if (!IsEditingIntent(e)) return;
        e.Handled = true;
        _viewModel.TryUnlockForEditing();
    }

    /// <summary>
    /// 입력 키가 "편집 의도" 인지 판별. 화살표·Home/End·Ctrl+C/Ctrl+A 등 읽기 전용 동작은 false.
    /// (PreviewTextInput 이 일반 문자 입력은 별도로 처리하므로 여기서는 특수 편집 키만 본다.)
    /// </summary>
    private static bool IsEditingIntent(KeyEventArgs e)
    {
        var ctrl = (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control;
        return e.Key switch
        {
            Key.Back or Key.Delete or Key.Enter or Key.Tab => true,
            Key.V or Key.X when ctrl                       => true,  // 붙여넣기 / 잘라내기
            _                                              => false,
        };
    }

    private void OnDragOver(object sender, DragEventArgs e)
    {
        if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;
        e.Effects = DragDropEffects.Copy;
        e.Handled = true;
    }

    private void OnDrop(object sender, DragEventArgs e)
    {
        if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;
        if (e.Data.GetData(DataFormats.FileDrop) is string[] files && files.Length > 0)
        {
            _viewModel?.OpenFile(files[0]);
        }
        e.Handled = true;
    }

    // ── 확대/축소 조절기 ──────────────────────────────────────────────────────

    private void OnFitToWidth(object sender, RoutedEventArgs e)
    {
        if (_viewModel is null) return;
        // 뷰포트 너비에서 PaperStackPanel 좌우 여백(32+32) 을 빼고 용지 논리 너비로 나눈다.
        const double hMargin = 64;
        double paperW = PaperBorder.Width;
        double viewW  = EditorScrollViewer.ViewportWidth;
        if (paperW <= 0 || viewW <= 0) return;
        _viewModel.ZoomPercent = (viewW - hMargin) / paperW * 100;
    }

    private void OnFitToPage(object sender, RoutedEventArgs e)
    {
        if (_viewModel is null) return;
        const double hMargin = 64; // StackPanel 좌우 여백
        const double vMargin = 76; // StackPanel 상(28)+하(48) 여백
        double paperW = PaperBorder.Width;
        double paperH = PaperBorder.MinHeight; // 콘텐츠가 길어도 한 페이지 분량 기준
        double viewW  = EditorScrollViewer.ViewportWidth;
        double viewH  = EditorScrollViewer.ViewportHeight;
        if (paperW <= 0 || paperH <= 0 || viewW <= 0 || viewH <= 0) return;
        double fitW = (viewW - hMargin) / paperW;
        double fitH = (viewH - vMargin) / paperH;
        _viewModel.ZoomPercent = Math.Min(fitW, fitH) * 100;
    }

    private void OnZoomTextBoxKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key != Key.Enter) return;
        ApplyZoomFromTextBox();
        BodyEditor.Focus();
        e.Handled = true;
    }

    private void OnZoomTextBoxLostFocus(object sender, RoutedEventArgs e)
        => ApplyZoomFromTextBox();

    private void ApplyZoomFromTextBox()
    {
        if (_viewModel is null) return;
        var text = TxtZoom.Text.Trim().TrimEnd('%');
        if (double.TryParse(text, out double val))
            _viewModel.ZoomPercent = val;
        else
            TxtZoom.Text = _viewModel.ZoomPercent.ToString("0");
    }

    // ── 글상자 (입력 > 글상자) ──────────────────────────────────────────────

    private void OnInsertTextBoxRequested(object? sender, TextBoxShape shape)
    {
        _drawingTextBox = true;
        _drawingShape   = shape;
        Mouse.OverrideCursor = Cursors.Cross;
        if (_viewModel is not null)
            _viewModel.StatusMessage = SR.StatusDrawTextBox;
    }

    private void EndDrawingMode()
    {
        _drawingTextBox = false;
        _drawingInProgress = false;
        Mouse.OverrideCursor = null;
        DrawPreviewRect.Visibility = Visibility.Collapsed;
        if (PaperBorder.IsMouseCaptured) PaperBorder.ReleaseMouseCapture();
        if (_viewModel is not null) _viewModel.StatusMessage = SR.StatusReady;
    }

    private void OnPaperPreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (!_drawingTextBox) return;

        var pos = e.GetPosition(PaperBorder);
        pos.X = Math.Clamp(pos.X, 0, PaperBorder.ActualWidth);
        pos.Y = Math.Clamp(pos.Y, 0, PaperBorder.ActualHeight);

        _drawStart = pos;
        _drawingInProgress = true;

        Canvas.SetLeft(DrawPreviewRect, pos.X);
        Canvas.SetTop(DrawPreviewRect, pos.Y);
        DrawPreviewRect.Width = 0;
        DrawPreviewRect.Height = 0;
        DrawPreviewRect.Visibility = Visibility.Visible;

        PaperBorder.CaptureMouse();
        e.Handled = true;
    }

    private void OnPaperPreviewMouseMove(object sender, MouseEventArgs e)
    {
        if (!_drawingInProgress) return;

        var pos = e.GetPosition(PaperBorder);
        pos.X = Math.Clamp(pos.X, 0, PaperBorder.ActualWidth);
        pos.Y = Math.Clamp(pos.Y, 0, PaperBorder.ActualHeight);

        double x = Math.Min(_drawStart.X, pos.X);
        double y = Math.Min(_drawStart.Y, pos.Y);
        double w = Math.Abs(pos.X - _drawStart.X);
        double h = Math.Abs(pos.Y - _drawStart.Y);

        Canvas.SetLeft(DrawPreviewRect, x);
        Canvas.SetTop(DrawPreviewRect, y);
        DrawPreviewRect.Width  = w;
        DrawPreviewRect.Height = h;
    }

    private void OnPaperPreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (!_drawingInProgress) return;

        var pos = e.GetPosition(PaperBorder);
        pos.X = Math.Clamp(pos.X, 0, PaperBorder.ActualWidth);
        pos.Y = Math.Clamp(pos.Y, 0, PaperBorder.ActualHeight);

        double x = Math.Min(_drawStart.X, pos.X);
        double y = Math.Min(_drawStart.Y, pos.Y);
        double w = Math.Abs(pos.X - _drawStart.X);
        double h = Math.Abs(pos.Y - _drawStart.Y);

        EndDrawingMode();

        // 너무 작은 드래그(클릭에 가까움) 는 생성 안 함.
        const double minDip = 10;
        if (w < minDip || h < minDip)
        {
            e.Handled = true;
            return;
        }

        var model = new TextBoxObject
        {
            Shape    = _drawingShape,
            XMm      = x / TextBoxOverlay.DipsPerMm,
            YMm      = y / TextBoxOverlay.DipsPerMm,
            WidthMm  = w / TextBoxOverlay.DipsPerMm,
            HeightMm = h / TextBoxOverlay.DipsPerMm,
            Status   = NodeStatus.Modified,
        };
        _viewModel?.AddFloatingObjectToCurrentSection(model);
        var overlay = AddTextBoxOverlay(model);
        SelectOverlay(overlay);
        overlay.BeginEditing();

        e.Handled = true;
    }

    // ── 글상자(부유 객체) 복사/잘라내기/붙여넣기 ──────────────────────
    private const string FloatingObjectClipboardFormat = "PolyDonky.FloatingObject.v1";

    /// <summary>
    /// 선택된 글상자를 복사한다. 안쪽 본문에 포커스가 있어도 텍스트 선택이 비어 있으면
    /// "글상자 자체 복사" 의도로 간주 — Word/PowerPoint 와 동일한 mental model.
    /// 안쪽 본문에 텍스트 선택이 있으면 가로채지 않고 일반 복사에 양보.
    /// </summary>
    private bool TryCopySelectedFloatingObject()
    {
        if (_selectedOverlay is null) return false;
        if (_selectedOverlay.InnerEditor.IsKeyboardFocusWithin
            && !_selectedOverlay.InnerEditor.Selection.IsEmpty)
            return false;

        var json = System.Text.Json.JsonSerializer.Serialize<FloatingObject>(
            _selectedOverlay.Model, JsonDefaults.Options);
        var dataObj = new System.Windows.DataObject();
        dataObj.SetData(FloatingObjectClipboardFormat, json);
        // Plain-text 폴백 — 다른 앱으로 붙여넣기 시 안쪽 텍스트만 가도록.
        dataObj.SetText(_selectedOverlay.Model.GetPlainText());
        Clipboard.SetDataObject(dataObj, copy: true);
        return true;
    }

    private bool TryCutSelectedFloatingObject()
    {
        if (!TryCopySelectedFloatingObject()) return false;
        var overlay = _selectedOverlay!;
        FloatingCanvas.Children.Remove(overlay);
        _viewModel?.RemoveFloatingObject(overlay.Model);
        _selectedOverlay = null;
        BodyEditor.Focus();
        return true;
    }

    private bool TryPasteFloatingObject()
    {
        if (!Clipboard.ContainsData(FloatingObjectClipboardFormat)) return false;
        // 텍스트 선택이 있을 때만 일반 텍스트 붙여넣기에 양보 (사용자 의도가 텍스트 교체).
        // 캐럿만 위치한 경우(=텍스트 선택 없음) 는 BodyEditor 든 InnerEditor 든
        // 글상자 클립보드 데이터를 우선 적용 — 사용자가 방금 글상자를 복사했다면
        // Ctrl+V 의 자연스러운 결과는 새 글상자 한 개를 캔버스에 띄우는 것.
        if (BodyEditor.IsKeyboardFocusWithin && !BodyEditor.Selection.IsEmpty) return false;
        if (_selectedOverlay?.InnerEditor.IsKeyboardFocusWithin == true
            && !_selectedOverlay.InnerEditor.Selection.IsEmpty)
            return false;

        var json = Clipboard.GetData(FloatingObjectClipboardFormat) as string;
        if (string.IsNullOrEmpty(json)) return false;

        FloatingObject? clone;
        try
        {
            clone = System.Text.Json.JsonSerializer.Deserialize<FloatingObject>(json, JsonDefaults.Options);
        }
        catch
        {
            return false;
        }
        if (clone is not TextBoxObject tb) return false;

        // 새 인스턴스 표시 — Id 재발급, 위치는 살짝 오프셋.
        tb.Id = null;
        tb.XMm += 5;
        tb.YMm += 5;
        tb.Status = NodeStatus.Modified;

        _viewModel?.AddFloatingObjectToCurrentSection(tb);
        var overlay = AddTextBoxOverlay(tb);
        SelectOverlay(overlay);
        return true;
    }

    /// <summary>현재 섹션의 FloatingObjects 를 캔버스에 다시 채워 그린다 (문서 로드 시).</summary>
    private void RebuildFloatingObjects()
    {
        FloatingCanvas.Children.Clear();
        _selectedOverlay = null;
        // 옛 InnerEditor 참조는 기각 — 다시 만들 overlay 의 InnerEditor 가 GotKeyboardFocus 시 갱신.
        _lastTextEditor = null;
        var section = _viewModel?.Document.Sections.FirstOrDefault();
        if (section is null) return;
        foreach (var obj in section.FloatingObjects.OfType<TextBoxObject>())
        {
            AddTextBoxOverlay(obj);
        }
    }

    private TextBoxOverlay AddTextBoxOverlay(TextBoxObject model)
    {
        var overlay = new TextBoxOverlay(model);
        Canvas.SetLeft(overlay, model.XMm * TextBoxOverlay.DipsPerMm);
        Canvas.SetTop(overlay,  model.YMm * TextBoxOverlay.DipsPerMm);
        overlay.Width  = model.WidthMm  * TextBoxOverlay.DipsPerMm;
        overlay.Height = model.HeightMm * TextBoxOverlay.DipsPerMm;

        overlay.Selected += (_, _) => SelectOverlay(overlay);
        // 마지막 포커스 추적 — 사용자가 메뉴(서식 → 글자 속성 등) 를 누르면 포커스가 메뉴로
        // 이동해 IsKeyboardFocusWithin 이 false 가 된다. _lastTextEditor 에 미리 기억해두면
        // 메뉴에서 연 다이얼로그가 정확한 편집 대상을 잡을 수 있다.
        overlay.InnerEditor.GotKeyboardFocus += (_, _) => _lastTextEditor = overlay.InnerEditor;

        overlay.BringForwardRequested += (_, _) =>
        {
            int idx = FloatingCanvas.Children.IndexOf(overlay);
            if (idx < FloatingCanvas.Children.Count - 1)
            {
                FloatingCanvas.Children.RemoveAt(idx);
                FloatingCanvas.Children.Insert(idx + 1, overlay);
                model.ZOrder++;
                _viewModel?.NotifyFloatingObjectChanged();
            }
        };

        overlay.SendBackRequested += (_, _) =>
        {
            int idx = FloatingCanvas.Children.IndexOf(overlay);
            if (idx > 0)
            {
                FloatingCanvas.Children.RemoveAt(idx);
                FloatingCanvas.Children.Insert(idx - 1, overlay);
                model.ZOrder--;
                _viewModel?.NotifyFloatingObjectChanged();
            }
        };

        overlay.AppearanceChangedCommitted += (_, _) => _viewModel?.NotifyFloatingObjectChanged();

        overlay.GeometryChangedCommitted += (_, _) =>
        {
            // Canvas DIP → mm 동기화. NaN 방어 (SetLeft 직후라 정상이지만 안전).
            double left = Canvas.GetLeft(overlay); if (double.IsNaN(left)) left = 0;
            double top  = Canvas.GetTop(overlay);  if (double.IsNaN(top))  top  = 0;
            model.XMm      = left / TextBoxOverlay.DipsPerMm;
            model.YMm      = top  / TextBoxOverlay.DipsPerMm;
            model.WidthMm  = overlay.ActualWidth  / TextBoxOverlay.DipsPerMm;
            model.HeightMm = overlay.ActualHeight / TextBoxOverlay.DipsPerMm;
            model.Status   = NodeStatus.Modified;
            _viewModel?.NotifyFloatingObjectChanged();
        };

        overlay.ContentChangedCommitted += (_, _) => _viewModel?.NotifyFloatingObjectChanged();

        overlay.DeleteRequested += (_, _) =>
        {
            FloatingCanvas.Children.Remove(overlay);
            _viewModel?.RemoveFloatingObject(model);
            if (ReferenceEquals(_selectedOverlay, overlay)) _selectedOverlay = null;
            if (ReferenceEquals(_lastTextEditor, overlay.InnerEditor)) _lastTextEditor = null;
            BodyEditor.Focus();
        };

        FloatingCanvas.Children.Add(overlay);
        return overlay;
    }

    private void SelectOverlay(TextBoxOverlay overlay)
    {
        if (!ReferenceEquals(_selectedOverlay, overlay) && _selectedOverlay is not null)
            _selectedOverlay.IsSelected = false;
        _selectedOverlay = overlay;
        overlay.IsSelected = true;
    }

    private void DeselectAllOverlays()
    {
        if (_selectedOverlay is null) return;
        _selectedOverlay.IsSelected = false;
        _selectedOverlay = null;
    }

}

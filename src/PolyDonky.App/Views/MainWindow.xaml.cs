using System;
using System.Collections.Generic;
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
using WpfShapes = System.Windows.Shapes;

namespace PolyDonky.App.Views;

public partial class MainWindow : Window
{
    private MainViewModel? _viewModel;
    private bool _suppressTextChanged;
    private DispatcherTimer? _statusTimer;

    // ── 임베드 이미지 사용자 드래그 ───────────────────────────────────────
    // WPF 의 Floater 드래그는 HorizontalAlignment 를 조용히 바꾸고, 클립보드 직렬화 과정에서
    // Tag(ImageBlock 참조)를 잃어버린다. PreviewMouseMove 에서 WPF 의 드래그를 완전 차단하고
    // 대신 직접 드래그 로직을 구현한다.
    // - MouseDown: 이미지 위인지 확인, Block + ImageBlock 참조 캡처
    // - MouseMove: 임계거리 초과 시 active 상태로 전환, 커서 변경
    // - MouseUp  : 드롭 X 좌표 기반으로 HAlign / WrapMode 결정 → Block 재구축
    private bool                            _suppressEmbeddedObjectDrag;
    private PolyDonky.Core.ImageBlock?      _embeddedDragModel;
    private System.Windows.Documents.Block? _embeddedDragBlock;
    private bool                            _embeddedDragActive;
    private Point                           _embeddedDragOrigin;

    private void OnEditorPreviewMouseDownTrackDrag(object sender, MouseButtonEventArgs e)
    {
        _embeddedDragModel  = null;
        _embeddedDragBlock  = null;
        _embeddedDragActive = false;

        if (!_drawingTextBox) DeselectAllOverlays();
        var pt = e.GetPosition(BodyEditor);

        // Alt + 클릭 → BehindText 그림(BodyEditor 뒤 UnderlayImageCanvas) 드래그 시작.
        // 일반 클릭은 본문 텍스트 선택을 위해 양보 — 텍스트 위에 그림이 깔린 영역에서도 편집 가능.
        if ((Keyboard.Modifiers & ModifierKeys.Alt) != 0 &&
            FindCanvasChildAt(UnderlayImageCanvas, pt) is { } underlayCtrl)
        {
            StartUnderlayImageDrag(underlayCtrl, e);
            e.Handled = true;
            return;
        }

        var found = FindEmbeddedObjectAt(e.OriginalSource as System.Windows.DependencyObject, pt);
        if (found is { container: System.Windows.Documents.Block blk } &&
            GetImageBlockFromBlock(blk) is { } imgModel)
        {
            _suppressEmbeddedObjectDrag = true;
            _embeddedDragModel  = imgModel;
            _embeddedDragBlock  = blk;
            _embeddedDragOrigin = pt;
        }
        else
        {
            _suppressEmbeddedObjectDrag = false;
        }
    }

    // ── BehindText 그림 Alt+드래그 ───────────────────────────────────────
    // UnderlayImageCanvas 자식은 BodyEditor 뒤에 있어 일반적으로 마우스를 받을 수 없다.
    // BodyEditor 가 mouse capture 를 가지고 이벤트를 underlay 그림에 직접 라우팅한다.
    private void StartUnderlayImageDrag(System.Windows.FrameworkElement fe, MouseButtonEventArgs e)
    {
        if (fe.Parent is not System.Windows.Controls.Canvas canvas) return;

        // 더블클릭 → 속성 다이얼로그
        if (e.ClickCount == 2 && fe.Tag is PolyDonky.Core.ImageBlock dblImg)
        {
            OpenOverlayImageProperties(dblImg);
            return;
        }

        _draggingOverlayImage = fe;
        _overlayDragStart     = e.GetPosition(canvas);
        _overlayDragStartLeft = System.Windows.Controls.Canvas.GetLeft(fe);
        _overlayDragStartTop  = System.Windows.Controls.Canvas.GetTop(fe);
        if (double.IsNaN(_overlayDragStartLeft)) _overlayDragStartLeft = 0;
        if (double.IsNaN(_overlayDragStartTop))  _overlayDragStartTop  = 0;
        _overlayDragMoved = false;
        BodyEditor.CaptureMouse();
        BodyEditor.MouseMove         += OnUnderlayImageDragMove;
        BodyEditor.MouseLeftButtonUp += OnUnderlayImageDragUp;
    }

    private void OnUnderlayImageDragMove(object sender, MouseEventArgs e)
    {
        if (_draggingOverlayImage is not { } fe) return;
        if (fe.Parent is not System.Windows.Controls.Canvas canvas) return;
        var pos = e.GetPosition(canvas);
        double dx = pos.X - _overlayDragStart.X;
        double dy = pos.Y - _overlayDragStart.Y;
        if (Math.Abs(dx) > 0.5 || Math.Abs(dy) > 0.5) _overlayDragMoved = true;
        System.Windows.Controls.Canvas.SetLeft(fe, _overlayDragStartLeft + dx);
        System.Windows.Controls.Canvas.SetTop (fe, _overlayDragStartTop  + dy);
    }

    private void OnUnderlayImageDragUp(object sender, MouseButtonEventArgs e)
    {
        if (_draggingOverlayImage is not { } fe) return;
        BodyEditor.ReleaseMouseCapture();
        BodyEditor.MouseMove         -= OnUnderlayImageDragMove;
        BodyEditor.MouseLeftButtonUp -= OnUnderlayImageDragUp;
        _draggingOverlayImage = null;

        if (_overlayDragMoved && fe.Tag is PolyDonky.Core.ImageBlock img)
        {
            double left = System.Windows.Controls.Canvas.GetLeft(fe);
            double top  = System.Windows.Controls.Canvas.GetTop(fe);
            if (double.IsNaN(left)) left = 0;
            if (double.IsNaN(top))  top  = 0;
            img.OverlayXMm = Services.FlowDocumentBuilder.DipToMm(left);
            img.OverlayYMm = Services.FlowDocumentBuilder.DipToMm(top);
            _viewModel?.MarkDirty();
        }
        e.Handled = true;
    }

    private void OnEditorPreviewMouseMoveBlockDrag(object sender, MouseEventArgs e)
    {
        if (!_suppressEmbeddedObjectDrag) return;
        if (e.LeftButton != MouseButtonState.Pressed)
        {
            _embeddedDragActive = false;
            _suppressEmbeddedObjectDrag = false;
            Mouse.OverrideCursor = null;
            return;
        }
        // WPF 기본 드래그(Floater 위치 변경 + Tag 손실)를 완전 차단.
        e.Handled = true;

        var pt = e.GetPosition(BodyEditor);
        if (!_embeddedDragActive &&
            (Math.Abs(pt.X - _embeddedDragOrigin.X) > 8 ||
             Math.Abs(pt.Y - _embeddedDragOrigin.Y) > 8))
        {
            _embeddedDragActive  = true;
            Mouse.OverrideCursor = Cursors.SizeAll;
        }
    }

    private void OnEditorPreviewMouseUpEmbedded(object sender, MouseButtonEventArgs e)
    {
        Mouse.OverrideCursor = null;
        _suppressEmbeddedObjectDrag = false;

        if (!_embeddedDragActive) { _embeddedDragActive = false; return; }

        bool   wasActive = _embeddedDragActive;
        var    model     = _embeddedDragModel;
        var    oldBlock  = _embeddedDragBlock;
        _embeddedDragActive = false;
        _embeddedDragModel  = null;
        _embeddedDragBlock  = null;

        if (!wasActive || model is null || oldBlock is null) return;

        // InFrontOfText / BehindText 는 Canvas 드래그가 처리. AsText / Inline 에서 WrapLeft/WrapRight
        // 로의 전환 또는 WrapLeft ↔ WrapRight 전환을 드롭 X 위치로 결정한다.
        var currentMode = model.WrapMode;
        if (currentMode is PolyDonky.Core.ImageWrapMode.InFrontOfText
                        or PolyDonky.Core.ImageWrapMode.BehindText
                        or PolyDonky.Core.ImageWrapMode.AsText)
            return;

        var    pt      = e.GetPosition(BodyEditor);
        double editorW = BodyEditor.ActualWidth;
        double third   = editorW / 3.0;

        if (currentMode == PolyDonky.Core.ImageWrapMode.Inline)
        {
            // Inline: 가로 위치에 따라 HAlign 만 바꾼다 (WrapMode 유지).
            model.HAlign = pt.X < third          ? PolyDonky.Core.ImageHAlign.Left
                         : pt.X > third * 2      ? PolyDonky.Core.ImageHAlign.Right
                         :                         PolyDonky.Core.ImageHAlign.Center;
        }
        else
        {
            // WrapLeft / WrapRight: 왼쪽 1/3 → WrapLeft, 중앙 1/3 → Inline 가운데, 오른쪽 1/3 → WrapRight.
            if (pt.X < third)
            {
                model.WrapMode = PolyDonky.Core.ImageWrapMode.WrapLeft;
                model.HAlign   = PolyDonky.Core.ImageHAlign.Left;
            }
            else if (pt.X > third * 2)
            {
                model.WrapMode = PolyDonky.Core.ImageWrapMode.WrapRight;
                model.HAlign   = PolyDonky.Core.ImageHAlign.Right;
            }
            else
            {
                model.WrapMode = PolyDonky.Core.ImageWrapMode.Inline;
                model.HAlign   = PolyDonky.Core.ImageHAlign.Center;
            }
        }

        var newBlock = Services.FlowDocumentBuilder.BuildImage(model);
        var doc      = BodyEditor.Document;
        doc.Blocks.InsertBefore(oldBlock, newBlock);
        doc.Blocks.Remove(oldBlock);
        RebuildOverlayImages();
        _viewModel?.MarkDirty();
    }

    // ── 도형 드래그 생성 / 오버레이 상태 ──────────────────────────
    private bool _drawingShape_active;
    private ShapeKind _drawingShape_kind = ShapeKind.Rectangle;

    // ── 도형 선택 상태 ────────────────────────────────────────────
    // 글상자 선택(_selectedOverlay)과 동일한 패턴 — 클릭으로 선택, ESC 로 해제,
    // Delete / Ctrl+C / Ctrl+X / Ctrl+V 로 통합 키보드 처리.
    private FrameworkElement? _selectedShapeCtrl;
    private PolyDonky.Core.ShapeObject? _selectedShape;

    // ── 폴리선/스플라인 클릭 입력 상태 ───────────────────────────
    // 각 클릭마다 점을 추가하고 더블클릭 또는 우클릭 메뉴로 마침.
    private bool               _drawingPolyline_active;
    private ShapeKind          _drawingPolyline_kind = ShapeKind.Polyline;
    private List<Point>        _drawingPolyline_points = new();
    private WpfShapes.Polyline? _polylinePreview;        // DrawPreviewCanvas 위의 고무줄 미리보기

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
            vm.InsertShapeRequested   += OnInsertShapeRequested;
            vm.RefreshSystemKeys();
            vm.RefreshMemoryUsage();
        }

        // RichTextBox 클릭 = 본문 편집 의도. 드래그 생성 모드가 아니면 글상자 선택 해제.
        // 동시에 임베드 객체(이미지·이모지) 위에서 드래그를 시작하는지 추적한다.
        BodyEditor.PreviewMouseLeftButtonDown += OnEditorPreviewMouseDownTrackDrag;
        BodyEditor.PreviewMouseMove           += OnEditorPreviewMouseMoveBlockDrag;
        BodyEditor.PreviewMouseLeftButtonUp   += OnEditorPreviewMouseUpEmbedded;

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
                if (_drawingTextBox || _drawingShape_active || _drawingPolyline_active)
                {
                    EndDrawingMode();
                    e.Handled = true;
                }
                else if (_selectedShape is not null)
                {
                    DeselectShape();
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

            case Key.Delete:
                // 선택된 객체(도형/글상자)가 있을 때만 가로채기. 본문 텍스트 삭제는 양보.
                if (TryDeleteSelectedObject()) e.Handled = true;
                break;

            case Key.Return:
            {
                int need = (_drawingPolyline_active &&
                            _drawingPolyline_kind is ShapeKind.Polygon or ShapeKind.ClosedSpline) ? 3 : 2;
                if (_drawingPolyline_active && _drawingPolyline_points.Count >= need)
                {
                    FinishPolylineShape();
                    e.Handled = true;
                }
                break;
            }

            // 부유 객체(도형·글상자) 자체의 복사/잘라내기/붙여넣기.
            // 객체 종류 무관하게 동작 — 새 객체 추가 시 TryDelete/TryCopy/TryPaste*Object 한 곳만 손보면 됨.
            // 안쪽 본문(InnerEditor)이 포커스 중이거나 본문 텍스트 selection 이 있으면 가로채지 않고
            // RichTextBox 의 일반 텍스트 클립보드 동작으로 넘긴다.
            case Key.C when (Keyboard.Modifiers & ModifierKeys.Control) != 0:
                if (TryCopySelectedObject()) e.Handled = true;
                break;
            case Key.X when (Keyboard.Modifiers & ModifierKeys.Control) != 0:
                if (TryCutSelectedObject()) e.Handled = true;
                break;
            case Key.V when (Keyboard.Modifiers & ModifierKeys.Control) != 0:
                if (TryPasteSelectedObject()) e.Handled = true;
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

        // 도형 오버레이 재구축
        RebuildOverlayShapes();
    }

    /// <summary>
    /// Document 모델을 순회해 InFrontOfText / BehindText 모드 ImageBlock 들을
    /// OverlayImageCanvas / UnderlayImageCanvas 에 절대 위치로 배치한다.
    /// </summary>
    private void RebuildOverlayImages()
    {
        OverlayImageCanvas.Children.Clear();
        UnderlayImageCanvas.Children.Clear();

        // 모델(_viewModel.Document)은 저장 시에만 FlowDocument 로부터 재구축되므로
        // 편집 중에는 반드시 live FlowDocument(BodyEditor.Document) 를 직접 순회해야 한다.
        foreach (var block in BodyEditor.Document.Blocks)
        {
            if (block.Tag is not PolyDonky.Core.ImageBlock img) continue;
            if (img.WrapMode is not (PolyDonky.Core.ImageWrapMode.InFrontOfText
                                  or PolyDonky.Core.ImageWrapMode.BehindText)) continue;

            var ctrl = Services.FlowDocumentBuilder.BuildOverlayImageControl(img);
            if (ctrl is null) continue;

            ctrl.Tag = img;
            System.Windows.Controls.Canvas.SetLeft(ctrl, Services.FlowDocumentBuilder.MmToDip(img.OverlayXMm));
            System.Windows.Controls.Canvas.SetTop(ctrl,  Services.FlowDocumentBuilder.MmToDip(img.OverlayYMm));
            ctrl.Cursor = System.Windows.Input.Cursors.SizeAll;
            // 우클릭은 PaperBorder.PreviewMouseRightButtonDown 통합 핸들러가 처리 — 개별 ContextMenu 불필요.

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

        System.Windows.Controls.Canvas.SetLeft(fe, _overlayDragStartLeft + dx);
        System.Windows.Controls.Canvas.SetTop (fe, _overlayDragStartTop  + dy);
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
            img.OverlayXMm = Services.FlowDocumentBuilder.DipToMm(left);
            img.OverlayYMm = Services.FlowDocumentBuilder.DipToMm(top);
            _viewModel?.MarkDirty();
        }
        e.Handled = true;
    }

    private void OpenOverlayImageProperties(PolyDonky.Core.ImageBlock img)
    {
        var dlg = new ImagePropertiesWindow(img) { Owner = this };
        if (dlg.ShowDialog() != true) return;

        // WrapMode 가 InFrontOfText/BehindText 이외로 바뀔 수 있으므로 플레이스홀더를
        // 새 Block 으로 교체한다. _viewModel.Document(저장 시 재구축)가 아니라
        // live FlowDocument 를 직접 수정해야 편집 중인 이미지가 누락되지 않는다.
        var placeholder = BodyEditor.Document.Blocks
            .FirstOrDefault(b => ReferenceEquals(b.Tag, img));
        if (placeholder is not null)
        {
            var newBlock = Services.FlowDocumentBuilder.BuildImage(img);
            BodyEditor.Document.Blocks.InsertBefore(placeholder, newBlock);
            BodyEditor.Document.Blocks.Remove(placeholder);
        }

        RebuildOverlayImages();
        _viewModel?.MarkDirty();
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

        // FlowDocument.PageWidth 를 본문 폭(종이 폭 − 좌여백 − 우여백)으로 갱신.
        // Build() 에서 초기값을 설정하지만, 여백 변경 시에도 즉시 반영되어야
        // 우측 정렬 객체(WrapRight Floater 등)가 정확한 위치에 그려진다.
        BodyEditor.Document.PageWidth =
            PolyDonky.App.Services.FlowDocumentBuilder.ComputeContentWidthDip(page);
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
        // live FlowDocument 를 함께 전달해 재빌드 전 모델 동기화 — 편집 중 이미지·글상자 손실 방지.
        dlg.StyleApplied += (_, styleSet) => _viewModel?.ApplyOutlineStyles(styleSet, BodyEditor.Document);
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
            // PageSettings 는 FlowDocument 블록 내용과 무관하다 (종이 크기·여백·배경색만 바뀜).
            // RebuildFlowDocument() 로 전체 재구성하면 저장 이후 편집한 이미지·글상자를 잃으므로
            // 레이아웃 속성만 직접 갱신한다.
            ApplyPageSettings(dlg.ResultSettings);
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

    private void OnInsertTable(object sender, RoutedEventArgs e)
    {
        var dlg = new TableInsertDialog { Owner = this };
        if (dlg.ShowDialog() != true || dlg.ResultTable is null) return;
        InsertTableBlock(BodyEditor, dlg.ResultTable);
        _viewModel?.MarkDirty();
        BodyEditor.Focus();
    }

    private static void InsertTableBlock(RichTextBox editor, PolyDonky.Core.Table table)
    {
        var wpfTable = PolyDonky.App.Services.FlowDocumentBuilder.BuildTable(table);
        var flowDoc  = editor.Document;

        // 캐럿이 속한 최상위 Block 바로 뒤에 표를 삽입한다.
        System.Windows.Documents.Block? current = null;
        var pos = editor.CaretPosition;
        if (pos.Paragraph is { } para)
        {
            foreach (var b in flowDoc.Blocks)
                if (b == para) { current = para; break; }
        }
        if (current is null)
        {
            foreach (var b in flowDoc.Blocks)
            {
                if (b.ContentStart.CompareTo(pos) <= 0 && pos.CompareTo(b.ContentEnd) <= 0)
                { current = b; break; }
            }
        }

        if (current is not null)
            flowDoc.Blocks.InsertAfter(current, wpfTable);
        else
            flowDoc.Blocks.Add(wpfTable);

        try { editor.CaretPosition = wpfTable.ContentEnd; }
        catch { /* 포지션 이동 실패는 무시 */ }
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

    // BodyEditor 내 컨텐츠(표·인라인 이미지/이모지) 우클릭 전용.
    // 오버레이/언더레이 객체는 OnPaperPreviewMouseRightButtonDown 에서 이미 처리했으므로
    // 여기에 도달하면 반드시 BodyEditor 본문 영역 우클릭이다.
    private void OnEmbeddedObjectContextMenuOpening(object sender, System.Windows.Controls.ContextMenuEventArgs e)
    {
        var pt   = System.Windows.Input.Mouse.GetPosition(BodyEditor);
        var menu = BodyEditor.ContextMenu;
        if (menu is null) return;

        // 표 셀 — 셀 속성 + 표 속성 항목 추가
        if (FindTableCellAtCaret(out var wpfTable, out var wpfRow, out var wpfCell) &&
            wpfTable is not null && wpfCell is not null &&
            wpfTable.Tag is PolyDonky.Core.Table coreTable)
        {
            AppendTablePropertyMenuItems(menu, wpfTable, wpfRow, wpfCell, coreTable);
        }

        // 인라인 이미지/이모지 — 속성 항목 추가 (표와 동시 존재 가능)
        if (FindEmbeddedObjectAt(e.OriginalSource, pt) is { } found)
        {
            AppendPropertyMenuItem(menu, () => OpenEmbeddedObjectProperties(found.img, found.container), "속성(_P)...");
        }
    }

    private bool FindTableCellAtCaret(
        out System.Windows.Documents.Table? wpfTable,
        out System.Windows.Documents.TableRow? wpfRow,
        out System.Windows.Documents.TableCell? wpfCell)
    {
        wpfTable = null; wpfRow = null; wpfCell = null;
        var para = BodyEditor.CaretPosition.Paragraph;
        if (para?.Parent is not System.Windows.Documents.TableCell tc) return false;
        wpfCell = tc;
        wpfRow  = tc.Parent as System.Windows.Documents.TableRow;
        var rowGroup = wpfRow?.Parent as System.Windows.Documents.TableRowGroup;
        wpfTable = rowGroup?.Parent as System.Windows.Documents.Table;
        return wpfTable is not null;
    }

    // ── 표 컨텍스트 메뉴 조립 ──────────────────────────────────────────────
    // 새 표 전용 기능을 추가할 때는 이 메서드 안에만 항목을 추가하면 된다.
    private void AppendTablePropertyMenuItems(
        System.Windows.Controls.ContextMenu menu,
        System.Windows.Documents.Table wpfTable,
        System.Windows.Documents.TableRow? wpfRow,
        System.Windows.Documents.TableCell wpfCell,
        PolyDonky.Core.Table coreTable)
    {
        var rowGroup = wpfTable.RowGroups.FirstOrDefault();
        int rowIdx  = wpfRow  is not null && rowGroup is not null ? rowGroup.Rows.IndexOf(wpfRow) : -1;
        int cellIdx = wpfRow  is not null ? wpfRow.Cells.IndexOf(wpfCell) : -1;
        PolyDonky.Core.TableCell? coreCell =
            rowIdx  >= 0 && rowIdx  < coreTable.Rows.Count &&
            cellIdx >= 0 && cellIdx < coreTable.Rows[rowIdx].Cells.Count
            ? coreTable.Rows[rowIdx].Cells[cellIdx]
            : null;

        bool hasNextCell  = wpfRow is not null && cellIdx + 1 < wpfRow.Cells.Count;
        bool hasNextRow   = rowIdx + 1 < coreTable.Rows.Count;
        bool canSplit     = (coreCell?.ColumnSpan ?? 1) > 1 || (coreCell?.RowSpan ?? 1) > 1;

        var items = new List<System.Windows.Controls.Control>();

        // ── 행 ─────────────────────────────────────────────────────────
        items.Add(new System.Windows.Controls.Separator());
        items.Add(MakeMenuItem("위에 행 삽입(_A)",   () => TableOp_InsertRow(wpfTable, coreTable, rowIdx, above: true)));
        items.Add(MakeMenuItem("아래에 행 삽입(_B)", () => TableOp_InsertRow(wpfTable, coreTable, rowIdx, above: false)));
        var delRow = MakeMenuItem("행 삭제(_R)", () => TableOp_DeleteRow(wpfTable, coreTable, rowIdx));
        delRow.IsEnabled = rowIdx >= 0;
        items.Add(delRow);

        // ── 열 ─────────────────────────────────────────────────────────
        items.Add(new System.Windows.Controls.Separator());
        items.Add(MakeMenuItem("왼쪽에 열 삽입(_L)",  () => TableOp_InsertColumn(wpfTable, coreTable, cellIdx, left: true)));
        items.Add(MakeMenuItem("오른쪽에 열 삽입(_G)", () => TableOp_InsertColumn(wpfTable, coreTable, cellIdx, left: false)));
        var delCol = MakeMenuItem("열 삭제(_C)", () => TableOp_DeleteColumn(wpfTable, coreTable, cellIdx));
        delCol.IsEnabled = cellIdx >= 0;
        items.Add(delCol);

        // ── 셀 병합·분할 ────────────────────────────────────────────────
        if (coreCell is not null && wpfRow is not null)
        {
            items.Add(new System.Windows.Controls.Separator());

            var mergeRight = MakeMenuItem("우측 셀과 병합(_M)",
                () => TableOp_MergeRight(wpfTable, coreTable, rowIdx, cellIdx, wpfRow, wpfCell, coreCell));
            mergeRight.IsEnabled = hasNextCell;
            items.Add(mergeRight);

            var mergeBelow = MakeMenuItem("아래 셀과 병합(_D)",
                () => TableOp_MergeBelow(wpfTable, coreTable, rowIdx, cellIdx, wpfRow, wpfCell, coreCell));
            mergeBelow.IsEnabled = hasNextRow;
            items.Add(mergeBelow);

            var split = MakeMenuItem("셀 분할(_S)",
                () => TableOp_SplitCell(wpfTable, coreTable, rowIdx, cellIdx, wpfRow, wpfCell, coreCell));
            split.IsEnabled = canSplit;
            items.Add(split);
        }

        // ── 속성 / 표 삭제 ──────────────────────────────────────────────
        items.Add(new System.Windows.Controls.Separator());
        if (coreCell is not null)
        {
            items.Add(MakeMenuItem("셀 속성(_P)...", () =>
            {
                var dlg = new CellPropertiesWindow(coreCell) { Owner = this };
                if (dlg.ShowDialog() == true)
                {
                    bool isHeader = rowIdx >= 0 && coreTable.Rows[rowIdx].IsHeader;
                    PolyDonky.App.Services.FlowDocumentBuilder.ApplyCellPropertiesToWpf(
                        wpfCell, coreCell, isHeader);
                    _viewModel?.MarkDirty();
                }
            }));
        }
        items.Add(MakeMenuItem("표 속성(_T)...", () =>
        {
            var dlg = new TablePropertiesWindow(coreTable) { Owner = this };
            if (dlg.ShowDialog() == true) _viewModel?.MarkDirty();
        }));
        items.Add(MakeMenuItem("표 삭제(_X)", () => TableOp_DeleteTable(wpfTable)));

        // 메뉴에 추가 후 닫힘 시 자동 제거
        foreach (var item in items) menu.Items.Add(item);
        void Cleanup(object? s, System.Windows.RoutedEventArgs _ev)
        {
            menu.Closed -= Cleanup;
            foreach (var ctrl in items) menu.Items.Remove(ctrl);
        }
        menu.Closed += Cleanup;
    }

    private static System.Windows.Controls.MenuItem MakeMenuItem(string header, Action onClick)
    {
        var item = new System.Windows.Controls.MenuItem { Header = header };
        item.Click += (_, _) => onClick();
        return item;
    }

    /// <summary>BodyEditor 컨텍스트 메뉴에 구분선 + 항목 하나를 추가하고 닫힘 시 제거.</summary>
    private static void AppendPropertyMenuItem(
        System.Windows.Controls.ContextMenu menu, Action onClick, string header)
    {
        var sep  = new System.Windows.Controls.Separator();
        var item = new System.Windows.Controls.MenuItem { Header = header };
        item.Click += (_, _) => onClick();
        menu.Items.Add(sep);
        menu.Items.Add(item);
        void Cleanup(object? s, System.Windows.RoutedEventArgs _ev)
        {
            menu.Closed -= Cleanup;
            menu.Items.Remove(sep);
            menu.Items.Remove(item);
        }
        menu.Closed += Cleanup;
    }

    // ── 표 구조 편집 연산 ────────────────────────────────────────────────
    // 모델(coreTable)과 WPF 표(wpfTable)를 항상 동시에 갱신해 동기화를 유지한다.

    private void TableOp_InsertRow(
        System.Windows.Documents.Table wpfTable, PolyDonky.Core.Table coreTable,
        int rowIdx, bool above)
    {
        var rowGroup = wpfTable.RowGroups.FirstOrDefault();
        if (rowGroup is null) return;

        int colCount = coreTable.Columns.Count > 0 ? coreTable.Columns.Count
            : (coreTable.Rows.Count > 0 ? coreTable.Rows.Max(r => r.Cells.Count) : 1);
        int insertAt = above ? Math.Max(rowIdx, 0) : Math.Min(rowIdx + 1, coreTable.Rows.Count);

        var newCoreRow = new PolyDonky.Core.TableRow();
        for (int c = 0; c < colCount; c++)
            newCoreRow.Cells.Add(new PolyDonky.Core.TableCell { Blocks = { Paragraph.Of(string.Empty) } });
        coreTable.Rows.Insert(insertAt, newCoreRow);

        var newWpfRow = new System.Windows.Documents.TableRow();
        for (int c = 0; c < colCount; c++)
            newWpfRow.Cells.Add(MakeEmptyWpfCell());
        rowGroup.Rows.Insert(insertAt, newWpfRow);

        _viewModel?.MarkDirty();
    }

    private void TableOp_DeleteRow(
        System.Windows.Documents.Table wpfTable, PolyDonky.Core.Table coreTable, int rowIdx)
    {
        if (rowIdx < 0 || rowIdx >= coreTable.Rows.Count) return;
        if (coreTable.Rows.Count <= 1)
        {
            MessageBox.Show(this, "마지막 행은 삭제할 수 없습니다.", "행 삭제",
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }
        coreTable.Rows.RemoveAt(rowIdx);
        var rowGroup = wpfTable.RowGroups.FirstOrDefault();
        if (rowGroup is not null && rowIdx < rowGroup.Rows.Count)
            rowGroup.Rows.RemoveAt(rowIdx);
        _viewModel?.MarkDirty();
    }

    private void TableOp_InsertColumn(
        System.Windows.Documents.Table wpfTable, PolyDonky.Core.Table coreTable,
        int cellIdx, bool left)
    {
        var rowGroup = wpfTable.RowGroups.FirstOrDefault();
        if (rowGroup is null) return;

        int insertAt    = left ? Math.Max(cellIdx, 0) : cellIdx + 1;
        double colWidth = coreTable.Columns.Count > 0 ? coreTable.Columns.Average(c => c.WidthMm) : 0;

        coreTable.Columns.Insert(
            Math.Min(insertAt, coreTable.Columns.Count),
            new PolyDonky.Core.TableColumn { WidthMm = colWidth });
        wpfTable.Columns.Insert(
            Math.Min(insertAt, wpfTable.Columns.Count),
            new System.Windows.Documents.TableColumn { Width = GridLength.Auto });

        for (int r = 0; r < coreTable.Rows.Count; r++)
        {
            var coreRow = coreTable.Rows[r];
            coreRow.Cells.Insert(
                Math.Min(insertAt, coreRow.Cells.Count),
                new PolyDonky.Core.TableCell { Blocks = { Paragraph.Of(string.Empty) } });

            if (r < rowGroup.Rows.Count)
            {
                var wpfRow = rowGroup.Rows[r];
                wpfRow.Cells.Insert(Math.Min(insertAt, wpfRow.Cells.Count), MakeEmptyWpfCell());
            }
        }
        _viewModel?.MarkDirty();
    }

    private void TableOp_DeleteColumn(
        System.Windows.Documents.Table wpfTable, PolyDonky.Core.Table coreTable, int cellIdx)
    {
        if (cellIdx < 0 || coreTable.Columns.Count <= 1)
        {
            if (coreTable.Columns.Count <= 1)
                MessageBox.Show(this, "마지막 열은 삭제할 수 없습니다.", "열 삭제",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }
        var rowGroup = wpfTable.RowGroups.FirstOrDefault();

        if (cellIdx < coreTable.Columns.Count) coreTable.Columns.RemoveAt(cellIdx);
        if (cellIdx < wpfTable.Columns.Count)  wpfTable.Columns.RemoveAt(cellIdx);

        for (int r = 0; r < coreTable.Rows.Count; r++)
        {
            var coreRow = coreTable.Rows[r];
            if (cellIdx < coreRow.Cells.Count) coreRow.Cells.RemoveAt(cellIdx);

            if (rowGroup is not null && r < rowGroup.Rows.Count)
            {
                var wpfRow = rowGroup.Rows[r];
                if (cellIdx < wpfRow.Cells.Count) wpfRow.Cells.RemoveAt(cellIdx);
            }
        }
        _viewModel?.MarkDirty();
    }

    private void TableOp_MergeRight(
        System.Windows.Documents.Table wpfTable, PolyDonky.Core.Table coreTable,
        int rowIdx, int cellIdx,
        System.Windows.Documents.TableRow wpfRow, System.Windows.Documents.TableCell wpfCell,
        PolyDonky.Core.TableCell coreCell)
    {
        if (cellIdx + 1 >= wpfRow.Cells.Count) return;

        coreCell.ColumnSpan = Math.Max(coreCell.ColumnSpan, 1) + 1;
        wpfCell.ColumnSpan  = coreCell.ColumnSpan;

        int nextIdx = cellIdx + 1;
        if (nextIdx < wpfRow.Cells.Count) wpfRow.Cells.RemoveAt(nextIdx);
        if (rowIdx < coreTable.Rows.Count && nextIdx < coreTable.Rows[rowIdx].Cells.Count)
            coreTable.Rows[rowIdx].Cells.RemoveAt(nextIdx);

        _viewModel?.MarkDirty();
    }

    private void TableOp_MergeBelow(
        System.Windows.Documents.Table wpfTable, PolyDonky.Core.Table coreTable,
        int rowIdx, int cellIdx,
        System.Windows.Documents.TableRow wpfRow, System.Windows.Documents.TableCell wpfCell,
        PolyDonky.Core.TableCell coreCell)
    {
        var rowGroup  = wpfTable.RowGroups.FirstOrDefault();
        int nextRowIdx = rowIdx + 1;
        if (nextRowIdx >= coreTable.Rows.Count) return;

        coreCell.RowSpan = Math.Max(coreCell.RowSpan, 1) + 1;
        wpfCell.RowSpan  = coreCell.RowSpan;

        var nextCoreRow = coreTable.Rows[nextRowIdx];
        if (cellIdx < nextCoreRow.Cells.Count) nextCoreRow.Cells.RemoveAt(cellIdx);

        if (rowGroup is not null && nextRowIdx < rowGroup.Rows.Count)
        {
            var nextWpfRow = rowGroup.Rows[nextRowIdx];
            if (cellIdx < nextWpfRow.Cells.Count) nextWpfRow.Cells.RemoveAt(cellIdx);
        }
        _viewModel?.MarkDirty();
    }

    private void TableOp_SplitCell(
        System.Windows.Documents.Table wpfTable, PolyDonky.Core.Table coreTable,
        int rowIdx, int cellIdx,
        System.Windows.Documents.TableRow wpfRow, System.Windows.Documents.TableCell wpfCell,
        PolyDonky.Core.TableCell coreCell)
    {
        var rowGroup = wpfTable.RowGroups.FirstOrDefault();

        if (coreCell.ColumnSpan > 1)
        {
            coreCell.ColumnSpan--;
            wpfCell.ColumnSpan = coreCell.ColumnSpan;

            var newCoreCell = new PolyDonky.Core.TableCell { Blocks = { Paragraph.Of(string.Empty) } };
            if (rowIdx < coreTable.Rows.Count)
                coreTable.Rows[rowIdx].Cells.Insert(cellIdx + 1, newCoreCell);
            wpfRow.Cells.Insert(cellIdx + 1, MakeEmptyWpfCell());
        }
        else if (coreCell.RowSpan > 1)
        {
            coreCell.RowSpan--;
            wpfCell.RowSpan = coreCell.RowSpan;

            int nextRowIdx = rowIdx + 1;
            if (nextRowIdx < coreTable.Rows.Count)
            {
                var newCoreCell = new PolyDonky.Core.TableCell { Blocks = { Paragraph.Of(string.Empty) } };
                coreTable.Rows[nextRowIdx].Cells.Insert(
                    Math.Min(cellIdx, coreTable.Rows[nextRowIdx].Cells.Count), newCoreCell);

                if (rowGroup is not null && nextRowIdx < rowGroup.Rows.Count)
                    rowGroup.Rows[nextRowIdx].Cells.Insert(
                        Math.Min(cellIdx, rowGroup.Rows[nextRowIdx].Cells.Count), MakeEmptyWpfCell());
            }
        }
        _viewModel?.MarkDirty();
    }

    private void TableOp_DeleteTable(System.Windows.Documents.Table wpfTable)
    {
        var block = BodyEditor.Document.Blocks
            .FirstOrDefault(b => ReferenceEquals(b, wpfTable));
        if (block is not null) BodyEditor.Document.Blocks.Remove(block);
        _viewModel?.MarkDirty();
    }

    private static System.Windows.Documents.TableCell MakeEmptyWpfCell()
    {
        double borderDip = PolyDonky.App.Services.FlowDocumentBuilder.PtToDip(0.75);
        double padLR     = PolyDonky.App.Services.FlowDocumentBuilder.MmToDip(1.5);
        double padTB     = PolyDonky.App.Services.FlowDocumentBuilder.MmToDip(1.0);
        var wcell = new System.Windows.Documents.TableCell
        {
            BorderBrush     = new WpfMedia.SolidColorBrush(WpfMedia.Color.FromRgb(0xC8, 0xC8, 0xC8)),
            BorderThickness = new Thickness(borderDip),
            Padding         = new Thickness(padLR, padTB, padLR, padTB),
        };
        wcell.Blocks.Add(new System.Windows.Documents.Paragraph(
            new System.Windows.Documents.Run(string.Empty)));
        return wcell;
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
            // 모드 전환 전 화면 위치 캡처 — overlay 모드(InFrontOfText/BehindText) 로 전환 시
            // 페이지 좌상단(0,0) 으로 점프하지 않고 현재 위치를 그대로 유지하도록 한다.
            // PaperBorder 기준 좌표 = OverlayImageCanvas/UnderlayImageCanvas 좌표.
            var prevMode = imageBlock.WrapMode;
            Point currentPos;
            try { currentPos = imgControl.TransformToVisual(PaperBorder).Transform(new Point(0, 0)); }
            catch { currentPos = new Point(0, 0); }

            var dlg = new ImagePropertiesWindow(imageBlock) { Owner = this };
            if (dlg.ShowDialog() == true)
            {
                // 모드 전환 (non-overlay → overlay) 이고 좌표가 기본값(0,0)이면 캡처한 위치 적용.
                bool toOverlay =
                    prevMode is not (PolyDonky.Core.ImageWrapMode.InFrontOfText
                                  or PolyDonky.Core.ImageWrapMode.BehindText)
                    && imageBlock.WrapMode is (PolyDonky.Core.ImageWrapMode.InFrontOfText
                                            or PolyDonky.Core.ImageWrapMode.BehindText);
                if (toOverlay && imageBlock.OverlayXMm == 0 && imageBlock.OverlayYMm == 0)
                {
                    imageBlock.OverlayXMm = Services.FlowDocumentBuilder.DipToMm(currentPos.X);
                    imageBlock.OverlayYMm = Services.FlowDocumentBuilder.DipToMm(currentPos.Y);
                }

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

    // 편집 > 지우기: 선택된 부유 객체(도형/글상자)가 있으면 그것을 삭제, 아니면 본문 텍스트 삭제.
    // RichTextBox 는 ApplicationCommands.Delete 를 자체 바인딩하지 않으므로 본문 분기에서는
    // 선택 영역을 지운다. selection 이 비어 있으면 캐럿 직후 한 글자(EditingCommands.Delete).
    private void OnEditDelete(object sender, RoutedEventArgs e)
    {
        if (TryDeleteSelectedObject()) return;
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

    private void OnInsertShapeRequested(object? sender, PolyDonky.Core.ShapeKind kind)
    {
        if (kind is ShapeKind.Line or ShapeKind.Polyline or ShapeKind.Spline
                 or ShapeKind.Polygon or ShapeKind.ClosedSpline)
        {
            _drawingPolyline_active = true;
            _drawingPolyline_kind   = kind;
            _drawingPolyline_points.Clear();
            Mouse.OverrideCursor = Cursors.Cross;
            if (_viewModel is not null)
                _viewModel.StatusMessage = SR.StatusDrawPolyline;
        }
        else
        {
            _drawingShape_active = true;
            _drawingShape_kind   = kind;
            Mouse.OverrideCursor = Cursors.Cross;
            if (_viewModel is not null)
                _viewModel.StatusMessage = SR.StatusDrawShape;
        }
    }

    private void EndDrawingMode()
    {
        _drawingTextBox       = false;
        _drawingShape_active  = false;
        _drawingInProgress    = false;
        _drawingPolyline_active = false;
        _drawingPolyline_points.Clear();
        ClearPolylinePreview();
        Mouse.OverrideCursor = null;
        DrawPreviewRect.Visibility = Visibility.Collapsed;
        if (PaperBorder.IsMouseCaptured) PaperBorder.ReleaseMouseCapture();
        if (_viewModel is not null) _viewModel.StatusMessage = SR.StatusReady;
    }

    private void ClearPolylinePreview()
    {
        if (_polylinePreview is not null)
        {
            DrawPreviewCanvas.Children.Remove(_polylinePreview);
            _polylinePreview = null;
        }
    }

    private void UpdatePolylinePreview(Point mousePos)
    {
        // DrawPreviewCanvas 위에 커밋된 점들 + 마우스 위치를 잇는 Polyline 을 그린다.
        if (_polylinePreview is null)
        {
            _polylinePreview = new WpfShapes.Polyline
            {
                Stroke = WpfMedia.Brushes.SteelBlue,
                StrokeThickness = 1.5,
                StrokeDashArray = new DoubleCollection { 4, 2 },
                IsHitTestVisible = false,
            };
            DrawPreviewCanvas.Children.Add(_polylinePreview);
        }

        _polylinePreview.Points.Clear();
        foreach (var pt in _drawingPolyline_points)
            _polylinePreview.Points.Add(pt);
        _polylinePreview.Points.Add(mousePos);
    }

    private void OnPaperPreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        // ── 폴리선/스플라인 클릭 입력 모드 ──────────────────────────────
        if (_drawingPolyline_active)
        {
            var pos = e.GetPosition(PaperBorder);
            pos.X = Math.Clamp(pos.X, 0, PaperBorder.ActualWidth);
            pos.Y = Math.Clamp(pos.Y, 0, PaperBorder.ActualHeight);

            if (e.ClickCount >= 2)
            {
                // 더블클릭: ClickCount==1 단계에서 이미 그 위치에 점이 추가되었으므로
                // 그대로 마감하면 더블클릭 위치가 마지막 점이 된다.
                int need = _drawingPolyline_kind is ShapeKind.Polygon or ShapeKind.ClosedSpline ? 3 : 2;
                if (_drawingPolyline_points.Count >= need)
                    FinishPolylineShape();
                else
                    EndDrawingMode();
            }
            else
            {
                // 단일 클릭: 점 추가
                _drawingPolyline_points.Add(pos);
                UpdatePolylinePreview(pos);

                // 직선은 2점 도달 시 자동 마감 (사용자: 시작점 클릭 → 끝점 클릭 → 끝)
                if (_drawingPolyline_kind == ShapeKind.Line && _drawingPolyline_points.Count >= 2)
                    FinishPolylineShape();
            }

            e.Handled = true;
            return;
        }

        if (!_drawingTextBox && !_drawingShape_active) return;

        var startPos = e.GetPosition(PaperBorder);
        startPos.X = Math.Clamp(startPos.X, 0, PaperBorder.ActualWidth);
        startPos.Y = Math.Clamp(startPos.Y, 0, PaperBorder.ActualHeight);

        _drawStart = startPos;
        _drawingInProgress = true;

        Canvas.SetLeft(DrawPreviewRect, startPos.X);
        Canvas.SetTop(DrawPreviewRect, startPos.Y);
        DrawPreviewRect.Width = 0;
        DrawPreviewRect.Height = 0;
        DrawPreviewRect.Visibility = Visibility.Visible;

        PaperBorder.CaptureMouse();
        e.Handled = true;
    }

    private void OnPaperPreviewMouseMove(object sender, MouseEventArgs e)
    {
        if (_drawingPolyline_active && _drawingPolyline_points.Count > 0)
        {
            var pos = e.GetPosition(PaperBorder);
            pos.X = Math.Clamp(pos.X, 0, PaperBorder.ActualWidth);
            pos.Y = Math.Clamp(pos.Y, 0, PaperBorder.ActualHeight);
            UpdatePolylinePreview(pos);
            return;
        }

        if (!_drawingInProgress || (!_drawingTextBox && !_drawingShape_active)) return;

        var pos2 = e.GetPosition(PaperBorder);
        pos2.X = Math.Clamp(pos2.X, 0, PaperBorder.ActualWidth);
        pos2.Y = Math.Clamp(pos2.Y, 0, PaperBorder.ActualHeight);

        double x = Math.Min(_drawStart.X, pos2.X);
        double y = Math.Min(_drawStart.Y, pos2.Y);
        double w = Math.Abs(pos2.X - _drawStart.X);
        double h = Math.Abs(pos2.Y - _drawStart.Y);

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

        bool wasShape = _drawingShape_active;
        EndDrawingMode();

        // 너무 작은 드래그(클릭에 가까움)는 생성 안 함.
        const double minDip = 10;
        if (w < minDip || h < minDip)
        {
            e.Handled = true;
            return;
        }

        if (wasShape)
        {
            FinishShapeDraw(x, y, w, h);
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

    // ── 통합 우클릭 핸들러 ──────────────────────────────────────────────────
    // 커서 아래 객체 타입을 우선순위대로 감지하고 해당 컨텍스트 메뉴를 조립한다.
    // 새 객체 타입 추가 시 여기에 else-if 블록만 추가하면 된다.
    private void OnPaperPreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
    {
        // ① 폴리선/스플라인 입력 모드
        if (_drawingPolyline_active)
        {
            OpenPolylineInputMenu();
            e.Handled = true;
            return;
        }

        var pt = e.GetPosition(BodyEditor);

        // ② 오버레이 도형 (InFrontOfText 우선, BehindText 차선)
        var hitShape = (FindCanvasChildAt(OverlayShapeCanvas, pt)
                     ?? FindCanvasChildAt(UnderlayShapeCanvas, pt))?.Tag as PolyDonky.Core.ShapeObject;
        if (hitShape is not null)
        {
            OpenContextMenu(BuildShapeMenu(hitShape));
            e.Handled = true;
            return;
        }

        // ③ 오버레이 이미지 (InFrontOfText 우선, BehindText 차선)
        var hitImage = (FindCanvasChildAt(OverlayImageCanvas, pt)
                     ?? FindCanvasChildAt(UnderlayImageCanvas, pt))?.Tag as PolyDonky.Core.ImageBlock;
        if (hitImage is not null)
        {
            OpenContextMenu(BuildOverlayImageMenu(hitImage));
            e.Handled = true;
            return;
        }

        // ④ BodyEditor 내 컨텐츠 — ContextMenuOpening 이 처리 (e.Handled = false)
    }

    private void OpenPolylineInputMenu()
    {
        int finishMin = _drawingPolyline_kind is ShapeKind.Polygon or ShapeKind.ClosedSpline ? 3 : 2;

        var itemFinish = new MenuItem { Header = "완료(_F)", IsEnabled = _drawingPolyline_points.Count >= finishMin };
        itemFinish.Click += (_, _) => FinishPolylineShape();

        var itemUndo = new MenuItem { Header = "마지막 점 취소(_Z)", IsEnabled = _drawingPolyline_points.Count >= 1 };
        itemUndo.Click += (_, _) =>
        {
            if (_drawingPolyline_points.Count > 0)
            {
                _drawingPolyline_points.RemoveAt(_drawingPolyline_points.Count - 1);
                var curPos = Mouse.GetPosition(PaperBorder);
                if (_drawingPolyline_points.Count > 0) UpdatePolylinePreview(curPos);
                else                                   ClearPolylinePreview();
            }
        };

        var itemClose = new MenuItem { Header = "시작점에 닫기(_L)", IsEnabled = _drawingPolyline_points.Count >= 3 };
        itemClose.Click += (_, _) =>
        {
            if (_drawingPolyline_points.Count >= 3)
            {
                _drawingPolyline_points.Add(_drawingPolyline_points[0]);
                FinishPolylineShape();
            }
        };

        var itemCancel = new MenuItem { Header = "전체 취소(_C)" };
        itemCancel.Click += (_, _) => EndDrawingMode();

        var menu = new ContextMenu();
        menu.Items.Add(itemFinish);
        menu.Items.Add(itemUndo);
        menu.Items.Add(itemClose);
        menu.Items.Add(new Separator());
        menu.Items.Add(itemCancel);
        OpenContextMenu(menu);
    }

    private ContextMenu BuildShapeMenu(PolyDonky.Core.ShapeObject shape)
    {
        var menu  = new ContextMenu();
        var props = new MenuItem { Header = "도형 속성(_P)..." };
        props.Click += (_, _) => OpenOverlayShapeProperties(shape);
        var del = new MenuItem { Header = "삭제(_D)" };
        del.Click += (_, _) => DeleteOverlayShape(shape);
        menu.Items.Add(props);
        menu.Items.Add(del);
        return menu;
    }

    private ContextMenu BuildOverlayImageMenu(PolyDonky.Core.ImageBlock img)
    {
        var menu  = new ContextMenu();
        var props = new MenuItem { Header = "그림 속성(_P)..." };
        props.Click += (_, _) => OpenOverlayImageProperties(img);
        menu.Items.Add(props);
        return menu;
    }

    private void OpenContextMenu(ContextMenu menu)
    {
        menu.PlacementTarget = PaperBorder;
        menu.IsOpen = true;
    }

    /// <summary>Canvas 자식 중 PaperBorder/BodyEditor 기준 좌표 pt 아래 첫 번째 요소를 반환.</summary>
    private static FrameworkElement? FindCanvasChildAt(Canvas canvas, Point pt)
    {
        foreach (UIElement child in canvas.Children)
        {
            if (child is not FrameworkElement fe) continue;
            double left = Canvas.GetLeft(fe); if (double.IsNaN(left)) left = 0;
            double top  = Canvas.GetTop(fe);  if (double.IsNaN(top))  top  = 0;
            double w = fe.ActualWidth  > 0 ? fe.ActualWidth  : fe.Width;
            double h = fe.ActualHeight > 0 ? fe.ActualHeight : fe.Height;
            if (double.IsNaN(w) || w <= 0 || double.IsNaN(h) || h <= 0) continue;
            if (pt.X >= left && pt.X <= left + w && pt.Y >= top && pt.Y <= top + h)
                return fe;
        }
        return null;
    }

    private void FinishPolylineShape()
    {
        const double DipsPerMm = TextBoxOverlay.DipsPerMm;
        var pts = _drawingPolyline_points;
        // 닫힌 면(Polygon/ClosedSpline)은 최소 3점, 열린 선은 최소 2점 필요.
        int minPts = _drawingPolyline_kind is ShapeKind.Polygon or ShapeKind.ClosedSpline ? 3 : 2;
        if (pts.Count < minPts) { EndDrawingMode(); return; }

        // 바운딩박스 계산
        double minX = double.MaxValue, minY = double.MaxValue;
        double maxX = double.MinValue, maxY = double.MinValue;
        foreach (var pt in pts)
        {
            if (pt.X < minX) minX = pt.X;
            if (pt.Y < minY) minY = pt.Y;
            if (pt.X > maxX) maxX = pt.X;
            if (pt.Y > maxY) maxY = pt.Y;
        }
        double wDip = Math.Max(maxX - minX, 1.0);
        double hDip = Math.Max(maxY - minY, 1.0);

        var kind  = _drawingPolyline_kind;
        var shape = new PolyDonky.Core.ShapeObject
        {
            Kind              = kind,
            WrapMode          = PolyDonky.Core.ImageWrapMode.InFrontOfText,
            WidthMm           = wDip / DipsPerMm,
            HeightMm          = hDip / DipsPerMm,
            OverlayXMm        = minX / DipsPerMm,
            OverlayYMm        = minY / DipsPerMm,
            StrokeColor       = "#2C3E50",
            StrokeThicknessPt = 1.5,
            FillColor         = kind is PolyDonky.Core.ShapeKind.Polygon
                                     or PolyDonky.Core.ShapeKind.ClosedSpline
                                 ? "#7FB3D9" : null,
            FillOpacity       = 0.7,
            Status            = NodeStatus.Modified,
        };

        // 점들을 바운딩박스 기준 상대 좌표(mm) 로 저장
        foreach (var pt in pts)
        {
            shape.Points.Add(new PolyDonky.Core.ShapePoint
            {
                X = (pt.X - minX) / DipsPerMm,
                Y = (pt.Y - minY) / DipsPerMm,
            });
        }

        EndDrawingMode();
        InsertShapeBlock(shape);
    }

    private void FinishShapeDraw(double xDip, double yDip, double wDip, double hDip)
    {
        const double DipsPerMm = TextBoxOverlay.DipsPerMm;

        var kind = _drawingShape_kind;
        var shape = new PolyDonky.Core.ShapeObject
        {
            Kind              = kind,
            WrapMode          = PolyDonky.Core.ImageWrapMode.InFrontOfText,
            WidthMm           = wDip / DipsPerMm,
            HeightMm          = hDip / DipsPerMm,
            OverlayXMm        = xDip / DipsPerMm,
            OverlayYMm        = yDip / DipsPerMm,
            StrokeColor       = "#2C3E50",
            StrokeThicknessPt = 1.5,
            FillColor         = kind is PolyDonky.Core.ShapeKind.Line
                                     or PolyDonky.Core.ShapeKind.Polyline
                                     or PolyDonky.Core.ShapeKind.Spline
                                 ? null : "#7FB3D9",
            FillOpacity = 0.7,
            Status      = NodeStatus.Modified,
        };

        // Line 계열: 드래그 방향을 바운딩박스 기준 상대 좌표(mm)로 저장.
        // _drawStart 가 bounding-box 어느 모서리인지는 (_drawStart - topLeft) 로 결정된다.
        if (kind is PolyDonky.Core.ShapeKind.Line
                 or PolyDonky.Core.ShapeKind.Polyline
                 or PolyDonky.Core.ShapeKind.Spline)
        {
            double sxDip = _drawStart.X - xDip; // 0 or wDip
            double syDip = _drawStart.Y - yDip; // 0 or hDip
            double exDip = wDip - sxDip;
            double eyDip = hDip - syDip;
            shape.Points.Add(new PolyDonky.Core.ShapePoint { X = sxDip / DipsPerMm, Y = syDip / DipsPerMm });
            shape.Points.Add(new PolyDonky.Core.ShapePoint { X = exDip / DipsPerMm, Y = eyDip / DipsPerMm });
            shape.HeightMm = Math.Max(shape.HeightMm, 1.0);
        }

        InsertShapeBlock(shape);
    }

    private void InsertShapeBlock(PolyDonky.Core.ShapeObject shape)
    {
        // 현재 캐럿 위치에 빈 paragraph 처럼 삽입 (overlay 모드이므로 placeholder 만 들어감)
        var block = Services.FlowDocumentBuilder.BuildShape(shape);
        var doc   = BodyEditor.Document;

        // 캐럿이 있는 단락 뒤에 삽입. 캐럿이 없으면 맨 끝에 추가.
        var caretBlock = BodyEditor.CaretPosition?.Paragraph
                         ?? BodyEditor.CaretPosition?.GetAdjacentElement(
                             System.Windows.Documents.LogicalDirection.Backward) as System.Windows.Documents.Block;
        if (caretBlock is not null && doc.Blocks.Contains(caretBlock))
            doc.Blocks.InsertAfter(caretBlock, block);
        else
            doc.Blocks.Add(block);

        _viewModel?.AddShapeToCurrentSection(shape);
        RebuildOverlayShapes();
    }

    // ── 도형 오버레이 재구축 ──────────────────────────────────────────────
    private void RebuildOverlayShapes()
    {
        // 재구축 전 선택 해제 — Children.Clear() 직후 _selectedShapeCtrl 이 stale 참조가 됨.
        DeselectShape();
        OverlayShapeCanvas.Children.Clear();
        UnderlayShapeCanvas.Children.Clear();

        foreach (var block in BodyEditor.Document.Blocks)
        {
            if (block.Tag is not PolyDonky.Core.ShapeObject shape) continue;
            if (shape.WrapMode is not (PolyDonky.Core.ImageWrapMode.InFrontOfText
                                    or PolyDonky.Core.ImageWrapMode.BehindText)) continue;

            var ctrl = Services.FlowDocumentBuilder.BuildOverlayShapeControl(shape);
            ctrl.Tag = shape;
            Canvas.SetLeft(ctrl, Services.FlowDocumentBuilder.MmToDip(shape.OverlayXMm));
            Canvas.SetTop(ctrl,  Services.FlowDocumentBuilder.MmToDip(shape.OverlayYMm));
            ctrl.Cursor = Cursors.SizeAll;

            // 우클릭은 PaperBorder.PreviewMouseRightButtonDown 통합 핸들러가 처리 — 개별 ContextMenu 불필요.
            ctrl.MouseLeftButtonDown += OnOverlayShapeMouseDown;

            var canvas = shape.WrapMode == PolyDonky.Core.ImageWrapMode.BehindText
                ? UnderlayShapeCanvas
                : OverlayShapeCanvas;
            canvas.Children.Add(ctrl);
        }
    }

    private void DeleteOverlayShape(PolyDonky.Core.ShapeObject shape)
    {
        var placeholder = BodyEditor.Document.Blocks
            .FirstOrDefault(b => ReferenceEquals(b.Tag, shape));
        if (placeholder is not null)
            BodyEditor.Document.Blocks.Remove(placeholder);
        _viewModel?.RemoveShape(shape);
        RebuildOverlayShapes();
    }

    private void OpenOverlayShapeProperties(PolyDonky.Core.ShapeObject shape)
    {
        var dlg = new ShapePropertiesWindow(shape) { Owner = this };
        if (dlg.ShowDialog() != true) return;

        // WrapMode 변경으로 Block 타입이 달라질 수 있으므로 교체
        var placeholder = BodyEditor.Document.Blocks
            .FirstOrDefault(b => ReferenceEquals(b.Tag, shape));
        if (placeholder is not null)
        {
            var newBlock = Services.FlowDocumentBuilder.BuildShape(shape);
            BodyEditor.Document.Blocks.InsertBefore(placeholder, newBlock);
            BodyEditor.Document.Blocks.Remove(placeholder);
        }
        RebuildOverlayShapes();
        _viewModel?.MarkDirty();
    }

    // ── 오버레이 도형 드래그 이동 상태 ───────────────────────────────────
    private FrameworkElement? _draggingOverlayShape;
    private Point  _overlayShapeDragStart;
    private double _overlayShapeDragStartLeft;
    private double _overlayShapeDragStartTop;
    private bool   _overlayShapeDragMoved;

    private void OnOverlayShapeMouseDown(object sender, MouseButtonEventArgs e)
    {
        if (sender is not FrameworkElement fe) return;

        if (e.ClickCount == 2 && fe.Tag is PolyDonky.Core.ShapeObject dblShape)
        {
            OpenOverlayShapeProperties(dblShape);
            e.Handled = true;
            return;
        }

        // 단일 클릭 → 선택 + 드래그 시작
        if (fe.Tag is PolyDonky.Core.ShapeObject clickShape)
            SelectShape(fe, clickShape);

        if (fe.Parent is not Canvas canvas) return;
        _draggingOverlayShape       = fe;
        _overlayShapeDragStart      = e.GetPosition(canvas);
        _overlayShapeDragStartLeft  = Canvas.GetLeft(fe);
        _overlayShapeDragStartTop   = Canvas.GetTop(fe);
        if (double.IsNaN(_overlayShapeDragStartLeft)) _overlayShapeDragStartLeft = 0;
        if (double.IsNaN(_overlayShapeDragStartTop))  _overlayShapeDragStartTop  = 0;
        _overlayShapeDragMoved = false;
        fe.CaptureMouse();
        fe.MouseMove         += OnOverlayShapeDragMove;
        fe.MouseLeftButtonUp += OnOverlayShapeDragUp;
        e.Handled = true;
    }

    private void OnOverlayShapeDragMove(object sender, MouseEventArgs e)
    {
        if (sender is not FrameworkElement fe || !ReferenceEquals(_draggingOverlayShape, fe)) return;
        if (fe.Parent is not Canvas canvas) return;
        var pos = e.GetPosition(canvas);
        double dx = pos.X - _overlayShapeDragStart.X;
        double dy = pos.Y - _overlayShapeDragStart.Y;
        if (Math.Abs(dx) > 0.5 || Math.Abs(dy) > 0.5) _overlayShapeDragMoved = true;
        Canvas.SetLeft(fe, _overlayShapeDragStartLeft + dx);
        Canvas.SetTop (fe, _overlayShapeDragStartTop  + dy);
    }

    private void OnOverlayShapeDragUp(object sender, MouseButtonEventArgs e)
    {
        if (sender is not FrameworkElement fe || !ReferenceEquals(_draggingOverlayShape, fe)) return;
        fe.ReleaseMouseCapture();
        fe.MouseMove         -= OnOverlayShapeDragMove;
        fe.MouseLeftButtonUp -= OnOverlayShapeDragUp;
        _draggingOverlayShape = null;

        if (_overlayShapeDragMoved && fe.Tag is PolyDonky.Core.ShapeObject s)
        {
            double left = Canvas.GetLeft(fe);
            double top  = Canvas.GetTop(fe);
            if (double.IsNaN(left)) left = 0;
            if (double.IsNaN(top))  top  = 0;
            s.OverlayXMm = Services.FlowDocumentBuilder.DipToMm(left);
            s.OverlayYMm = Services.FlowDocumentBuilder.DipToMm(top);
            _viewModel?.MarkDirty();
        }
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
        // 다른 종류 선택 해제
        DeselectShape();
    }

    private void DeselectAllOverlays()
    {
        if (_selectedOverlay is not null)
        {
            _selectedOverlay.IsSelected = false;
            _selectedOverlay = null;
        }
        DeselectShape();
    }

    // ── 도형 선택 / 키보드 일반화 ─────────────────────────────────
    private void SelectShape(FrameworkElement ctrl, PolyDonky.Core.ShapeObject shape)
    {
        if (ReferenceEquals(_selectedShape, shape)) return;
        DeselectShape();
        // 다른 종류 선택 해제 (글상자)
        if (_selectedOverlay is not null)
        {
            _selectedOverlay.IsSelected = false;
            _selectedOverlay = null;
        }

        _selectedShape     = shape;
        _selectedShapeCtrl = ctrl;
        ctrl.Effect = new System.Windows.Media.Effects.DropShadowEffect
        {
            Color       = WpfMedia.Colors.DodgerBlue,
            ShadowDepth = 0,
            BlurRadius  = 14,
            Opacity     = 0.9,
        };
        // 윈도우로 키보드 포커스 — Delete/Ctrl+C 가 BodyEditor 가 아닌 윈도우에서 처리되도록.
        Focus();
        Keyboard.Focus(this);
    }

    private void DeselectShape()
    {
        if (_selectedShapeCtrl is not null)
            _selectedShapeCtrl.Effect = null;
        _selectedShape     = null;
        _selectedShapeCtrl = null;
    }

    /// <summary>선택된 객체(도형/글상자)를 종류 무관하게 삭제. 새 객체 추가 시 여기에 한 줄만 더하면 된다.</summary>
    private bool TryDeleteSelectedObject()
    {
        if (_selectedShape is { } shape)
        {
            DeleteOverlayShape(shape);
            DeselectShape();
            return true;
        }
        if (_selectedOverlay is { } overlay)
        {
            // 안쪽 본문에 텍스트 selection 이 있으면 일반 텍스트 삭제에 양보.
            if (overlay.InnerEditor.IsKeyboardFocusWithin && !overlay.InnerEditor.Selection.IsEmpty)
                return false;
            FloatingCanvas.Children.Remove(overlay);
            _viewModel?.RemoveFloatingObject(overlay.Model);
            _selectedOverlay = null;
            BodyEditor.Focus();
            return true;
        }
        return false;
    }

    /// <summary>선택된 객체를 종류 무관하게 클립보드로 복사.</summary>
    private bool TryCopySelectedObject()
    {
        if (_selectedShape is { } shape) return CopyShapeToClipboard(shape);
        return TryCopySelectedFloatingObject();
    }

    /// <summary>잘라내기 = 복사 후 삭제.</summary>
    private bool TryCutSelectedObject()
    {
        if (_selectedShape is { } shape)
        {
            if (!CopyShapeToClipboard(shape)) return false;
            DeleteOverlayShape(shape);
            DeselectShape();
            return true;
        }
        return TryCutSelectedFloatingObject();
    }

    /// <summary>붙여넣기는 클립보드 포맷에 따라 자동 분기 — 도형 → 글상자 순.</summary>
    private bool TryPasteSelectedObject()
    {
        if (Clipboard.ContainsData(BlockClipboardFormat))   return TryPasteBlockFromClipboard();
        if (Clipboard.ContainsData(FloatingObjectClipboardFormat)) return TryPasteFloatingObject();
        return false;
    }

    private const string BlockClipboardFormat = "PolyDonky.Block.v1";

    private bool CopyShapeToClipboard(PolyDonky.Core.ShapeObject shape)
    {
        var json = System.Text.Json.JsonSerializer.Serialize<PolyDonky.Core.Block>(
            shape, PolyDonky.Core.JsonDefaults.Options);
        var dataObj = new DataObject();
        dataObj.SetData(BlockClipboardFormat, json);
        if (!string.IsNullOrEmpty(shape.LabelText))
            dataObj.SetText(shape.LabelText);
        Clipboard.SetDataObject(dataObj, copy: true);
        return true;
    }

    private bool TryPasteBlockFromClipboard()
    {
        var json = Clipboard.GetData(BlockClipboardFormat) as string;
        if (string.IsNullOrEmpty(json)) return false;
        PolyDonky.Core.Block? block;
        try { block = System.Text.Json.JsonSerializer.Deserialize<PolyDonky.Core.Block>(
                  json, PolyDonky.Core.JsonDefaults.Options); }
        catch { return false; }

        if (block is PolyDonky.Core.ShapeObject shape)
        {
            shape.Id = null;
            shape.OverlayXMm += 5;
            shape.OverlayYMm += 5;
            shape.Status = NodeStatus.Modified;
            InsertShapeBlock(shape);
            return true;
        }
        return false;
    }

}

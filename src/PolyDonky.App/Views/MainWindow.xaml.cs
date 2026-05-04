using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
using PolyDonky.App.Pagination;
using PolyDonky.App.Services;
using PolyDonky.App.ViewModels;
using PolyDonky.Core;
using SR = PolyDonky.App.Properties.Resources;
using WpfDocs = System.Windows.Documents;
using WpfMedia = System.Windows.Media;
using WpfShapes = System.Windows.Shapes;

namespace PolyDonky.App.Views;

public partial class MainWindow : Window
{
    private MainViewModel? _viewModel;
    private bool _suppressTextChanged;
    private bool _suppressPasteCommand;
    private DispatcherTimer? _statusTimer;
    private DictionaryWindow? _dictWindow;

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

    // ── 마퀴(범위 드래그) 멀티-선택 상태 ────────────────────────────────────────
    // 페이지 전체 대상 — 텍스트 블록 + 오버레이(이미지/도형/표/글상자)를 한 번에 선택.
    // - 빈 여백에서 드래그: 사각형 내 모든 개체 선택
    // - Ctrl+클릭 오버레이: 해당 개체를 선택에 추가/제거
    private readonly List<FrameworkElement> _multiSelectedControls = new();
    private bool _marqueeSelecting;

    private void OnEditorPreviewMouseDownTrackDrag(object sender, MouseButtonEventArgs e)
    {
        _embeddedDragModel  = null;
        _embeddedDragBlock  = null;
        _embeddedDragActive = false;

        if (!_drawingTextBox) DeselectAllOverlays();
        var pt = e.GetPosition(BodyEditor);

        // 표 열 경계선 위: 열 너비 드래그 시작
        if (_tableColResizeHovering &&
            TryHitTableColumnBorder(pt, out var rcWpf, out var rcCore, out int rcIdx, out _))
        {
            _suppressEmbeddedObjectDrag = false;
            StartTableColumnResize(rcWpf!, rcCore!, rcIdx, pt.X);
            e.Handled = true;
            return;
        }

        // Alt + 클릭 → BehindText 그림(BodyEditor 뒤 UnderlayImageCanvas) 드래그 시작.
        // 일반 클릭은 본문 텍스트 선택을 위해 양보 — 텍스트 위에 그림이 깔린 영역에서도 편집 가능.
        if ((Keyboard.Modifiers & ModifierKeys.Alt) != 0 &&
            FindCanvasChildAt(UnderlayImageCanvas, pt) is { } underlayCtrl)
        {
            StartUnderlayImageDrag(underlayCtrl, e);
            e.Handled = true;
            return;
        }

        // (Ctrl+클릭 오버레이 토글·멀티-선택 해제는 PaperHost.PreviewMouseLeftButtonDown
        //  통합 핸들러가 처리 — 오버레이 컨트롤은 BodyEditor 의 형제라 여기로는 안 옴.)

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
            CommitOverlayDragPosition(img, left, top);
            _viewModel?.MarkDirty();
        }
        e.Handled = true;
    }

    private void OnEditorPreviewMouseMoveBlockDrag(object sender, MouseEventArgs e)
    {
        var pt = e.GetPosition(BodyEditor);

        // 우선순위 1: 표 열 너비 드래그 중
        if (_tableColResizeActive)
        {
            if (e.LeftButton != MouseButtonState.Pressed) { FinishTableColumnResize(); return; }
            double delta    = pt.X - _colRszStartX;
            double newLeft  = Math.Max(_colRszInitLeft  + delta, TableColResizeMinDip);
            double newRight = _colRszRightCol != null
                            ? Math.Max(_colRszInitRight - delta, TableColResizeMinDip)
                            : 0;
            _colRszLeftCol!.Width = new GridLength(newLeft);
            if (_colRszRightCol != null)
                _colRszRightCol.Width = new GridLength(newRight);
            e.Handled = true;
            return;
        }

        // 우선순위 2: 임베드 오브젝트(이미지) 드래그 억제
        if (_suppressEmbeddedObjectDrag)
        {
            if (e.LeftButton != MouseButtonState.Pressed)
            {
                _embeddedDragActive = false;
                _suppressEmbeddedObjectDrag = false;
                Mouse.OverrideCursor = null;
                return;
            }
            e.Handled = true;
            if (!_embeddedDragActive &&
                (Math.Abs(pt.X - _embeddedDragOrigin.X) > 8 ||
                 Math.Abs(pt.Y - _embeddedDragOrigin.Y) > 8))
            {
                _embeddedDragActive  = true;
                Mouse.OverrideCursor = Cursors.SizeAll;
            }
            return;
        }

        // 우선순위 3: 아무 드래그도 없을 때 — 표 열 경계선 커서 표시
        if (!_drawingShape_active && !_drawingPolyline_active && !_drawingTextBox &&
            e.LeftButton != MouseButtonState.Pressed)
        {
            bool onBorder = TryHitTableColumnBorder(pt, out _, out _, out _, out _);
            if (onBorder != _tableColResizeHovering)
            {
                _tableColResizeHovering = onBorder;
                Mouse.OverrideCursor    = (onBorder || _colDivHovering) ? Cursors.SizeWE : null;
            }
        }
    }

    private void OnEditorPreviewMouseUpEmbedded(object sender, MouseButtonEventArgs e)
    {
        // 표 열 너비 드래그 완료
        if (_tableColResizeActive)
        {
            FinishTableColumnResize();
            e.Handled = true;
            return;
        }

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
    // 직선(Line) 자동마감 직후 발생하는 ClickCount==2 이벤트를 억제하기 위한 플래그.
    // ClickCount==1 에서 2점 도달→자동마감 시 true 로 설정, 다음 PreviewMouseLeftButtonDown 에서 소비.
    private bool _suppressNextClickAfterLineFinish;

    // ── 페이지 구분선 ────────────────────────────────────────────────
    private double _pageHeightDip;  // 한 페이지 높이(DIP). 0이면 표시 안 함.

    // ── 조판부호 보기 ────────────────────────────────────────────────
    private bool _showTypesettingMarks;

    // ── BodyEditor 호환 심(shim) ─────────────────────────────────────
    // 활성 페이지 RTB 를 단일 편집기처럼 다루기 위한 프로퍼티.
    // Selection·CaretPosition·Document.Blocks 등 단일-RTB API 가 필요한 곳에서 사용한다.
    private RichTextBox BodyEditor
        => PageEditorHost.ActiveEditor ?? PageEditorHost.FirstEditor
           ?? throw new InvalidOperationException("페이지 에디터가 초기화되지 않았습니다.");

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
        // 페이지 텍스트 변경 이벤트 — SetupPages 가 각 RTB 에 연결한 뒤 여기로 집결.
        PageEditorHost.PageTextChanged += OnEditorTextChanged;
        // 상태 표시줄 Insert/CapsLock/NumLock 갱신을 위해 윈도우 레벨 키 입력 가로채기.
        PreviewKeyDown += OnPreviewKeyDown;

        // IME(한글 등) 조합 중에는 RTB 재구성을 미룬다 — 조합 도중 SetupPageEditors 가
        // RTB 를 새로 만들면 IME 조합 상태(ㅈ + ㅏ → 자)가 끊겨 한 글자가 둘로 갈라진다.
        AddHandler(System.Windows.Input.TextCompositionManager.PreviewTextInputStartEvent,
            new System.Windows.Input.TextCompositionEventHandler(OnImeCompositionStart),
            handledEventsToo: true);
        AddHandler(System.Windows.Input.TextCompositionManager.PreviewTextInputEvent,
            new System.Windows.Input.TextCompositionEventHandler(OnImeCompositionEnd),
            handledEventsToo: true);
    }

    /// <summary>가장 최근에 키보드 포커스를 가졌던 RichTextBox.</summary>
    private RichTextBox? _lastTextEditor;

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        if (DataContext is MainViewModel vm)
        {
            _viewModel = vm;
            vm.LiveDocumentProvider = () => PageEditorHost.PageCount > 0
                ? ParseAllPageEditors()
                : null;
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

        // 각 페이지 RTB 에 대한 이벤트 구독·속성 설정은 ConfigurePageRtb 콜백에서 수행.
        // SetupPageEditors → PageEditorHost.SetupPages 호출 시 RTB 마다 ConfigurePageRtb 가 적용된다.

        // OnLoaded 시점에는 WPF 첫 렌더링이 아직 완료되지 않아 오프스크린 RTB 의
        // GetCharacterRect 가 Y=0 을 반환할 수 있다. Background 우선순위로 지연해
        // 첫 렌더링 완료 후 올바른 측정값으로 재-페이지네이션한다.
        Dispatcher.BeginInvoke(DispatcherPriority.Background, () =>
        {
            if (_viewModel?.Document is not null && _pageGeometry is not null)
            {
                UpdatePaginatedDoc();
                SetupPageEditors();
                RebuildOverlays();
            }
        });

        _statusTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(1),
        };
        _statusTimer.Tick += OnStatusTimerTick;
        _statusTimer.Start();

        _showTypesettingMarks = LanguageService.ShowTypesettingMarks;
        MiTypesettingMarks.IsChecked = _showTypesettingMarks;
    }

    /// <summary>
    /// 붙여넣기 완료 직후 FlowDocument 를 순회해 Tag=null 인 Wpf.Table 에 Core.Table 을 부착한다.
    /// 이 시점에 EnsureCoreTable 을 호출해두면 우클릭 메뉴·열 리사이즈가 즉시 정상 동작한다.
    /// DataObject.AddPastingHandler 는 실제 삽입 전에 발화하므로, 삽입 완료 후 처리를 위해
    /// Dispatcher.BeginInvoke(DispatcherPriority.Background) 로 지연 실행한다.
    /// (PolyDonky.FlowSelection.v1 경로는 Tag 가 처음부터 살아오지만, 외부 앱 XAML 붙여넣기 등의
    ///  fallback 경로 안전망으로 유지.)
    /// </summary>
    private void OnBodyEditorPasting(object sender, DataObjectPastingEventArgs e)
    {
        Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Background, () =>
        {
            foreach (var block in PageEditorHost.AllBlocks)
            {
                if (block is System.Windows.Documents.Table wpfTable && wpfTable.Tag is null)
                    EnsureCoreTable(wpfTable);
                else if (block is System.Windows.Documents.Section sec)
                    foreach (var inner in sec.Blocks.OfType<System.Windows.Documents.Table>())
                        if (inner.Tag is null) EnsureCoreTable(inner);
            }
        });
    }

    // ── 본문 RichTextBox 멀티-블록 클립보드 (Tag 완전 보존) ─────────────────

    /// <summary>본문 RichTextBox 가 다루는 multi-block 선택의 Core JSON 직렬화 포맷.</summary>
    private const string FlowSelectionClipboardFormat = "PolyDonky.FlowSelection.v1";

    private void OnBodyEditorPreviewExecuted(object sender, System.Windows.Input.ExecutedRoutedEventArgs e)
    {
        // 부유 객체 전용 핸들러가 이미 처리한 경우(e.Handled=true)는 여기 도달하지 않음.
        if (e.Command == ApplicationCommands.Copy)
        {
            if (TryCopyFlowSelection(cut: false)) e.Handled = true;
        }
        else if (e.Command == ApplicationCommands.Cut)
        {
            if (TryCopyFlowSelection(cut: true)) e.Handled = true;
        }
        else if (e.Command == ApplicationCommands.Paste)
        {
            if (!_suppressPasteCommand && TryPasteFlowSelection()) e.Handled = true;
        }
    }

    /// <summary>본문 선택 영역을 Core.Block 리스트로 추출해 PolyDonky.FlowSelection.v1 포맷으로 클립보드에 저장.</summary>
    private bool TryCopyFlowSelection(bool cut)
    {
        // 단 교차 선택: 각 RTB 의 선택 텍스트를 순서대로 결합해 클립보드에 넣는다.
        if (_crossSelActive)
            return TryCopyCrossColumnSelection(cut);

        var sel = BodyEditor.Selection;
        if (sel.IsEmpty) return false;

        var coreBlocks = ExtractCoreSelection();
        if (coreBlocks.Count == 0) return false;

        // ID 충돌 방지를 위해 새 ID 발급 + Modified 표시
        foreach (var b in coreBlocks) ResetCoreBlockId(b);

        try
        {
            var json = System.Text.Json.JsonSerializer.Serialize(
                coreBlocks, PolyDonky.Core.JsonDefaults.Options);

            var dataObj = new System.Windows.DataObject();
            dataObj.SetData(FlowSelectionClipboardFormat, json);
            dataObj.SetText(sel.Text);  // 다른 앱에서 plain text 로 받을 수 있도록

            // XamlPackage / RTF 도 함께 — 외부 앱 호환용 (우리 앱에서는 우선 FlowSelection 사용).
            // 주의: SetData 의 타입은 포맷별 약속을 정확히 따라야 한다 — 어긋나면 받는 쪽이
            // byte[].ToString() = "System.Byte[]" 를 텍스트로 받아들일 수 있다.
            //   • XamlPackage → MemoryStream (binary). using 으로 dispose 하면 클립보드에서 못 읽으므로 그대로 둠.
            //   • RTF         → string (ASCII RTF 텍스트).
            try
            {
                var msX = new System.IO.MemoryStream();
                sel.Save(msX, System.Windows.DataFormats.XamlPackage);
                msX.Position = 0;
                dataObj.SetData(System.Windows.DataFormats.XamlPackage, msX);
            }
            catch { /* 일부 선택은 XamlPackage 직렬화 실패 — 무시 */ }

            try
            {
                using var msR = new System.IO.MemoryStream();
                sel.Save(msR, System.Windows.DataFormats.Rtf);
                var rtf = System.Text.Encoding.ASCII.GetString(msR.ToArray());
                dataObj.SetData(System.Windows.DataFormats.Rtf, rtf);
            }
            catch { /* RTF 직렬화 실패는 무시 */ }

            System.Windows.Clipboard.SetDataObject(dataObj, copy: true);
        }
        catch
        {
            return false;
        }

        if (cut)
        {
            sel.Text = string.Empty;
            _viewModel?.MarkDirty();
        }
        return true;
    }

    /// <summary>
    /// 단 교차 선택(여러 RTB)이 활성일 때의 복사/잘라내기.
    /// 각 RTB 의 선택 텍스트를 순서대로 결합해 클립보드에 넣는다.
    /// </summary>
    private bool TryCopyCrossColumnSelection(bool cut)
    {
        var selected = PageEditorHost.PageEditors
            .Where(r => !r.Selection.IsEmpty)
            .ToList();
        if (selected.Count == 0) return false;

        var text = string.Concat(selected.Select(r => r.Selection.Text));
        try { System.Windows.Clipboard.SetText(text); }
        catch { return false; }

        if (cut)
        {
            foreach (var r in selected)
                r.Selection.Text = string.Empty;
            ClearCrossColumnSelection();
            _viewModel?.MarkDirty();
        }
        return true;
    }

    /// <summary>
    /// PolyDonky.FlowSelection.v1 클립보드 데이터를 캐럿 위치에 붙여넣는다.</summary>
    private bool TryPasteFlowSelection()
    {
        if (!System.Windows.Clipboard.ContainsData(FlowSelectionClipboardFormat)) return false;
        if (System.Windows.Clipboard.GetData(FlowSelectionClipboardFormat) is not string json
            || string.IsNullOrEmpty(json)) return false;

        List<PolyDonky.Core.Block>? blocks;
        try
        {
            blocks = System.Text.Json.JsonSerializer.Deserialize<List<PolyDonky.Core.Block>>(
                json, PolyDonky.Core.JsonDefaults.Options);
        }
        catch { return false; }
        if (blocks is null || blocks.Count == 0) return false;

        foreach (var b in blocks) ResetCoreBlockId(b);

        if (!BodyEditor.Selection.IsEmpty) BodyEditor.Selection.Text = string.Empty;
        var caret = BodyEditor.CaretPosition;

        // 단일 단락 콘텐츠 → 캐럿 위치에 인라인 삽입 (새 단락 블록 생성 않음)
        if (blocks.Count == 1 && blocks[0] is PolyDonky.Core.Paragraph singlePara)
        {
            var newCaret = InsertParagraphInline(singlePara, caret);
            if (newCaret != null) try { BodyEditor.CaretPosition = newCaret; } catch { }
            _viewModel?.MarkDirty();
            return true;
        }

        // 셀 안에 단락만 붙여넣을 때 — 셀의 Blocks 컬렉션에 직접 단락을 삽입한다.
        // (multi-block 일반 경로로 가면 enclosing Table 뒤에 붙어 셀 밖으로 나간다)
        if (IsCaretInTableCell(caret) && blocks.All(b => b is PolyDonky.Core.Paragraph))
        {
            InsertParagraphsIntoCell(blocks.Cast<PolyDonky.Core.Paragraph>().ToList(), caret);
            _viewModel?.MarkDirty();
            return true;
        }

        // 복수 블록 또는 표·이미지·도형 — 앵커 뒤에 새 Block 으로 삽입
        var doc = BodyEditor.Document;
        System.Windows.Documents.Block? anchor = FindTopLevelAnchorForCaret(doc, caret);

        // 셀 안에 Table 을 붙여넣으면 중첩 표로 레이아웃이 무너진다.
        // 이 경우에만 InsertBefore(table.NextBlock) 으로 enclosing Table 뒤에 삽입하고,
        // 그 외 (텍스트·도형 등) 는 InsertAfter(anchor) 를 유지해 셀 안에 정상 삽입되도록 한다.
        bool avoidTableInTable = anchor is System.Windows.Documents.Table
            && IsCaretInTableCell(caret)
            && blocks.Any(b => b is PolyDonky.Core.Table);

        System.Windows.Documents.Block? insertBefore = avoidTableInTable ? anchor!.NextBlock : null;

        bool hasOverlay = false;
        foreach (var coreBlock in blocks)
        {
            // 글상자(TextBoxObject) — FlowDocument 앵커 없이 FloatingCanvas 에 직접 배치됨.
            // 모델에 먼저 등록하고, 캔버스 배치는 아래 RebuildFloatingObjects() 가 일괄 수행.
            if (coreBlock is PolyDonky.Core.TextBoxObject tbo)
            {
                tbo.OverlayXMm += 5; tbo.OverlayYMm += 5;
                _viewModel?.AddOverlayBlockToCurrentSection(tbo);
                hasOverlay = true;
                continue;
            }

            // 오버레이 도형 — RTB 앵커 없이 모델에만 등록(RebuildOverlayShapes 가 모델을 읽음).
            // 앵커를 RTB 에 넣으면 ParseAllPageEditors 가 본문 블록으로 파싱해 모델의 도형과 중복된다.
            if (coreBlock is PolyDonky.Core.ShapeObject shpOvl &&
                shpOvl.WrapMode is PolyDonky.Core.ImageWrapMode.InFrontOfText or PolyDonky.Core.ImageWrapMode.BehindText)
            {
                shpOvl.OverlayXMm += 5; shpOvl.OverlayYMm += 5;
                _viewModel?.AddShapeToCurrentSection(shpOvl);
                hasOverlay = true;
                continue;
            }

            // 오버레이 이미지 — RTB 앵커 없이 모델에만 등록(RebuildOverlayImages 가 모델을 읽음).
            if (coreBlock is PolyDonky.Core.ImageBlock imgBlk &&
                imgBlk.WrapMode is PolyDonky.Core.ImageWrapMode.InFrontOfText or PolyDonky.Core.ImageWrapMode.BehindText)
            {
                imgBlk.OverlayXMm += 5; imgBlk.OverlayYMm += 5;
                _viewModel?.AddOverlayBlockToCurrentSection(imgBlk);
                hasOverlay = true;
                continue;
            }

            // 오버레이 표 — RTB 앵커 없이 모델에만 등록(RebuildOverlayTables 가 모델을 읽음).
            if (coreBlock is PolyDonky.Core.Table tblOvl &&
                tblOvl.WrapMode != PolyDonky.Core.TableWrapMode.Block)
            {
                tblOvl.OverlayXMm += 5; tblOvl.OverlayYMm += 5;
                _viewModel?.AddOverlayBlockToCurrentSection(tblOvl);
                hasOverlay = true;
                continue;
            }

            var wpfBlock = BuildWpfBlockFromCore(coreBlock);
            if (wpfBlock is null) continue;
            if (insertBefore != null)
            {
                doc.Blocks.InsertBefore(insertBefore, wpfBlock);
                insertBefore = wpfBlock.NextBlock;
            }
            else if (anchor != null)
            {
                doc.Blocks.InsertAfter(anchor, wpfBlock);
                anchor = wpfBlock;
            }
            else
            {
                doc.Blocks.Add(wpfBlock);
            }
        }

        // 커서를 마지막 삽입 블록 끝으로 이동 (비-오버레이 블록 붙여넣기)
        if (!hasOverlay && anchor != null)
        {
            try { BodyEditor.CaretPosition = anchor.ContentEnd; } catch { }
        }

        if (hasOverlay)
        {
            RebuildOverlayImages(); RebuildOverlayShapes(); RebuildOverlayTables();
            RebuildFloatingObjects();
            // 오버레이 재구축 후 포커스가 캔버스에 묶일 수 있으므로 본문 RTB 로 복원
            BodyEditor.Focus();
        }
        _viewModel?.MarkDirty();
        return true;
    }

    /// <summary>
    private static bool IsCaretInTableCell(System.Windows.Documents.TextPointer caret)
    {
        var el = caret.Parent as System.Windows.FrameworkContentElement;
        while (el != null)
        {
            if (el is System.Windows.Documents.TableCell) return true;
            el = el.Parent as System.Windows.FrameworkContentElement;
        }
        return false;
    }

    /// FlowDocument 의 최상위 Block 중 caret 을 포함하는 것을 반환한다.
    /// caret 이 TableCell 안에 있으면 parent chain 으로 Table 을 직접 찾아 반환 —
    /// ContentRange 비교만 사용하면 WPF Table 의 InsertAfter 가 셀 내부에 삽입되는 문제가 발생한다.
    private static System.Windows.Documents.Block? FindTopLevelAnchorForCaret(
        System.Windows.Documents.FlowDocument doc,
        System.Windows.Documents.TextPointer  caret)
    {
        // 1. parent chain 에서 doc.Blocks 에 포함된 최상위 Block 을 찾는다
        var el = caret.Parent as System.Windows.FrameworkContentElement;
        while (el != null)
        {
            if (el is System.Windows.Documents.Block b && doc.Blocks.Contains(b))
                return b;
            el = el.Parent as System.Windows.FrameworkContentElement;
        }

        // 2. 폴백: ContentRange 비교 (부유 객체·일반 단락 등 parent chain 이 단순한 경우)
        foreach (var b in doc.Blocks)
        {
            if (PolyDonky.App.Services.PageBreakPadder.IsPagePadding(b)) continue;
            if (b.ContentStart.CompareTo(caret) <= 0 && caret.CompareTo(b.ContentEnd) <= 0)
                return b;
        }
        return null;
    }

    /// <summary>
    /// 표 셀 안에 여러 단락을 붙여넣는다. 첫 단락은 캐럿 위치에 inline 으로 합치고,
    /// 나머지 단락은 셀의 Blocks 컬렉션에 새 Paragraph 로 차례로 삽입한다.
    /// </summary>
    private void InsertParagraphsIntoCell(
        List<PolyDonky.Core.Paragraph>     paragraphs,
        System.Windows.Documents.TextPointer caret)
    {
        if (paragraphs.Count == 0) return;

        // 첫 단락은 캐럿 위치에 inline 합치기
        InsertParagraphInline(paragraphs[0], caret);

        if (paragraphs.Count == 1) return;

        // 캐럿 단락의 부모 셀을 찾아 그 Blocks 컬렉션에 추가 — 호출 후 caret 은 합쳐진 단락 안에 있다.
        var insertPos  = caret.GetInsertionPosition(System.Windows.Documents.LogicalDirection.Forward);
        var targetPara = insertPos.Paragraph;
        if (targetPara?.Parent is not System.Windows.Documents.TableCell cell) return;

        System.Windows.Documents.Block lastInserted = targetPara;
        for (int i = 1; i < paragraphs.Count; i++)
        {
            var newPara = Services.FlowDocumentBuilder.BuildParagraph(paragraphs[i]);
            cell.Blocks.InsertAfter(lastInserted, newPara);
            lastInserted = newPara;
        }
    }

    /// <summary>
    /// Core.Paragraph 의 Inline 들을 캐럿 위치 단락에 인라인으로 삽입한다.
    /// 캐럿이 Run 중간에 있으면 그 Run 을 분리하고 사이에 삽입한다.
    /// </summary>
    /// <returns>삽입된 마지막 인라인의 ContentEnd (커서 복원용). 삽입 실패 시 null.</returns>
    private System.Windows.Documents.TextPointer? InsertParagraphInline(
        PolyDonky.Core.Paragraph corePara,
        System.Windows.Documents.TextPointer caretPos)
    {
        var tempPara = Services.FlowDocumentBuilder.BuildParagraph(corePara);
        var inlines = tempPara.Inlines.ToList();
        foreach (var il in inlines.ToArray()) tempPara.Inlines.Remove(il);  // parent 분리

        var insertPos  = caretPos.GetInsertionPosition(System.Windows.Documents.LogicalDirection.Forward);
        var targetPara = insertPos.Paragraph;

        if (targetPara is null)
        {
            string plain = string.Concat(corePara.Runs.Select(r => r.Text));
            if (!string.IsNullOrEmpty(plain)) insertPos.InsertTextInRun(plain);
            return null;
        }

        if (inlines.Count == 0) return null;

        System.Windows.Documents.Run? splitRun = null;
        int splitOffset = 0;
        System.Windows.Documents.Inline? insertAfter = null;

        foreach (var il in targetPara.Inlines)
        {
            if (il.ContentEnd.CompareTo(insertPos) <= 0) { insertAfter = il; continue; }
            if (il.ContentStart.CompareTo(insertPos) >= 0) break;
            if (il is System.Windows.Documents.Run r)
            {
                splitOffset = new System.Windows.Documents.TextRange(r.ContentStart, insertPos).Text.Length;
                if (splitOffset > 0 && splitOffset < r.Text.Length)
                    splitRun = r;
                else
                    insertAfter = splitOffset == 0 ? null : (System.Windows.Documents.Inline)r;
            }
            break;
        }

        System.Windows.Documents.Inline lastPasted;
        if (splitRun != null)
        {
            var before = CloneWpfRun(splitRun, splitRun.Text[..splitOffset]);
            var after  = CloneWpfRun(splitRun, splitRun.Text[splitOffset..]);
            targetPara.Inlines.InsertBefore(splitRun, before);
            foreach (var il in inlines) targetPara.Inlines.InsertBefore(splitRun, il);
            targetPara.Inlines.InsertBefore(splitRun, after);
            targetPara.Inlines.Remove(splitRun);
            lastPasted = inlines[inlines.Count - 1];
        }
        else if (insertAfter != null)
        {
            foreach (var il in Enumerable.Reverse(inlines))
                targetPara.Inlines.InsertAfter(insertAfter, il);
            lastPasted = inlines[inlines.Count - 1];
        }
        else
        {
            var first = targetPara.Inlines.FirstInline;
            if (first != null)
                foreach (var il in Enumerable.Reverse(inlines)) targetPara.Inlines.InsertBefore(first, il);
            else
                foreach (var il in inlines) targetPara.Inlines.Add(il);
            lastPasted = inlines[inlines.Count - 1];
        }

        try { return lastPasted.ContentEnd; } catch { return null; }
    }

    private static System.Windows.Documents.Run CloneWpfRun(System.Windows.Documents.Run src, string text)
    {
        var r = new System.Windows.Documents.Run(text)
        {
            FontFamily        = src.FontFamily,
            FontSize          = src.FontSize,
            FontWeight        = src.FontWeight,
            FontStyle         = src.FontStyle,
            Foreground        = src.Foreground,
            Background        = src.Background,
            BaselineAlignment = src.BaselineAlignment,
            Tag               = src.Tag,
        };
        foreach (var td in src.TextDecorations) r.TextDecorations.Add(td);
        return r;
    }

    /// <summary>BodyEditor.Selection 영역에 걸쳐 있는 모든 Block 을 Core.Block 으로 추출해 깊은 복사.</summary>
    private List<PolyDonky.Core.Block> ExtractCoreSelection()
    {
        var sel = BodyEditor.Selection;
        var result = new List<PolyDonky.Core.Block>();
        if (sel.IsEmpty) return result;

        foreach (var block in BodyEditor.Document.Blocks)
        {
            // 합성 페이지-갭 패딩 단락은 클립보드에 절대 포함되면 안 됨.
            if (PolyDonky.App.Services.PageBreakPadder.IsPagePadding(block)) continue;
            if (block.ContentEnd.CompareTo(sel.Start) <= 0) continue;
            if (block.ContentStart.CompareTo(sel.End) >= 0) break;

            if (block is System.Windows.Documents.Table wpfTable)
                EnsureCoreTable(wpfTable);

            PolyDonky.Core.Block? core;

            // 일반 Paragraph 가 선택 범위를 부분적으로 덮는 경우 → 선택된 Run/Inline 만 추출.
            // 오버레이 앵커 단락(Tag = Table/ImageBlock/ShapeObject) 은 통째로 추출.
            if (block is System.Windows.Documents.Paragraph wpfPara
                && block.Tag is not PolyDonky.Core.Table
                && block.Tag is not PolyDonky.Core.ImageBlock
                && block.Tag is not PolyDonky.Core.ShapeObject)
            {
                bool clippedAtStart = block.ContentStart.CompareTo(sel.Start) < 0;
                bool clippedAtEnd   = block.ContentEnd.CompareTo(sel.End)     > 0;
                core = (clippedAtStart || clippedAtEnd)
                    ? PolyDonky.App.Services.FlowDocumentParser.ParseParagraphClipped(wpfPara, sel.Start, sel.End)
                    : PolyDonky.App.Services.FlowDocumentParser.ParseSingleBlock(block);
            }
            else
            {
                core = PolyDonky.App.Services.FlowDocumentParser.ParseSingleBlock(block);
            }

            if (core is null) continue;

            try
            {
                var jsonClone = System.Text.Json.JsonSerializer.Serialize(core, PolyDonky.Core.JsonDefaults.Options);
                var clone = System.Text.Json.JsonSerializer.Deserialize<PolyDonky.Core.Block>(jsonClone, PolyDonky.Core.JsonDefaults.Options);
                if (clone != null) result.Add(clone);
            }
            catch { }
        }

        return result;
    }

    /// <summary>Core.Block 을 WPF Block 으로 변환 (FlowDocumentBuilder 디스패처).</summary>
    private static System.Windows.Documents.Block? BuildWpfBlockFromCore(PolyDonky.Core.Block coreBlock)
    {
        return coreBlock switch
        {
            PolyDonky.Core.Table t when t.WrapMode == PolyDonky.Core.TableWrapMode.Block
                => PolyDonky.App.Services.FlowDocumentBuilder.BuildTable(t),
            // 오버레이 표 — 앵커 단락만 삽입하고 캔버스 시각 요소는 RebuildOverlayTables() 가 담당.
            PolyDonky.Core.Table t
                => PolyDonky.App.Services.FlowDocumentBuilder.BuildTableAnchor(t),
            PolyDonky.Core.ImageBlock img
                => PolyDonky.App.Services.FlowDocumentBuilder.BuildImage(img),
            PolyDonky.Core.ShapeObject sh
                => PolyDonky.App.Services.FlowDocumentBuilder.BuildShape(sh),
            PolyDonky.Core.Paragraph p
                => PolyDonky.App.Services.FlowDocumentBuilder.BuildParagraph(p),
            _ => null,
        };
    }

    /// <summary>붙여넣기로 새 인스턴스가 되었음을 표시 — ID 재발급 및 Modified 상태.</summary>
    private static void ResetCoreBlockId(PolyDonky.Core.Block block)
    {
        block.Id = null;
        block.Status = PolyDonky.Core.NodeStatus.Modified;

        // 자식 Block 도 ID 재발급 (중첩 표 셀 등)
        if (block is PolyDonky.Core.Table t)
        {
            foreach (var row in t.Rows)
                foreach (var cell in row.Cells)
                    foreach (var inner in cell.Blocks)
                        ResetCoreBlockId(inner);
        }
        // (Run 은 Id/Status 가 없음 — 자식 Run 재발급 불필요)
    }

    private void OnUnloaded(object sender, RoutedEventArgs e)
    {
        if (_statusTimer is not null)
        {
            _statusTimer.Stop();
            _statusTimer.Tick -= OnStatusTimerTick;
            _statusTimer = null;
        }

        _dictWindow?.ForceClose();
        _dictWindow = null;
    }

    private void OnStatusTimerTick(object? sender, EventArgs e)
    {
        _viewModel?.RefreshSystemKeys();
        _viewModel?.RefreshMemoryUsage();
    }

    private void OnPreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.F8
            && (Keyboard.Modifiers & (ModifierKeys.Control | ModifierKeys.Shift)) == (ModifierKeys.Control | ModifierKeys.Shift))
        {
            ToggleTypesettingMarks();
            e.Handled = true;
            return;
        }

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
                if (_marqueeSelecting)
                {
                    _marqueeSelecting = false;
                    DrawPreviewRect.Visibility = Visibility.Collapsed;
                    if (PaperHost.IsMouseCaptured) PaperHost.ReleaseMouseCapture();
                    e.Handled = true;
                }
                else if (_drawingTextBox || _drawingShape_active || _drawingPolyline_active)
                {
                    EndDrawingMode();
                    e.Handled = true;
                }
                else if (_selectedShape is not null)
                {
                    DeselectShape();
                    e.Handled = true;
                }
                else if (_multiSelectedControls.Count > 0)
                {
                    ClearMultiSelect();
                    e.Handled = true;
                }
                else if (_selectedOverlay is not null)
                {
                    // 1단계: 안쪽 본문 편집 중이면 chrome 만 선택 상태로 전환 (포커스를 overlay 로 이동).
                    //        이후 Ctrl+C 가 글상자 자체를 복사하도록 해주는 진입점.
                    // 2단계: chrome 만 선택된 상태에서 다시 누르면 완전 해제.
                    if (_selectedOverlay.IsEditorFocusWithin)
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

            case Key.A when (Keyboard.Modifiers & ModifierKeys.Control) != 0:
                // Ctrl+A — 본문 텍스트 + 모든 오버레이(이미지/도형/표/글상자) 통합 선택.
                // 본문 InnerEditor(글상자 안쪽) 포커스 중이면 양보 (글상자 안쪽 텍스트만 SelectAll).
                if (_selectedOverlay?.IsEditorFocusWithin == true) break;
                SelectAllIncludingOverlays();
                e.Handled = true;
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
            //
            // 본문 RichTextBox 의 multi-block 선택(이미지·표·도형 포함) 클립보드 보존은
            // CommandManager.AddPreviewExecutedHandler 에서 PolyDonky.FlowSelection.v1 포맷으로 처리.
            case Key.C when (Keyboard.Modifiers & ModifierKeys.Control) != 0:
                if (TryCopySelectedObject()) e.Handled = true;
                break;
            case Key.X when (Keyboard.Modifiers & ModifierKeys.Control) != 0:
                if (TryCutSelectedObject()) e.Handled = true;
                break;
            case Key.V when (Keyboard.Modifiers & ModifierKeys.Control) != 0:
                if (TryPasteSelectedObject()) e.Handled = true;
                break;

            case Key.L when (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control:
                OnFormatChar(this, new RoutedEventArgs());
                e.Handled = true;
                break;

            case Key.T when (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control:
                OnFormatPara(this, new RoutedEventArgs());
                e.Handled = true;
                break;

            case Key.P when (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control:
                OnPreviewClick(this, new RoutedEventArgs());
                e.Handled = true;
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
        else if (e.PropertyName == nameof(MainViewModel.IsWriteProtected))
        {
            bool ro = _viewModel.IsWriteProtected;
            foreach (var rtb in PageEditorHost.PageEditors)
                rtb.IsReadOnly = ro;
        }
    }

    private void ApplyFlowDocument(System.Windows.Documents.FlowDocument _)
    {
        // per-page 모드에서는 ViewModel 의 FlowDocument 를 직접 BodyEditor 에 대입하지 않는다.
        // 대신 _viewModel.Document 를 DocumentPaginator 로 페이지네이트한 뒤,
        // PerPageDocumentSplitter 가 생성한 슬라이스로 각 페이지 RTB 를 초기화한다.

        // 1. 페이지 기하 정보 갱신
        var page = _viewModel?.Document.Sections.FirstOrDefault()?.Page
                   ?? new PolyDonky.Core.PageSettings();
        _pageGeometry  = new PolyDonky.App.Services.PageGeometry(page);
        PaperHost.Width = _pageGeometry.PageWidthDip;
        _pageHeightDip  = _pageGeometry.PageHeightDip;

        // 2. _viewModel.Document 로부터 PaginatedDocument 캐시 갱신
        UpdatePaginatedDoc();

        // 3. 페이지별 RTB 초기화
        SetupPageEditors();

        // 4. 오버레이 전체 재구축 (내부에서 RebuildPageFrames 호출)
        RebuildOverlays();
    }

    /// <summary>
    /// _viewModel.Document 로부터 PaginatedDocument 를 계산해 캐시한다.
    /// 문서 로드 직후에만 호출 — 그 시점엔 _viewModel.Document 가 디스크와 일치한다.
    /// 편집 중에는 <see cref="ScheduleLivePaginationRefresh"/> 가 라이브 FlowDocument 에서
    /// 직접 모델을 파싱해 캐시를 갱신한다.
    /// </summary>
    private void UpdatePaginatedDoc()
    {
        var doc = _viewModel?.Document;
        if (doc is null || _pageGeometry is null) return;
        _currentPaginatedDoc = FlowDocumentPaginationAdapter.Paginate(doc);
    }

    /// <summary>
    /// _currentPaginatedDoc 으로부터 페이지별 슬라이스를 만들고 PageEditorHost 를 초기화한다.
    /// _currentPaginatedDoc 이 null 이거나 _pageGeometry 가 없으면 아무 것도 하지 않는다.
    /// </summary>
    private void SetupPageEditors()
    {
        if (_currentPaginatedDoc is null || _pageGeometry is null) return;
        var slices = PerPageDocumentSplitter.Split(_currentPaginatedDoc);
        _suppressTextChanged    = true;
        _suppressPasteCommand   = true;
        try
        {
            PageEditorHost.SetupPages(slices, _pageGeometry, ConfigurePageRtb);
        }
        finally
        {
            _suppressTextChanged = false;
        }
        // Input 우선순위까지 미뤄서 RestoreCaretToLastEditor 포커스 이벤트가 모두 처리된 뒤 해제
        Dispatcher.BeginInvoke(DispatcherPriority.Input, () => _suppressPasteCommand = false);
    }

    /// <summary>
    /// 새로 생성된 페이지 RTB 에 이벤트 핸들러·속성을 등록하는 콜백.
    /// PageEditorHost.SetupPages 가 각 RTB 생성 직후 호출한다.
    /// </summary>
    private void ConfigurePageRtb(RichTextBox rtb)
    {
        // 테마 Foreground 바인딩 (FlowDocument 는 RTB Foreground 를 상속하지 않아 별도 바인딩 필요)
        rtb.Document.SetResourceReference(
            System.Windows.Documents.FlowDocument.ForegroundProperty, "OnSurface");
        rtb.Document.Background = Brushes.Transparent;

        rtb.IsReadOnly            = _viewModel?.IsWriteProtected ?? false;
        rtb.SpellCheck.IsEnabled  = false;
        rtb.Foreground            = (WpfMedia.Brush)FindResource("OnSurface");
        // 비활성(포커스 없는) RTB 에서도 선택 영역이 시각적으로 유지되도록.
        // 단 경계 교차 선택 시 이전 단의 선택을 표시하기 위해 필요.
        rtb.IsInactiveSelectionHighlightEnabled = true;
        rtb.GotKeyboardFocus += OnPageRtbGotFocusClearCrossSel;

        rtb.PreviewKeyDown    += OnEditorPreviewKeyDown;
        rtb.PreviewTextInput  += OnEditorPreviewTextInput;

        rtb.PreviewMouseLeftButtonDown += OnEditorPreviewMouseDownTrackDrag;
        rtb.PreviewMouseMove           += OnEditorPreviewMouseMoveBlockDrag;
        rtb.PreviewMouseLeftButtonUp   += OnEditorPreviewMouseUpEmbedded;
        // RTB 내부 ScrollViewer 가 휠 이벤트를 소비해 본문만 내부 스크롤되는 것을 막고
        // 외부 EditorScrollViewer 로 전달 — 본문·오버레이가 함께 스크롤되도록 한다.
        rtb.PreviewMouseWheel          += OnPageRtbPreviewMouseWheel;

        rtb.ContextMenu             = new System.Windows.Controls.ContextMenu();
        rtb.ContextMenuOpening      += OnEmbeddedObjectContextMenuOpening;
        rtb.PreviewMouseDoubleClick += OnEmbeddedObjectDoubleClick;

        rtb.MouseLeave += (_, _) =>
        {
            if (!_tableColResizeActive)
            {
                _tableColResizeHovering = false;
                Mouse.OverrideCursor = _colDivHovering ? Cursors.SizeWE : null;
            }
        };

        // _lastTextEditor 갱신 — 마지막으로 포커스를 가진 RTB 를 추적한다.
        rtb.GotKeyboardFocus += (_, _) => _lastTextEditor = rtb;

        DataObject.AddPastingHandler(rtb, OnBodyEditorPasting);
        CommandManager.AddPreviewExecutedHandler(rtb, OnBodyEditorPreviewExecuted);
    }

    /// <summary>
    /// 페이지 RTB 의 마우스 휠 이벤트를 외부 EditorScrollViewer 로 전달한다.
    /// 기본 동작: RTB 내부 ScrollViewer 가 휠을 소비해 본문만 내부 스크롤 → 오버레이와 어긋남.
    /// 처리: e.Handled = true 로 RTB 의 처리를 차단하고, 동일 Delta 의 새 MouseWheelEventArgs 를
    /// EditorScrollViewer 로 RaiseEvent 해서 페이지·오버레이가 함께 스크롤되도록 한다.
    /// </summary>
    private void OnPageRtbPreviewMouseWheel(object sender, MouseWheelEventArgs e)
    {
        if (EditorScrollViewer is null) return;
        e.Handled = true;
        var newArgs = new MouseWheelEventArgs(e.MouseDevice, e.Timestamp, e.Delta)
        {
            RoutedEvent = UIElement.MouseWheelEvent,
            Source      = sender,
        };
        EditorScrollViewer.RaiseEvent(newArgs);
    }

    /// <summary>
    /// 모든 페이지 RTB 를 파싱해 본문 블록을 결합하고, _viewModel.Document 의 오버레이 블록을 추가한 뒤
    /// 새로운 PolyDonkyument 를 반환한다.
    /// </summary>
    private PolyDonkyument ParseAllPageEditors()
    {
        var original = _viewModel?.Document;
        var freshDoc = new PolyDonkyument();
        if (original is not null)
        {
            freshDoc.Metadata      = original.Metadata;
            freshDoc.Styles        = original.Styles;
            freshDoc.Provenance    = original.Provenance;
            freshDoc.Watermark     = original.Watermark;
            freshDoc.OutlineStyles = original.OutlineStyles;
        }

        var section = new PolyDonky.Core.Section();
        if (original?.Sections.FirstOrDefault() is { } origSection)
            section.Page = origSection.Page;
        freshDoc.Sections.Add(section);

        // 각 페이지 RTB 를 파싱해 본문 블록 수집 (오버레이는 제외 — per-page RTB 에는 없다)
        var rawBlocks = new System.Collections.Generic.List<PolyDonky.Core.Block>();
        foreach (var rtb in PageEditorHost.PageEditors)
        {
            var pageDoc = PolyDonky.App.Services.FlowDocumentParser.Parse(rtb.Document);
            if (pageDoc.Sections.FirstOrDefault() is { } ps)
                foreach (var b in ps.Blocks)
                    rawBlocks.Add(b);
        }
        // 줄 단위 분할로 생성된 조각 단락(§f0/§f1 접미사)을 원본 단락으로 재결합한다.
        foreach (var b in MergeColumnFragments(rawBlocks))
            section.Blocks.Add(b);

        // 오버레이 블록 (_viewModel.Document 가 stable source) — 글상자는 RTB 에 앵커가 없고,
        // 오버레이 모드 표/그림/도형의 앵커 단락은 PerPageDocumentSplitter 가 BodyBlocks 에서 제외하므로
        // 둘 다 RTB 파싱으로 복원되지 않는다. 따라서 모델에서 직접 인계.
        // 주의: Table/ImageBlock/ShapeObject 는 모두 IOverlayAnchored 를 구현하므로 단순히
        // `b is IOverlayAnchored` 로 거르면 block-mode 표 등 본문 블록까지 포함되어 RTB 파싱본과
        // 중복되고 ("표가 한 개 더 맨 끝에 붙는" 증상) 페이지 끝에 추가 배정된다.
        // IsOverlayMode 로 진짜 오버레이만 추린다.
        if (original?.Sections.FirstOrDefault() is { } origOverlay)
        {
            foreach (var b in origOverlay.Blocks)
            {
                if (Pagination.FlowDocumentPaginationAdapter.IsOverlayMode(b))
                    section.Blocks.Add(b);
            }
        }

        return freshDoc;
    }

    /// <summary>
    /// MapBodyBlocksToPages 의 줄 단위 분할이 생성한 조각 단락들을 재결합한다.
    /// 조각 단락은 Id 에 "§f0" / "§f1" 접미사를 포함한다.
    /// 첫 조각(§f0)을 결과에 추가하고, 이어지는 조각(§f1, §f2 …)의 텍스트 런을 거기에 병합한다.
    /// </summary>
    private static System.Collections.Generic.IEnumerable<PolyDonky.Core.Block>
        MergeColumnFragments(System.Collections.Generic.IList<PolyDonky.Core.Block> blocks)
    {
        const string FragSep   = "§f";
        const string GenPrefix = "§g";

        var result      = new System.Collections.Generic.List<PolyDonky.Core.Block>();
        var openFrags   = new System.Collections.Generic.Dictionary<string, PolyDonky.Core.Paragraph>();
        var seenFragIdx = new System.Collections.Generic.Dictionary<string, System.Collections.Generic.HashSet<string>>();

        foreach (var block in blocks)
        {
            if (block is PolyDonky.Core.Paragraph p && p.Id is { } id)
            {
                int sepIdx = id.LastIndexOf(FragSep, StringComparison.Ordinal);
                if (sepIdx >= 0)
                {
                    string groupId     = id[..sepIdx];
                    string fragIdxStr  = id[(sepIdx + FragSep.Length)..];

                    // 같은 group 의 같은 fragIdx 가 두 번째 등장이면 — WPF 엔터 분할로 인해 같은
                    // §f Id 를 공유하게 된 두 반쪽 중 둘째. 사용자가 새로 만든 단락이므로
                    // §f 머지 대상에서 빼고 그대로 추가한다(엔터가 시각적으로 사라지는 버그 방지).
                    if (!seenFragIdx.TryGetValue(groupId, out var seen))
                    {
                        seen = new System.Collections.Generic.HashSet<string>(StringComparer.Ordinal);
                        seenFragIdx[groupId] = seen;
                    }
                    if (!seen.Add(fragIdxStr))
                    {
                        p.Id = null;
                        result.Add(p);
                        continue;
                    }

                    bool isFirst = fragIdxStr == "0";
                    if (isFirst)
                    {
                        // Id 를 원본으로 복원 (임시 생성 Id 이면 null)
                        p.Id = groupId.StartsWith(GenPrefix, StringComparison.Ordinal) ? null : groupId;
                        result.Add(p);
                        openFrags[groupId] = p;
                    }
                    else
                    {
                        if (openFrags.TryGetValue(groupId, out var target))
                            foreach (var run in p.Runs) target.Runs.Add(run);
                        else
                        {
                            // 짝 없는 이어지는 조각 — 단독으로 추가
                            p.Id = groupId.StartsWith(GenPrefix, StringComparison.Ordinal) ? null : groupId;
                            result.Add(p);
                        }
                    }
                    continue;
                }
            }
            result.Add(block);
        }

        return result;
    }

    private bool _liveRefreshQueued;
    private bool _isImeComposing;
    private bool _pendingLiveRefresh;

    private void OnImeCompositionStart(object sender, System.Windows.Input.TextCompositionEventArgs e)
    {
        _isImeComposing = true;
    }

    private void OnImeCompositionEnd(object sender, System.Windows.Input.TextCompositionEventArgs e)
    {
        _isImeComposing = false;
        // 조합 중에 보류된 재페이지네이션 요청이 있으면 지금 처리.
        if (_pendingLiveRefresh)
        {
            _pendingLiveRefresh = false;
            ScheduleLivePaginationRefresh();
        }
    }

    /// <summary>
    /// 라이브 FlowDocument 를 Parse → Paginate 해서 _currentPaginatedDoc 캐시를 최신화한다.
    /// Background 우선순위로 디스패치해 UI 응답성 보장.
    /// 동시에 여러 번 큐잉되지 않도록 _liveRefreshQueued 플래그로 합치기.
    /// 페이지 수가 동일하고 페이지별 본문 블록 수도 동일하면 RTB 는 재구성하지 않아 커서 위치를 유지한다.
    /// 페이지 수 또는 페이지별 블록 분배가 달라지면(오버플로우/언더플로우/표 삽입) RTB 를 재구성해
    /// 본문을 페이지 경계에 맞게 재분배한다 — 이 경우 커서는 마지막 RTB 끝으로 이동한다.
    /// </summary>
    private void ScheduleLivePaginationRefresh()
    {
        if (_liveRefreshQueued) return;
        if (_viewModel?.Document is null || _pageGeometry is null) return;

        _liveRefreshQueued = true;
        Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Background, () =>
        {
            _liveRefreshQueued = false;

            // IME 조합 중에는 재구성을 보류 — RTB 재생성이 IME 조합 상태를 깨뜨려
            // 한글 한 글자(ㅈ + ㅏ)가 두 글자로 분리되는 문제를 방지.
            // 조합이 끝나면 OnImeCompositionEnd 가 다시 호출한다.
            if (_isImeComposing)
            {
                _pendingLiveRefresh = true;
                return;
            }

            try
            {
                // 모든 페이지 RTB 를 파싱해 결합된 PolyDonkyument 를 만들고 재페이지네이트.
                var freshDoc = ParseAllPageEditors();
                _currentPaginatedDoc = FlowDocumentPaginationAdapter.Paginate(freshDoc);

                // 재구성 필요 판정:
                //   ① 페이지 수가 달라졌거나
                //   ② 페이지별 본문 블록 수 또는 ID 시퀀스가 달라졌으면
                //      (예: 단락이 자라서 다음 페이지로 밀려났지만 다른 단락이 채워져 개수가 같은 경우 포함)
                // RTB 를 재구성해 본문을 페이지별로 다시 분배한다.
                bool needsRebuild = NeedsPageRebuild();
                if (needsRebuild)
                {
                    var savedCaret = SaveCaretState();
                    SetupPageEditors();
                    RestoreCaretState(savedCaret);
                }

                RebuildPageFrames();
            }
            catch
            {
                // 실패 시 캐시 유지.
            }
        });
    }

    /// <summary>
    /// 새로 계산된 <see cref="_currentPaginatedDoc"/> 의 페이지별 블록 분배가
    /// 현재 페이지 RTB 들과 다르면 true.
    /// 블록 수뿐 아니라 블록 ID 시퀀스까지 비교해 "개수는 같지만 다른 블록으로 채워진" 경우도 검출한다.
    /// 다단의 경우: 페이지당 단 수 × 페이지 수 = 총 RTB 수이므로 각 페이지의 모든 단 RTB 에서
    /// 블록 ID 를 수집해 페이지 전체 블록 목록과 비교한다.
    /// </summary>
    private bool NeedsPageRebuild()
    {
        if (_currentPaginatedDoc is null) return false;

        int newPageCount = _currentPaginatedDoc.PageCount;
        int oldPageCount = PageEditorHost.PageCount;
        if (newPageCount != oldPageCount) return true;

        int colCount = _pageGeometry?.ColumnCount ?? 1;
        var rtbs     = PageEditorHost.PageEditors;

        for (int i = 0; i < newPageCount; i++)
        {
            var newPage = _currentPaginatedDoc.Pages[i];

            if (colCount <= 1)
            {
                // 단일 단: RTB 인덱스 == 페이지 인덱스
                if (i >= rtbs.Count) return true;
                var oldIds = CollectBodyBlockIds(rtbs[i].Document.Blocks);

                if (newPage.BodyBlocks.Count != oldIds.Count) return true;

                for (int j = 0; j < newPage.BodyBlocks.Count; j++)
                {
                    var newId = newPage.BodyBlocks[j].Source.Id;
                    var oldId = oldIds[j];
                    // 둘 다 ID 가 있으면 직접 비교. 하나라도 null 이면 신규 단락이므로 카운트만 체크.
                    if (newId != null && oldId != null && newId != oldId)
                        return true;
                }
            }
            else
            {
                // 다단: 페이지 내 단별 RTB 의 블록 분배를 새 결과와 비교한다.
                // 페이지 합계만 비교하면 "단끼리 재분배되었지만 합은 같은" 경우(긴 단락이
                // col 0 에서 col 1 로 넘어가야 하는 케이스) 를 놓쳐 RTB 가 stale 상태로 남는다.
                for (int col = 0; col < colCount; col++)
                {
                    int rtbIdx = i * colCount + col;
                    if (rtbIdx >= rtbs.Count) return true;

                    var oldIds = CollectBodyBlockIds(rtbs[rtbIdx].Document.Blocks);
                    var newColBlocks = newPage.BodyBlocks.Where(b => b.ColumnIndex == col).ToList();

                    if (newColBlocks.Count != oldIds.Count) return true;

                    for (int j = 0; j < newColBlocks.Count; j++)
                    {
                        var newId = newColBlocks[j].Source.Id;
                        var oldId = oldIds[j];
                        if (newId != null && oldId != null && newId != oldId)
                            return true;
                    }
                }
            }
        }
        return false;
    }

    /// <summary>
    /// FlowDocument.Blocks 를 재귀적으로 순회해 본문 블록의 ID 목록을 반환한다
    /// (List 내부 ListItem.Blocks 포함, 오버레이 앵커 제외).
    /// </summary>
    private static List<string?> CollectBodyBlockIds(System.Windows.Documents.BlockCollection blocks)
    {
        var ids = new List<string?>();
        foreach (var b in blocks)
        {
            if (b.Tag is PolyDonky.Core.Block coreBlock
                && !Pagination.FlowDocumentPaginationAdapter.IsOverlayMode(coreBlock))
                ids.Add(coreBlock.Id);
            if (b is System.Windows.Documents.List list)
                foreach (var li in list.ListItems)
                    ids.AddRange(CollectBodyBlockIds(li.Blocks));
        }
        return ids;
    }

    private record CaretState(int PageIndex, int Offset, int VirtualDocOffset = -1);

    /// <summary>
    /// 현재 포커스된 페이지 RTB 와 캐럿 오프셋을 저장한다.
    /// SetupPageEditors 이전에 호출해 재구성 후 커서 위치를 복원한다.
    /// </summary>
    private CaretState? SaveCaretState()
    {
        var rtbs = PageEditorHost.PageEditors;
        for (int i = 0; i < rtbs.Count; i++)
        {
            var rtb = rtbs[i];
            if (!rtb.IsKeyboardFocusWithin && !rtb.IsFocused) continue;
            try
            {
                int rtbOffset = rtb.Document.ContentStart
                    .GetOffsetToPosition(rtb.CaretPosition);

                // 가상 문서 위치(VDP) = 이전 RTB 들의 콘텐츠 합산 + 이 RTB 내 오프셋.
                // 단 간 블록 재분배 후에도 전체 VDP 합은 보존되므로 복원 시
                // 어느 RTB 로 블록이 이동했는지 자동으로 추적된다.
                int vdp = rtbOffset;
                for (int j = 0; j < i; j++)
                {
                    vdp += rtbs[j].Document.ContentStart
                        .GetOffsetToPosition(rtbs[j].Document.ContentEnd);
                }

                return new CaretState(i, rtbOffset, vdp);
            }
            catch { break; }
        }
        return null;
    }

    /// <summary>
    /// <see cref="SaveCaretState"/> 로 저장한 위치로 커서를 복원한다.
    /// VDP 방식을 우선 시도하고, 실패하면 RTB 인덱스·오프셋으로 폴백한다.
    /// </summary>
    private void RestoreCaretState(CaretState? saved)
    {
        var rtbs = PageEditorHost.PageEditors;
        if (rtbs.Count == 0) return;

        if (saved is not null)
        {
            // 1순위: 가상 문서 위치로 복원
            // — 블록이 다른 단 RTB 로 이동했을 때도 올바른 위치를 찾는다.
            if (saved.VirtualDocOffset >= 0 && TryRestoreByVdp(rtbs, saved.VirtualDocOffset))
                return;

            // 2순위: 원래 RTB 인덱스 + 오프셋 (폴백)
            int pageIdx = Math.Min(saved.PageIndex, rtbs.Count - 1);
            var rtb = rtbs[pageIdx];
            try
            {
                var start     = rtb.Document.ContentStart;
                int maxOffset = start.GetOffsetToPosition(rtb.Document.ContentEnd);
                int target    = Math.Clamp(saved.Offset, 0, maxOffset);
                var pos = start.GetPositionAtOffset(
                    target, System.Windows.Documents.LogicalDirection.Forward);
                if (pos != null)
                {
                    rtb.CaretPosition = pos;
                    rtb.Focus();
                    Keyboard.Focus(rtb);
                    return;
                }
            }
            catch { }
        }

        RestoreCaretToLastEditor();
    }

    private bool TryRestoreByVdp(System.Collections.Generic.IReadOnlyList<System.Windows.Controls.RichTextBox> rtbs, int vdp)
    {
        int remaining = vdp;
        for (int i = 0; i < rtbs.Count; i++)
        {
            var rtb    = rtbs[i];
            int rtbLen = rtb.Document.ContentStart
                             .GetOffsetToPosition(rtb.Document.ContentEnd);

            bool isLast = i == rtbs.Count - 1;
            if (remaining <= rtbLen || isLast)
            {
                try
                {
                    int target = Math.Clamp(remaining, 0, rtbLen);
                    var pos = rtb.Document.ContentStart.GetPositionAtOffset(
                        target, System.Windows.Documents.LogicalDirection.Forward);
                    if (pos != null)
                    {
                        rtb.CaretPosition = pos;
                        rtb.Focus();
                        Keyboard.Focus(rtb);
                        return true;
                    }
                }
                catch { }
                return false;
            }
            remaining -= rtbLen;
        }
        return false;
    }

    /// <summary>
    /// SetupPageEditors 후 마지막 페이지 RTB 끝에 커서를 두고 포커스를 준다.
    /// <see cref="RestoreCaretState"/> 가 실패했을 때 최후 수단으로 호출한다.
    /// </summary>
    private void RestoreCaretToLastEditor()
    {
        var rtbs = PageEditorHost.PageEditors;
        if (rtbs.Count == 0) return;
        var last = rtbs[rtbs.Count - 1];
        last.CaretPosition = last.Document.ContentEnd;
        last.Focus();
        Keyboard.Focus(last);
    }

    /// <summary>
    /// section.Blocks 전체를 훑어 InFrontOfText / BehindText / 글상자 등 모든
    /// 부유 객체를 적절한 캔버스에 다시 채워 그린다.
    /// 종류별 세부 로직은 RebuildOverlayImages / RebuildOverlayShapes /
    /// RebuildOverlayTables / RebuildFloatingObjects 가 담당.
    /// </summary>
    private void RebuildOverlays()
    {
        RebuildFloatingObjects();   // 글상자 (TextBoxObject)
        RebuildOverlayImages();     // 이미지 (ImageBlock)
        RebuildOverlayShapes();     // 도형 (ShapeObject)
        RebuildOverlayTables();     // 표 (Table)
        // 오버레이가 새 페이지로 이동·생성될 수 있으므로 페이지 프레임도 다시 계산.
        RebuildPageFrames();
    }

    /// <summary>
    /// Document 모델을 순회해 InFrontOfText / BehindText 모드 ImageBlock 들을
    /// OverlayImageCanvas / UnderlayImageCanvas 에 절대 위치로 배치한다.
    /// </summary>
    private void RebuildOverlayImages()
    {
        ClearMultiSelect();
        OverlayImageCanvas.Children.Clear();
        UnderlayImageCanvas.Children.Clear();

        // per-page 모드에서 오버레이 이미지는 어떤 페이지 RTB 에도 포함되지 않으므로
        // _viewModel.Document (삽입·삭제가 반영된 stable source) 를 직접 순회한다.
        var overlaySection = _viewModel?.Document.Sections.FirstOrDefault();
        if (overlaySection is null) return;
        foreach (var coreBlock in overlaySection.Blocks)
        {
            if (coreBlock is not PolyDonky.Core.ImageBlock img) continue;
            if (img.WrapMode is not (PolyDonky.Core.ImageWrapMode.InFrontOfText
                                  or PolyDonky.Core.ImageWrapMode.BehindText)) continue;

            var ctrl = Services.FlowDocumentBuilder.BuildOverlayImageControl(img);
            if (ctrl is null) continue;

            ctrl.Tag = img;
            PlaceOverlay(ctrl, img);
            ctrl.Cursor = System.Windows.Input.Cursors.SizeAll;
            // 우클릭은 PaperHost.PreviewMouseRightButtonDown 통합 핸들러가 처리 — 개별 ContextMenu 불필요.

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

    // ── 표 열 너비 드래그 리사이즈 상태 ──────────────────────────────
    private const double TableColResizeHitDip = 8.0;
    private const double TableColResizeMinDip = 18.0;
    private bool   _tableColResizeActive;
    private bool   _tableColResizeHovering;
    private System.Windows.Documents.Table?       _colRszWpf;
    private PolyDonky.Core.Table?                 _colRszCore;
    private System.Windows.Documents.TableColumn? _colRszLeftCol;
    private System.Windows.Documents.TableColumn? _colRszRightCol;
    private int    _colRszLeftIdx;
    private double _colRszInitLeft;
    private double _colRszInitRight;
    private double _colRszStartX;

    // ── 단(column) 너비 드래그 리사이즈 상태 ──────────────────────────
    private const double ColDivHitDip    = 8.0;
    private const double ColDivMinWidthDip = 20.0;
    private bool     _colDivDragging;
    private bool     _colDivHovering;
    private int      _colDivDragLeftIdx;
    private double   _colDivDragStartX;
    private double[] _colDivDragStartWidths = Array.Empty<double>();
    private readonly List<WpfShapes.Line> _columnDividerLines = new();

    // ── 단 경계 교차 텍스트 선택 상태 ────────────────────────────────────
    // Shift+방향키로 단 경계를 넘을 때 비활성 RTB 의 선택을 유지(IsInactiveSelectionHighlightEnabled)
    // 하고, _crossSelActive 플래그로 복사/편집 명령이 여러 RTB 의 선택을 하나로 처리하도록 한다.
    private bool _crossSelActive;
    private bool _inCrossSelNavigation;  // 프로그래매틱 포커스 이동 중 GotKeyboardFocus 무시 플래그

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
            CommitOverlayDragPosition(img, left, top);
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

    private PolyDonky.App.Services.PageGeometry? _pageGeometry;
    private int                _currentPageCount    = 1;
    private PaginatedDocument? _currentPaginatedDoc;           // 로드 시점 또는 ScheduleLivePaginationRefresh 후 최신
    private bool               _suppressPageFrameRebuild;      // RebuildPageFrames 안에서 MinHeight 변경이 재귀로 돌아오는 것을 차단

    private void ApplyPageSettings(PolyDonky.Core.PageSettings? page)
    {
        if (page is null) return;

        _pageGeometry  = new PolyDonky.App.Services.PageGeometry(page);
        PaperHost.Width = _pageGeometry.PageWidthDip;
        _pageHeightDip  = _pageGeometry.PageHeightDip;

        // 편집 중 페이지 설정 변경 — 라이브 RTB 내용을 파싱해 새 설정으로 재페이지네이트.
        if (PageEditorHost.PageCount > 0)
        {
            var freshDoc = ParseAllPageEditors();
            if (freshDoc.Sections.FirstOrDefault() is { } s) s.Page = page;
            _currentPaginatedDoc = FlowDocumentPaginationAdapter.Paginate(freshDoc);
            SetupPageEditors();
            RebuildOverlays();

            // 첫 Paginate 의 오프스크린 RTB 측정이 GetCharacterRect=NaN 으로 실패해
            // 모든 본문 블록이 page 0 으로 몰릴 수 있다(시각 트리 밖에서 fresh FlowDocument
            // 의 텍스트 레이아웃이 확정되지 않는 케이스). RTB 가 시각 트리에 부착된 직후
            // Render 우선순위로 한 번 더 재페이지네이트해 정확한 분배를 강제한다.
            // 두 번째 측정은 PageEditorHost 자식 RTB 들이 이미 layout pass 를 거친 뒤이므로
            // GetCharacterRect 가 안정된 Y 를 반환한다.
            Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Render, () =>
            {
                if (_pageGeometry is null) return;
                var redoFresh = ParseAllPageEditors();
                if (redoFresh.Sections.FirstOrDefault() is { } rs) rs.Page = page;
                var redoPaginated = FlowDocumentPaginationAdapter.Paginate(redoFresh);
                _currentPaginatedDoc = redoPaginated;
                if (NeedsPageRebuild())
                {
                    var savedCaret = SaveCaretState();
                    SetupPageEditors();
                    RestoreCaretState(savedCaret);
                    RebuildOverlays();
                }
                else
                {
                    RebuildPageFrames();
                }
            });
        }
        else
        {
            RebuildPageFrames();
        }
    }

/// <summary>
    /// PageBackgroundCanvas 에 페이지 별 흰색 Border (그림자 + 점선 여백 가이드 + "N페이지" 라벨)를 다시 그린다.
    /// 페이지 수 우선순위:
    ///   1. <see cref="_currentPaginatedDoc"/> — 로드 시점·ScheduleLivePaginationRefresh 후 최신값.
    ///   2. PageEditorHost.PageCount / 오버레이 anchor max — 편집 중간 폴백.
    /// </summary>
    internal void RebuildPageFrames()
    {
        if (_pageGeometry is null) return;
        if (_suppressPageFrameRebuild) return;

        _suppressPageFrameRebuild = true;
        try
        {
            RebuildPageFramesCore();
        }
        finally
        {
            _suppressPageFrameRebuild = false;
        }
    }

    private void RebuildPageFramesCore()
    {
        var pg = _pageGeometry!;
        _columnDividerLines.Clear();

        // 페이지 수 계산.
        // 1순위: WPF DocumentPaginator 로 산출한 정확값 (_currentPaginatedDoc).
        // 2순위: 편집 중간 짧은 구간 — PageEditorHost 현재 페이지 수 또는 오버레이 anchor max.
        int maxAnchorIndex = ComputeMaxAnchorPageIndex();
        int pageCount = _currentPaginatedDoc?.PageCount
            ?? Math.Max(PageEditorHost.PageCount, maxAnchorIndex + 1);
        if (pageCount < 1) pageCount = 1;
        _currentPageCount = pageCount;

        // PaperHost 의 전체 높이 = N 페이지 + (N-1) 갭.
        double totalHeight = pg.TotalHeightDip(pageCount);
        if (Math.Abs(PaperHost.MinHeight - totalHeight) > 0.5)
            PaperHost.MinHeight = totalHeight;

        // 오버레이 캔버스를 페이지 경계에서 클립 — 본문 텍스트가 per-page RTB 에서 페이지마다
        // 잘려 보이는 것과 동일하게, 모든 부유 객체(글상자·도형·이미지·표) 도 페이지 경계 밖
        // (특히 페이지 간 갭) 에서 잘려 보이도록 한다. 미리보기/인쇄와 동일한 시각 결과를 보장한다.
        var overlayClip = PageViewBuilder.BuildPageClipGeometry(pg, pageCount);
        OverlayShapeCanvas.Clip  = overlayClip;
        UnderlayShapeCanvas.Clip = overlayClip;
        OverlayImageCanvas.Clip  = overlayClip;
        UnderlayImageCanvas.Clip = overlayClip;
        OverlayTableCanvas.Clip  = overlayClip;
        UnderlayTableCanvas.Clip = overlayClip;
        FloatingCanvas.Clip      = overlayClip;

        // PageBackgroundCanvas 클리어 후 페이지마다 다시 그리기.
        PageBackgroundCanvas.Children.Clear();

        var page = _viewModel?.Document.Sections.FirstOrDefault()?.Page;
        bool showGuides = page?.ShowMarginGuides ?? true;
        WpfMedia.Brush pageBg = ResolvePaperBackground(page);

        for (int i = 0; i < pageCount; i++)
        {
            double topY = i * pg.PageStrideDip;

            // 페이지 외곽 (흰색 Border + 그림자)
            var pageBorder = new Border
            {
                Width  = pg.PageWidthDip,
                Height = pg.PageHeightDip,
                Background  = pageBg,
                BorderBrush = (WpfMedia.Brush)FindResource("Divider"),
                BorderThickness = new Thickness(0.5),
                Effect = new System.Windows.Media.Effects.DropShadowEffect
                {
                    BlurRadius   = 18,
                    ShadowDepth  = 6,
                    Opacity      = 0.38,
                    Direction    = 270,
                    Color        = WpfMedia.Colors.Black,
                },
                IsHitTestVisible = false,
            };
            System.Windows.Controls.Canvas.SetLeft(pageBorder, 0);
            System.Windows.Controls.Canvas.SetTop (pageBorder, topY);
            PageBackgroundCanvas.Children.Add(pageBorder);

            // 여백 가이드 (점선 사각형) — 본문 영역
            if (showGuides)
            {
                var guide = new System.Windows.Shapes.Rectangle
                {
                    Width  = pg.PageWidthDip - pg.PadLeftDip - pg.PadRightDip,
                    Height = pg.PageHeightDip - pg.PadTopDip - pg.PadBottomDip,
                    Stroke = new SolidColorBrush(WpfMedia.Color.FromArgb(0x66, 0x00, 0x78, 0xD4)),
                    StrokeThickness = 0.7,
                    StrokeDashArray = new System.Windows.Media.DoubleCollection { 4, 3 },
                    Fill = WpfMedia.Brushes.Transparent,
                    IsHitTestVisible = false,
                };
                System.Windows.Controls.Canvas.SetLeft(guide, pg.PadLeftDip);
                System.Windows.Controls.Canvas.SetTop (guide, topY + pg.PadTopDip);
                PageBackgroundCanvas.Children.Add(guide);
            }

            // 단 구분선 — 다단인 경우 단 사이 갭 중앙에 세로선.
            // 표시 여부/색/두께/스타일은 PageSettings.ColumnDivider* 로 사용자 설정.
            // ShowMarginGuides 와는 독립적이며, 인쇄에는 출력되지 않는 편집창 전용 가이드.
            if (pg.ColumnCount > 1
                && page is not null
                && page.ColumnDividerVisible
                && page.ColumnDividerStyle != ColumnDividerStyle.None)
            {
                double bodyTop    = topY + pg.PadTopDip;
                double bodyHeight = pg.PageHeightDip - pg.PadTopDip - pg.PadBottomDip;
                var dividerBrush  = ParseColumnDividerBrush(page.ColumnDividerColor);
                double dividerThk = page.ColumnDividerThicknessPt * (96.0 / 72.0);
                System.Windows.Media.DoubleCollection? dividerDash = page.ColumnDividerStyle switch
                {
                    ColumnDividerStyle.Dashed => new System.Windows.Media.DoubleCollection { 4, 3 },
                    ColumnDividerStyle.Dotted => new System.Windows.Media.DoubleCollection { 1, 2 },
                    _                         => null,
                };

                for (int c = 0; c < pg.ColumnCount - 1; c++)
                {
                    double colW = c < pg.ColWidthsDip.Length ? pg.ColWidthsDip[c] : pg.ColWidthDip;
                    double divX = pg.PadLeftDip
                                + pg.ColumnXOffsetDip(c)
                                + colW
                                + pg.ColGapDip / 2.0;
                    var divider = new WpfShapes.Line
                    {
                        X1 = divX, X2 = divX,
                        Y1 = bodyTop, Y2 = bodyTop + bodyHeight,
                        Stroke = dividerBrush,
                        StrokeThickness = dividerThk,
                        IsHitTestVisible = false,
                        Tag = c,
                    };
                    if (dividerDash is not null) divider.StrokeDashArray = dividerDash;
                    _columnDividerLines.Add(divider);
                    PageBackgroundCanvas.Children.Add(divider);
                }
            }

            // 페이지 번호 라벨 (좌상단)
            var label = new System.Windows.Controls.TextBlock
            {
                Text       = $"{i + 1}페이지",
                FontSize   = 10,
                Foreground = new SolidColorBrush(WpfMedia.Color.FromArgb(0xA0, 0x50, 0x50, 0xB4)),
                IsHitTestVisible = false,
            };
            System.Windows.Controls.Canvas.SetLeft(label, 6);
            System.Windows.Controls.Canvas.SetTop (label, topY + 2);
            PageBackgroundCanvas.Children.Add(label);

        }

        RenderWatermark(pg, pageCount);
        RebuildTypesettingMarks();
    }

    // ── 조판부호 보기 ───────────────────────────────────────────────────────

    private void OnTypesettingMarksToggle(object sender, RoutedEventArgs e)
        => ToggleTypesettingMarks();

    private void ToggleTypesettingMarks()
    {
        _showTypesettingMarks = !_showTypesettingMarks;
        MiTypesettingMarks.IsChecked = _showTypesettingMarks;
        LanguageService.SetShowTypesettingMarks(_showTypesettingMarks);
        RebuildTypesettingMarks();
    }

    private void RebuildTypesettingMarks()
    {
        TypesettingMarksCanvas.Children.Clear();
        if (!_showTypesettingMarks || _pageGeometry is null) return;

        var pg        = _pageGeometry;
        int pageCount = _currentPageCount < 1 ? 1 : _currentPageCount;
        var page      = _viewModel?.Document.Sections.FirstOrDefault()?.Page;

        double headerBandH = page is not null
            ? Services.FlowDocumentBuilder.MmToDip(Math.Max(0.0, page.MarginTopMm - page.MarginHeaderMm))
            : 0.0;
        double footerBandH = page is not null
            ? Services.FlowDocumentBuilder.MmToDip(Math.Max(0.0, page.MarginBottomMm - page.MarginFooterMm))
            : 0.0;
        double headerOffsetDip = page is not null
            ? Services.FlowDocumentBuilder.MmToDip(page.MarginHeaderMm)
            : 0.0;

        var headerFill  = new SolidColorBrush(WpfMedia.Color.FromArgb(0x28, 0x00, 0x78, 0xD4));
        var footerFill  = new SolidColorBrush(WpfMedia.Color.FromArgb(0x28, 0xFF, 0x80, 0x00));
        var headerFg    = new SolidColorBrush(WpfMedia.Color.FromArgb(0xCC, 0x00, 0x4C, 0xAA));
        var footerFg    = new SolidColorBrush(WpfMedia.Color.FromArgb(0xCC, 0xB8, 0x50, 0x00));

        for (int i = 0; i < pageCount; i++)
        {
            double topY = i * pg.PageStrideDip;

            if (headerBandH > 1.0)
            {
                var band = new WpfShapes.Rectangle
                {
                    Width  = pg.PageWidthDip,
                    Height = headerBandH,
                    Fill   = headerFill,
                    IsHitTestVisible = false,
                };
                System.Windows.Controls.Canvas.SetLeft(band, 0);
                System.Windows.Controls.Canvas.SetTop (band, topY + headerOffsetDip);
                TypesettingMarksCanvas.Children.Add(band);

                var lbl = new System.Windows.Controls.TextBlock
                {
                    Text       = SR.TyposettingMarkHeader,
                    FontSize   = 9,
                    Foreground = headerFg,
                    IsHitTestVisible = false,
                };
                System.Windows.Controls.Canvas.SetLeft(lbl, pg.PadLeftDip);
                System.Windows.Controls.Canvas.SetTop (lbl, topY + headerOffsetDip + 1);
                TypesettingMarksCanvas.Children.Add(lbl);
            }

            if (footerBandH > 1.0)
            {
                double footerTopDip = topY + pg.PageHeightDip - pg.PadBottomDip;
                var band = new WpfShapes.Rectangle
                {
                    Width  = pg.PageWidthDip,
                    Height = footerBandH,
                    Fill   = footerFill,
                    IsHitTestVisible = false,
                };
                System.Windows.Controls.Canvas.SetLeft(band, 0);
                System.Windows.Controls.Canvas.SetTop (band, footerTopDip);
                TypesettingMarksCanvas.Children.Add(band);

                var lbl = new System.Windows.Controls.TextBlock
                {
                    Text       = SR.TyposettingMarkFooter,
                    FontSize   = 9,
                    Foreground = footerFg,
                    IsHitTestVisible = false,
                };
                System.Windows.Controls.Canvas.SetLeft(lbl, pg.PadLeftDip);
                System.Windows.Controls.Canvas.SetTop (lbl, footerTopDip + 1);
                TypesettingMarksCanvas.Children.Add(lbl);
            }
        }

        AddOverlayAnchorBadges();
    }

    private void AddOverlayAnchorBadges()
    {
        void AddBadge(double left, double top, string icon, string tip)
        {
            var border = new Border
            {
                Background    = new SolidColorBrush(WpfMedia.Color.FromArgb(0xCC, 0x00, 0x50, 0xC0)),
                CornerRadius  = new CornerRadius(3),
                Padding       = new Thickness(3, 1, 3, 1),
                IsHitTestVisible = false,
                ToolTip       = tip,
                Child = new System.Windows.Controls.TextBlock
                {
                    Text       = icon,
                    FontSize   = 9,
                    Foreground = WpfMedia.Brushes.White,
                    IsHitTestVisible = false,
                },
            };
            System.Windows.Controls.Canvas.SetLeft(border, left);
            System.Windows.Controls.Canvas.SetTop (border, top);
            TypesettingMarksCanvas.Children.Add(border);
        }

        void ScanCanvas(System.Windows.Controls.Canvas canvas, string icon, string tipKey)
        {
            foreach (System.Windows.FrameworkElement child in canvas.Children)
            {
                double l = System.Windows.Controls.Canvas.GetLeft(child);
                double t = System.Windows.Controls.Canvas.GetTop(child);
                if (!double.IsNaN(l) && !double.IsNaN(t))
                    AddBadge(l, t, icon, tipKey);
            }
        }

        ScanCanvas(OverlayImageCanvas,   "🖼", SR.TyposettingMarkImage);
        ScanCanvas(UnderlayImageCanvas,  "🖼", SR.TyposettingMarkImageBehind);
        ScanCanvas(OverlayShapeCanvas,   "△", SR.TyposettingMarkShape);
        ScanCanvas(UnderlayShapeCanvas,  "△", SR.TyposettingMarkShapeBehind);
        ScanCanvas(OverlayTableCanvas,   "⊞", SR.TyposettingMarkTable);
        ScanCanvas(UnderlayTableCanvas,  "⊞", SR.TyposettingMarkTableBehind);
        ScanCanvas(FloatingCanvas,       "▤", SR.TyposettingMarkTextBox);
    }

    private void RenderWatermark(PageGeometry pg, int pageCount)
    {
        var watermark = _viewModel?.Document.Watermark;
        PageViewBuilder.BuildWatermarkLayer(WatermarkCanvas, watermark, pg, pageCount);
    }

    // ── 오버레이 좌표 헬퍼 — 페이지-로컬 (AnchorPageIndex, X/Y mm) ↔ Canvas 절대 DIP ──

    /// <summary>오버레이 컨트롤을 (AnchorPageIndex, OverlayXMm/YMm) 가 가리키는 Canvas 절대 위치에 배치.</summary>
    private void PlaceOverlay(System.Windows.FrameworkElement ctrl, PolyDonky.Core.IOverlayAnchored anchor)
    {
        var pg = _pageGeometry;
        if (pg is null) return;
        var (xDip, yDip) = pg.ToAbsoluteDip(anchor.AnchorPageIndex, anchor.OverlayXMm, anchor.OverlayYMm);
        System.Windows.Controls.Canvas.SetLeft(ctrl, xDip);
        System.Windows.Controls.Canvas.SetTop (ctrl, yDip);
    }

    /// <summary>드래그 종료 후 Canvas 절대 좌표를 (AnchorPageIndex, OverlayXMm/YMm) 으로 분해해 모델에 반영.</summary>
    private void CommitOverlayDragPosition(PolyDonky.Core.IOverlayAnchored anchor, double leftDip, double topDip)
    {
        var pg = _pageGeometry;
        if (pg is null) return;
        var (pageIndex, xMm, yMm) = pg.ToPageLocal(leftDip, topDip);
        anchor.AnchorPageIndex = pageIndex;
        anchor.OverlayXMm      = xMm;
        anchor.OverlayYMm      = yMm;
    }

    /// <summary>현재 라이브 FlowDocument 의 모든 오버레이를 훑어 최대 AnchorPageIndex 반환.</summary>
    private int ComputeMaxAnchorPageIndex()
    {
        int max = 0;
        var section = _viewModel?.Document.Sections.FirstOrDefault();
        if (section is not null)
        {
            foreach (var b in section.Blocks)
            {
                if (b is PolyDonky.Core.IOverlayAnchored a)
                {
                    if (a.AnchorPageIndex > max) max = a.AnchorPageIndex;
                }
            }
        }
        // per-page 모드에서 오버레이 앵커는 _viewModel.Document 가 정본 — 두 번째 루프 불필요.
        return max;
    }

    /// <summary>페이지 배경 브러시 — 사용자 지정 색상 우선, 없으면 테마 Surface 리소스.</summary>
    private WpfMedia.Brush ResolvePaperBackground(PolyDonky.Core.PageSettings? page)
    {
        if (!string.IsNullOrEmpty(page?.PaperColor))
        {
            try
            {
                var c = (WpfMedia.Color)WpfMedia.ColorConverter.ConvertFromString(page.PaperColor)!;
                return new SolidColorBrush(c);
            }
            catch { /* 파싱 실패 시 테마 색상으로 폴백 */ }
        }
        return (WpfMedia.Brush)FindResource("Surface");
    }

    private static WpfMedia.Brush ParseColumnDividerBrush(string? hex)
    {
        if (!string.IsNullOrWhiteSpace(hex))
        {
            try
            {
                var s = hex.Trim();
                if (!s.StartsWith('#')) s = '#' + s;
                var c = (WpfMedia.Color)WpfMedia.ColorConverter.ConvertFromString(s)!;
                return new SolidColorBrush(c);
            }
            catch { /* 폴백 */ }
        }
        return new SolidColorBrush(WpfMedia.Color.FromArgb(0x66, 0x88, 0x88, 0x88));
    }

    private void OnEditorTextChanged(object sender, TextChangedEventArgs e)
    {
        if (_suppressTextChanged) return;
        _viewModel?.MarkDirty();
        // 본문이 변경되면 페이지 수가 달라질 수 있으므로 라이브 페이지네이션 갱신 디바운스.
        ScheduleLivePaginationRefresh();
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

    private void OnPreviewClick(object sender, System.Windows.RoutedEventArgs e)
    {
        if (_viewModel is null) return;
        if (!_viewModel.Document.IsPrintable)
        {
            System.Windows.MessageBox.Show(SR.DocInfoNotPrintable, SR.MenuFilePreviewPrint,
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        // per-page 모드: 각 RTB 에서 현재 본문을 파싱해 최신 스냅샷을 만든다.
        // _viewModel.GetPreviewDocument() 는 로드 시점 FlowDocument 를 쓰므로 편집 내용이 반영되지 않는다.
        var snapshot = PageEditorHost.PageCount > 0
            ? ParseAllPageEditors()
            : _viewModel.GetPreviewDocument();
        var wnd = new PrintPreviewWindow(snapshot) { Owner = this };
        wnd.Show();
    }

    private void OnDictMenuClick(object sender, System.Windows.RoutedEventArgs e)
    {
        var query = BodyEditor.Selection?.IsEmpty == false
            ? BodyEditor.Selection.Text.Trim()
            : null;

        if (_dictWindow is null)
            _dictWindow = new DictionaryWindow(query) { Owner = this };
        else if (!string.IsNullOrWhiteSpace(query))
            _dictWindow.SearchFor(query);

        if (!_dictWindow.IsVisible)
            _dictWindow.Show();
        else
            _dictWindow.Activate();
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
        // 글자 속성: 선택이 있으면 그 영역만, 없으면 caret 위치(이후 입력에 적용) — SelectAll 강제 안 함.
        var editor = GetActiveTextEditor();
        var dlg = new CharFormatWindow(editor) { Owner = this };
        if (dlg.ShowDialog() == true)
        {
            _viewModel?.MarkDirty();
            editor.Focus();
        }
    }

    private void OnFormatPara(object sender, RoutedEventArgs e)
    {
        // 문단 속성은 selection 이 비어있어도 caret 단락에만 적용 — EnsureInnerSelectionForDialog
        // (= SelectAll) 호출하지 않음. ParaFormatWindow.CollectParagraphs 가 caret 단락 1개 반환.
        var editor = GetActiveTextEditor();
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

    // ── 표 열 너비 드래그 리사이즈 ────────────────────────────────────────────

    private bool TryHitTableColumnBorder(Point pt,
        out System.Windows.Documents.Table? wpfTable,
        out PolyDonky.Core.Table? coreTable,
        out int leftColIdx,
        out double borderX)
    {
        wpfTable = null; coreTable = null; leftColIdx = -1; borderX = 0;

        var tp = BodyEditor.GetPositionFromPoint(pt, true);
        if (tp == null) return false;

        // Walk up to TableCell
        System.Windows.Documents.TableCell? cell = null;
        DependencyObject? cur = tp.Parent;
        while (cur is System.Windows.FrameworkContentElement fce)
        {
            if (cur is System.Windows.Documents.TableCell tc) { cell = tc; break; }
            cur = fce.Parent;
        }
        if (cell == null) return false;
        if (cell.ColumnSpan > 1) return false;

        var row      = cell.Parent as System.Windows.Documents.TableRow;
        var rowGroup = row?.Parent as System.Windows.Documents.TableRowGroup;
        var wTable   = rowGroup?.Parent as System.Windows.Documents.Table;
        if (wTable is null) return false;
        // 붙여넣기 표는 Tag 가 null 이므로 즉석 부착
        var cTable = EnsureCoreTable(wTable);

        int cellIdx  = row!.Cells.IndexOf(cell);
        int colCount = wTable.Columns.Count;
        if (cellIdx < 0) return false;

        var firstRow = rowGroup!.Rows[0];

        // 이 셀의 오른쪽 경계선 = 다음 셀 콘텐츠 시작 X
        if (cellIdx < colCount - 1 && cellIdx + 1 < firstRow.Cells.Count)
        {
            var nextRect = firstRow.Cells[cellIdx + 1].ContentStart
                                .GetCharacterRect(System.Windows.Documents.LogicalDirection.Forward);
            if (!nextRect.IsEmpty && Math.Abs(pt.X - nextRect.Left) <= TableColResizeHitDip)
            {
                wpfTable = wTable; coreTable = cTable;
                leftColIdx = cellIdx; borderX = nextRect.Left;
                return true;
            }
        }

        // 이 셀의 왼쪽 경계선 = 이 셀 콘텐츠 시작 X (이전 셀의 오른쪽 경계)
        if (cellIdx > 0 && cellIdx < firstRow.Cells.Count)
        {
            var thisRect = firstRow.Cells[cellIdx].ContentStart
                                .GetCharacterRect(System.Windows.Documents.LogicalDirection.Forward);
            if (!thisRect.IsEmpty && Math.Abs(pt.X - thisRect.Left) <= TableColResizeHitDip)
            {
                wpfTable = wTable; coreTable = cTable;
                leftColIdx = cellIdx - 1; borderX = thisRect.Left;
                return true;
            }
        }

        return false;
    }

    private void StartTableColumnResize(
        System.Windows.Documents.Table wpfTable,
        PolyDonky.Core.Table coreTable,
        int leftColIdx, double startX)
    {
        _tableColResizeActive = true;
        _colRszWpf     = wpfTable;
        _colRszCore    = coreTable;
        _colRszLeftIdx = leftColIdx;
        _colRszStartX  = startX;
        _colRszLeftCol  = wpfTable.Columns[leftColIdx];
        _colRszRightCol = leftColIdx + 1 < wpfTable.Columns.Count
                        ? wpfTable.Columns[leftColIdx + 1]
                        : null;

        // 현재 렌더된 열 너비 스냅샷 (셀 콘텐츠 시작 X 차이로 추정)
        var firstRow = wpfTable.RowGroups[0].Rows[0];
        double CellLeft(int idx)
        {
            if (idx >= firstRow.Cells.Count) return double.NaN;
            var r = firstRow.Cells[idx].ContentStart.GetCharacterRect(System.Windows.Documents.LogicalDirection.Forward);
            return r.IsEmpty ? double.NaN : r.Left;
        }

        double x0 = CellLeft(leftColIdx);
        double x1 = CellLeft(leftColIdx + 1);
        double x2 = CellLeft(leftColIdx + 2);

        _colRszInitLeft = !double.IsNaN(x0) && !double.IsNaN(x1) && x1 > x0
            ? x1 - x0
            : _colRszLeftCol.Width.IsAbsolute ? _colRszLeftCol.Width.Value : 80;

        _colRszInitRight = _colRszRightCol == null ? 0
            : !double.IsNaN(x1) && !double.IsNaN(x2) && x2 > x1
            ? x2 - x1
            : _colRszRightCol.Width.IsAbsolute ? _colRszRightCol.Width.Value : 80;

        // Auto 너비를 절대값으로 고정해야 드래그 중 다른 열이 튀지 않음
        _colRszLeftCol.Width = new GridLength(_colRszInitLeft);
        if (_colRszRightCol != null)
            _colRszRightCol.Width = new GridLength(_colRszInitRight);

        BodyEditor.CaptureMouse();
        Mouse.OverrideCursor = Cursors.SizeWE;
    }

    private void FinishTableColumnResize()
    {
        if (!_tableColResizeActive) return;
        _tableColResizeActive   = false;
        _tableColResizeHovering = false;
        if (BodyEditor.IsMouseCaptured) BodyEditor.ReleaseMouseCapture();
        Mouse.OverrideCursor = null;

        if (_colRszCore != null && _colRszLeftCol != null)
        {
            int li = _colRszLeftIdx;
            int ri = li + 1;
            double leftDip  = _colRszLeftCol.Width.IsAbsolute  ? _colRszLeftCol.Width.Value  : 0;
            double rightDip = _colRszRightCol?.Width.IsAbsolute == true ? _colRszRightCol.Width.Value : 0;
            if (li < _colRszCore.Columns.Count)
                _colRszCore.Columns[li].WidthMm = Services.FlowDocumentBuilder.DipToMm(leftDip);
            if (ri < _colRszCore.Columns.Count && _colRszRightCol != null)
                _colRszCore.Columns[ri].WidthMm = Services.FlowDocumentBuilder.DipToMm(rightDip);
        }

        _colRszWpf      = null;
        _colRszCore     = null;
        _colRszLeftCol  = null;
        _colRszRightCol = null;
        _viewModel?.MarkDirty();
    }

    // ── 단 너비 드래그 리사이즈 헬퍼 ────────────────────────────────────────

    /// <summary>
    /// PaperHost 기준 좌표 <paramref name="ptInPaper"/> 가 단 구분선 Hit 영역 안에 있으면 true.
    /// <paramref name="leftColIdx"/> 는 경계 왼쪽 단 인덱스 (0-based).
    /// </summary>
    private bool TryHitColumnDivider(Point ptInPaper, out int leftColIdx)
    {
        leftColIdx = -1;
        var pg = _pageGeometry;
        if (pg is null || pg.ColumnCount <= 1) return false;

        // Y 가 어느 한 페이지의 본문 영역 안이어야 한다 (페이지 사이 갭·머리글/바닥글 영역은 제외).
        double stride    = pg.PageStrideDip;
        double yInPage   = stride > 0 ? ptInPaper.Y % stride : ptInPaper.Y;
        if (yInPage < pg.PadTopDip || yInPage > pg.PageHeightDip - pg.PadBottomDip)
            return false;

        for (int c = 0; c < pg.ColumnCount - 1; c++)
        {
            double colW = c < pg.ColWidthsDip.Length ? pg.ColWidthsDip[c] : pg.ColWidthDip;
            double divX = pg.PadLeftDip + pg.ColumnXOffsetDip(c) + colW + pg.ColGapDip / 2.0;
            if (Math.Abs(ptInPaper.X - divX) <= ColDivHitDip)
            {
                leftColIdx = c;
                return true;
            }
        }
        return false;
    }

    /// <summary>단 너비 드래그 중 마우스 이동 — 구분선 위치를 실시간으로 업데이트한다.</summary>
    private void HandleColDivDragMove(MouseEventArgs e)
    {
        var pg = _pageGeometry;
        if (pg is null) return;

        var pos   = e.GetPosition(PaperHost);
        double delta = pos.X - _colDivDragStartX;
        int l = _colDivDragLeftIdx, r = l + 1;
        if (r >= _colDivDragStartWidths.Length) return;

        double total = _colDivDragStartWidths[l] + _colDivDragStartWidths[r];
        double newL  = Math.Clamp(_colDivDragStartWidths[l] + delta, ColDivMinWidthDip, total - ColDivMinWidthDip);

        // 드래그된 경계의 새 X = padLeft + (변경되지 않은 이전 단들의 X 합) + newL + gap/2
        double newDivX = pg.PadLeftDip + pg.ColumnXOffsetDip(l) + newL + pg.ColGapDip / 2.0;

        foreach (var line in _columnDividerLines)
        {
            if (line.Tag is int idx && idx == l)
            {
                line.X1 = newDivX;
                line.X2 = newDivX;
            }
        }
    }

    /// <summary>단 너비 드래그 완료 — 모델 갱신 후 페이지 재구성.</summary>
    private void FinishColDivDrag(MouseButtonEventArgs e)
    {
        _colDivDragging = false;
        _colDivHovering = false;
        if (PaperHost.IsMouseCaptured) PaperHost.ReleaseMouseCapture();
        Mouse.OverrideCursor = null;

        var pos   = e.GetPosition(PaperHost);
        double delta = pos.X - _colDivDragStartX;
        int l = _colDivDragLeftIdx, r = l + 1;
        if (r >= _colDivDragStartWidths.Length || _pageGeometry is null) { e.Handled = true; return; }

        double total = _colDivDragStartWidths[l] + _colDivDragStartWidths[r];
        double newL  = Math.Clamp(_colDivDragStartWidths[l] + delta, ColDivMinWidthDip, total - ColDivMinWidthDip);
        double newR  = total - newL;

        var sec = _viewModel?.Document.Sections.FirstOrDefault();
        if (sec is null) { e.Handled = true; return; }

        var newWidths = (double[])_pageGeometry.ColWidthsDip.Clone();
        newWidths[l] = newL;
        newWidths[r] = newR;
        sec.Page.ColumnWidthsMm = newWidths.Select(w => Services.FlowDocumentBuilder.DipToMm(w)).ToList();

        ApplyPageSettings(sec.Page);
        _viewModel?.MarkDirty();
        e.Handled = true;
    }

    // ── 단 경계 교차 텍스트 선택 헬퍼 ────────────────────────────────────────

    /// <summary>
    /// RTB 포커스 변경 시 — 비-Shift 원인(클릭 등)이면 단 교차 선택을 모두 해제한다.
    /// _inCrossSelNavigation 플래그가 설정된 프로그래매틱 이동은 무시.
    /// </summary>
    private void OnPageRtbGotFocusClearCrossSel(object sender, KeyboardFocusChangedEventArgs e)
    {
        if (_inCrossSelNavigation) return;
        if ((Keyboard.Modifiers & ModifierKeys.Shift) != 0) return;
        ClearCrossColumnSelection();
    }

    /// <summary>
    /// 모든 비활성 RTB 의 선택을 해제하고 교차 선택 상태를 초기화한다.
    /// </summary>
    private void ClearCrossColumnSelection()
    {
        if (!_crossSelActive) return;
        _crossSelActive = false;
        foreach (var rtb in PageEditorHost.PageEditors)
        {
            if (rtb.IsKeyboardFocusWithin) continue;
            if (!rtb.Selection.IsEmpty)
            {
                var cp = rtb.CaretPosition ?? rtb.Document.ContentStart;
                rtb.Selection.Select(cp, cp);
            }
        }
    }

    /// <summary>
    /// Shift+방향키가 RTB 경계에 도달했을 때 현재 RTB 선택을 경계까지 확장하고
    /// 인접 RTB 로 캐럿을 이동한다.
    /// </summary>
    private bool TryExtendCrossColumnSelection(
        RichTextBox rtb, WpfDocs.TextPointer caret, int idx, KeyEventArgs e)
    {
        // 방향 판단: 경계 조건이 아닌 키는 WPF 기본 처리에 양보
        bool? fwd = e.Key switch
        {
            Key.Right when caret.GetNextInsertionPosition(WpfDocs.LogicalDirection.Forward)  is null => true,
            Key.Down  when caret.GetLineStartPosition(1)   is null => true,
            Key.Left  when caret.GetNextInsertionPosition(WpfDocs.LogicalDirection.Backward) is null => false,
            Key.Up    when caret.GetLineStartPosition(-1)  is null => false,
            _ => (bool?)null,
        };
        if (fwd is null) return false;

        var pages = PageEditorHost.PageEditors;
        int targetIdx = fwd.Value ? idx + 1 : idx - 1;
        if (targetIdx < 0 || targetIdx >= pages.Count) return false;

        // 현재 RTB 의 선택 고정점(anchor):
        //   선택 없음 → 캐럿 그대로
        //   앞으로 확장 중 → 선택 시작(Start)이 고정점
        //   뒤로 확장 중  → 선택 끝(End)이 고정점
        var sel    = rtb.Selection;
        var anchor = sel.IsEmpty ? caret
            : fwd.Value ? sel.Start : sel.End;

        // 현재 RTB 선택을 경계까지 확장
        if (fwd.Value)
        {
            var endPos = rtb.Document.ContentEnd
                .GetInsertionPosition(WpfDocs.LogicalDirection.Backward) ?? rtb.Document.ContentEnd;
            rtb.Selection.Select(anchor, endPos);
        }
        else
        {
            var startPos = rtb.Document.ContentStart
                .GetInsertionPosition(WpfDocs.LogicalDirection.Forward) ?? rtb.Document.ContentStart;
            rtb.Selection.Select(startPos, anchor);
        }

        _crossSelActive = true;

        // 인접 RTB 로 캐럿 이동 (선택 없이, GotKeyboardFocus 는 무시)
        _inCrossSelNavigation = true;
        try
        {
            var target    = pages[targetIdx];
            target.Focus();
            var newCaret  = fwd.Value
                ? target.Document.ContentStart
                    .GetInsertionPosition(WpfDocs.LogicalDirection.Forward) ?? target.Document.ContentStart
                : target.Document.ContentEnd
                    .GetInsertionPosition(WpfDocs.LogicalDirection.Backward) ?? target.Document.ContentEnd;
            target.CaretPosition = newCaret;
        }
        finally
        {
            _inCrossSelNavigation = false;
        }

        e.Handled = true;
        return true;
    }

    // ── BodyEditor 우클릭 통합 핸들러 ──────────────────────────────────────
    //
    // BodyEditor 본문 영역의 우클릭 메뉴는 모두 이 한 곳에서 빌드한다.
    // (오버레이/언더레이 캔버스 객체는 OnPaperPreviewMouseRightButtonDown 이
    //  e.Handled = true 로 가로채므로 여기에 도달하지 않는다.)
    //
    // 매번 메뉴를 비우고 처음부터 다시 빌드 → 같은 선택 상태에서는 항상 같은 메뉴.
    private void OnEmbeddedObjectContextMenuOpening(object sender, System.Windows.Controls.ContextMenuEventArgs e)
    {
        var menu = BodyEditor.ContextMenu;
        if (menu is null) return;

        menu.Items.Clear();

        // ① 기본 클립보드 메뉴 (모든 컨텍스트 공통)
        menu.Items.Add(new System.Windows.Controls.MenuItem
        {
            Header = "잘라내기(_T)", Command = ApplicationCommands.Cut,   InputGestureText = "Ctrl+X"
        });
        menu.Items.Add(new System.Windows.Controls.MenuItem
        {
            Header = "복사(_C)",     Command = ApplicationCommands.Copy,  InputGestureText = "Ctrl+C"
        });
        menu.Items.Add(new System.Windows.Controls.MenuItem
        {
            Header = "붙여넣기(_P)", Command = ApplicationCommands.Paste, InputGestureText = "Ctrl+V"
        });

        // ① 서식 메뉴 — 텍스트 컨텍스트 공통
        menu.Items.Add(new System.Windows.Controls.Separator());
        var miFormatChar = new System.Windows.Controls.MenuItem
        {
            Header = "글자 서식(_L)...", InputGestureText = "Ctrl+L"
        };
        miFormatChar.Click += OnFormatChar;
        menu.Items.Add(miFormatChar);

        var miFormatPara = new System.Windows.Controls.MenuItem
        {
            Header = "문단 서식(_T)...", InputGestureText = "Ctrl+T"
        };
        miFormatPara.Click += OnFormatPara;
        menu.Items.Add(miFormatPara);

        // ② 표 컨텍스트 — 멀티 셀 선택이 우선, 없으면 캐럿 위치 셀
        if (FindSelectedTableCells() is { Count: > 1 } multiCells)
        {
            AppendMultiCellMenuItems(menu, multiCells);
        }
        else if (FindTableCellAtCaret(out var wpfTable, out var wpfRow, out var wpfCell) &&
                 wpfTable is not null && wpfCell is not null)
        {
            // 클립보드에서 붙여넣기 된 표는 Tag 가 비어있을 수 있어 즉석에서 Core.Table 부착
            var coreTable = EnsureCoreTable(wpfTable);
            AppendTablePropertyMenuItems(menu, wpfTable, wpfRow, wpfCell, coreTable);
        }

        // ③ 인라인 이미지/이모지 — 속성 항목
        var pt = System.Windows.Input.Mouse.GetPosition(BodyEditor);
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

    // ── 멀티 셀 선택 감지 ──────────────────────────────────────────────────

    private record SelectedCell(
        System.Windows.Documents.Table WpfTable,
        System.Windows.Documents.TableRowGroup RowGroup,
        System.Windows.Documents.TableRow WpfRow,
        System.Windows.Documents.TableCell WpfCell,
        PolyDonky.Core.Table CoreTable,
        PolyDonky.Core.TableCell CoreCell,
        int RowIdx, int CellIdx);

    /// <summary>TextPointer 의 조상 체인에서 가장 가까운 TableCell 을 찾는다 (List·Span 등 중간 컨테이너 통과).</summary>
    private static System.Windows.Documents.TableCell? FindAncestorCell(System.Windows.Documents.TextPointer? tp)
    {
        if (tp == null) return null;
        DependencyObject? cur = tp.Parent;
        while (cur is System.Windows.FrameworkContentElement fce)
        {
            if (cur is System.Windows.Documents.TableCell tc) return tc;
            cur = fce.Parent;
        }
        return null;
    }

    /// <summary>
    /// WPF Table 의 Tag 에 PolyDonky.Core.Table 이 비어있으면(예: 클립보드 붙여넣기로 새로 만들어진 표),
    /// WPF 구조를 미러링한 신선한 Core.Table 을 만들어 부착해 반환한다. 이 헬퍼는 표 우클릭 메뉴·열
    /// 리사이즈·멀티셀 메뉴 등 모든 Tag 의존 코드 경로에서 호출되어 라운드트립 안전성을 보장한다.
    /// (저장 경로의 FlowDocumentParser.ParseTable 도 Tag 가 없을 때 동일하게 처리하므로 모순 없음.)
    /// </summary>
    private static PolyDonky.Core.Table EnsureCoreTable(System.Windows.Documents.Table wpfTable)
    {
        if (wpfTable.Tag is PolyDonky.Core.Table existing) return existing;

        var coreTable = new PolyDonky.Core.Table();

        // 컬럼: WPF GridLength 가 절대값이면 mm 로 환산, 아니면 Auto(0).
        foreach (var col in wpfTable.Columns)
        {
            double widthMm = 0;
            if (col.Width.IsAbsolute && col.Width.Value > 0)
                widthMm = PolyDonky.App.Services.FlowDocumentBuilder.DipToMm(col.Width.Value);
            coreTable.Columns.Add(new PolyDonky.Core.TableColumn { WidthMm = widthMm });
        }

        // 행/셀 구조 미러링 (셀 본문은 빈 Paragraph 하나로 시드 — 저장 시 Parser 가 실제 본문을 채움)
        var rowGroup = wpfTable.RowGroups.FirstOrDefault();
        if (rowGroup != null)
        {
            foreach (var wpfRow in rowGroup.Rows)
            {
                var coreRow = new PolyDonky.Core.TableRow();
                foreach (var wpfCell in wpfRow.Cells)
                {
                    var coreCell = new PolyDonky.Core.TableCell
                    {
                        ColumnSpan = wpfCell.ColumnSpan > 0 ? wpfCell.ColumnSpan : 1,
                        RowSpan    = wpfCell.RowSpan    > 0 ? wpfCell.RowSpan    : 1,
                    };
                    coreCell.Blocks.Add(PolyDonky.Core.Paragraph.Of(string.Empty));
                    coreRow.Cells.Add(coreCell);
                }
                coreTable.Rows.Add(coreRow);
            }
        }

        wpfTable.Tag = coreTable;
        return coreTable;
    }

    private List<SelectedCell>? FindSelectedTableCells()
    {
        if (BodyEditor.Selection.IsEmpty) return null;

        // Paragraph.Parent 직접 검사 대신 조상 체인을 끝까지 거슬러 올라간다.
        // (List · Span · 중첩 Block 등으로 중간 단계가 있을 수 있음)
        var startWpfCell = FindAncestorCell(BodyEditor.Selection.Start);

        // Selection.End 가 셀 오른쪽 경계 또는 표 뒤의 Paragraph에 위치하면
        // FindAncestorCell 이 null 을 반환하거나 다음 셀을 반환해 off-by-one 발생.
        // 한 위치 뒤로 물러나면 항상 마지막으로 선택된 셀 안쪽에 머문다.
        var endRaw = BodyEditor.Selection.End;
        var endPos = endRaw.GetPositionAtOffset(-1, System.Windows.Documents.LogicalDirection.Backward) ?? endRaw;
        var endWpfCell = FindAncestorCell(endPos);

        if (startWpfCell == null || endWpfCell == null) return null;
        if (ReferenceEquals(startWpfCell, endWpfCell)) return null; // 단일 셀

        var startRow  = startWpfCell.Parent as System.Windows.Documents.TableRow;
        var rowGroup  = startRow?.Parent as System.Windows.Documents.TableRowGroup;
        var wpfTable  = rowGroup?.Parent as System.Windows.Documents.Table;
        if (wpfTable is null || rowGroup is null || startRow is null) return null;
        // 붙여넣기 표는 Tag 가 null 일 수 있으므로 즉석 부착
        var coreTable = EnsureCoreTable(wpfTable);

        // 끝 셀이 같은 표에 속하는지 확인
        var endRow = endWpfCell.Parent as System.Windows.Documents.TableRow;
        var endRowGroup = endRow?.Parent as System.Windows.Documents.TableRowGroup;
        if (!ReferenceEquals(endRowGroup, rowGroup)) return null;

        int startRowIdx  = rowGroup.Rows.IndexOf(startRow);
        int startCellIdx = startRow.Cells.IndexOf(startWpfCell);
        int endRowIdx    = rowGroup.Rows.IndexOf(endRow!);
        int endCellIdx   = endRow!.Cells.IndexOf(endWpfCell);

        int minRow  = Math.Min(startRowIdx, endRowIdx);
        int maxRow  = Math.Max(startRowIdx, endRowIdx);
        int minCell = Math.Min(startCellIdx, endCellIdx);
        int maxCell = Math.Max(startCellIdx, endCellIdx);

        var result = new List<SelectedCell>();
        for (int r = minRow; r <= maxRow; r++)
        {
            if (r >= rowGroup.Rows.Count || r >= coreTable.Rows.Count) break;
            var wRow  = rowGroup.Rows[r];
            var coRow = coreTable.Rows[r];
            for (int c = minCell; c <= maxCell; c++)
            {
                if (c >= wRow.Cells.Count || c >= coRow.Cells.Count) break;
                result.Add(new SelectedCell(wpfTable, rowGroup, wRow, wRow.Cells[c],
                                            coreTable, coRow.Cells[c], r, c));
            }
        }
        return result.Count > 1 ? result : null;
    }

    // ── 멀티 셀 컨텍스트 메뉴 ─────────────────────────────────────────────

    private void AppendMultiCellMenuItems(
        System.Windows.Controls.ContextMenu menu,
        List<SelectedCell> cells)
    {
        var items = new List<System.Windows.Controls.Control>();
        var first = cells[0];

        items.Add(new System.Windows.Controls.Separator());

        // 선택 셀 병합
        bool canMerge = cells.All(c => c.CoreCell.ColumnSpan == 1 && c.CoreCell.RowSpan == 1);
        var mergeItem = MakeMenuItem($"선택 셀 {cells.Count}개 병합(_M)",
            () => TableOp_MergeSelectedCells(cells));
        mergeItem.IsEnabled = canMerge;
        items.Add(mergeItem);

        // 선택 셀 내용 지우기
        items.Add(MakeMenuItem("선택 셀 내용 지우기(_E)", () =>
        {
            foreach (var sc in cells)
            {
                sc.WpfCell.Blocks.Clear();
                sc.WpfCell.Blocks.Add(new System.Windows.Documents.Paragraph(
                    new System.Windows.Documents.Run(string.Empty)));
                sc.CoreCell.Blocks.Clear();
                sc.CoreCell.Blocks.Add(PolyDonky.Core.Paragraph.Of(string.Empty));
            }
            _viewModel?.MarkDirty();
        }));

        // 선택 셀 배경색
        items.Add(MakeMenuItem("선택 셀 배경색(_C)...", () =>
        {
            var dlg = new System.Windows.Forms.ColorDialog { FullOpen = true, AnyColor = true };
            if (dlg.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;
            var c    = dlg.Color;
            string hex = $"#{c.R:X2}{c.G:X2}{c.B:X2}";
            var brush = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(c.R, c.G, c.B));
            foreach (var sc in cells)
            {
                sc.CoreCell.BackgroundColor = hex;
                sc.WpfCell.Background = brush;
            }
            _viewModel?.MarkDirty();
        }));

        // 선택 셀 속성 (첫 번째 셀 기준으로 다이얼로그 열고, OK 시 전체 적용)
        items.Add(MakeMenuItem($"선택 셀 속성(_P) [{cells.Count}개]...", () =>
        {
            var dlg = new CellPropertiesWindow(first.CoreCell) { Owner = this };
            if (dlg.ShowDialog() != true) return;
            bool isHdr = first.CoreTable.Rows[first.RowIdx].IsHeader;
            // 첫 셀의 변경값을 나머지 셀에 복사
            foreach (var sc in cells)
            {
                sc.CoreCell.TextAlign         = first.CoreCell.TextAlign;
                sc.CoreCell.PaddingTopMm      = first.CoreCell.PaddingTopMm;
                sc.CoreCell.PaddingBottomMm   = first.CoreCell.PaddingBottomMm;
                sc.CoreCell.PaddingLeftMm     = first.CoreCell.PaddingLeftMm;
                sc.CoreCell.PaddingRightMm    = first.CoreCell.PaddingRightMm;
                sc.CoreCell.BorderThicknessPt = first.CoreCell.BorderThicknessPt;
                sc.CoreCell.BorderColor       = first.CoreCell.BorderColor;
                sc.CoreCell.BackgroundColor   = first.CoreCell.BackgroundColor;
                bool cellIsHdr = first.CoreTable.Rows[sc.RowIdx].IsHeader;
                PolyDonky.App.Services.FlowDocumentBuilder.ApplyCellPropertiesToWpf(
                    sc.WpfCell, sc.CoreCell, cellIsHdr, sc.CoreTable);
            }
            _viewModel?.MarkDirty();
        }));

        // 표 속성 (표 전체)
        items.Add(new System.Windows.Controls.Separator());
        items.Add(MakeMenuItem("표 속성(_T)...", () =>
        {
            var dlg = new TablePropertiesWindow(first.CoreTable) { Owner = this };
            if (dlg.ShowDialog() == true)
            {
                PolyDonky.App.Services.FlowDocumentBuilder.ApplyTableLevelPropertiesToWpf(
                    first.WpfTable, first.CoreTable);
                _viewModel?.MarkDirty();
            }
        }));

        foreach (var item in items) menu.Items.Add(item);
        void Cleanup(object? s, System.Windows.RoutedEventArgs _ev)
        {
            menu.Closed -= Cleanup;
            foreach (var ctrl in items) menu.Items.Remove(ctrl);
        }
        menu.Closed += Cleanup;
    }

    // ── 선택 셀 병합 ──────────────────────────────────────────────────────

    private void TableOp_MergeSelectedCells(List<SelectedCell> cells)
    {
        if (cells.Count < 2) return;
        var coreTable = cells[0].CoreTable;
        var rowGroup  = cells[0].RowGroup;

        int minRow  = cells.Min(c => c.RowIdx);
        int maxRow  = cells.Max(c => c.RowIdx);
        int minCell = cells.Min(c => c.CellIdx);
        int maxCell = cells.Max(c => c.CellIdx);
        int newColSpan = maxCell - minCell + 1;
        int newRowSpan = maxRow  - minRow  + 1;

        // 첫 셀에 병합된 내용 집결
        var firstWpf  = cells[0].WpfCell;
        var firstCore = cells[0].CoreCell;

        for (int i = 1; i < cells.Count; i++)
        {
            var sc = cells[i];
            foreach (var b in sc.CoreCell.Blocks)
                firstCore.Blocks.Add(b);
            foreach (System.Windows.Documents.Block b in sc.WpfCell.Blocks.ToList())
            {
                sc.WpfCell.Blocks.Remove(b);
                firstWpf.Blocks.Add(b);
            }
        }

        firstWpf.ColumnSpan  = newColSpan;
        firstWpf.RowSpan     = newRowSpan;
        firstCore.ColumnSpan = newColSpan;
        firstCore.RowSpan    = newRowSpan;

        // 첫 셀 제외한 나머지 제거 (역순으로)
        foreach (var sc in cells.Skip(1).Reverse())
        {
            sc.WpfRow.Cells.Remove(sc.WpfCell);
            coreTable.Rows[sc.RowIdx].Cells.Remove(sc.CoreCell);
        }

        _viewModel?.MarkDirty();
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
            if (dlg.ShowDialog() == true)
            {
                if (coreTable.WrapMode == PolyDonky.Core.TableWrapMode.Block)
                {
                    // Block 모드 유지 → WPF Table 에 직접 적용
                    PolyDonky.App.Services.FlowDocumentBuilder.ApplyTableLevelPropertiesToWpf(wpfTable, coreTable);
                }
                else
                {
                    // 오버레이 모드로 전환 — FlowDocument 내 WPF Table 을 앵커 단락으로 교체
                    var anchor = new System.Windows.Documents.Paragraph
                    {
                        Tag        = coreTable,
                        Margin     = new Thickness(0),
                        FontSize   = 0.1,
                        Foreground = System.Windows.Media.Brushes.Transparent,
                        Background = System.Windows.Media.Brushes.Transparent,
                    };
                    BodyEditor.Document.Blocks.InsertBefore(wpfTable, anchor);
                    BodyEditor.Document.Blocks.Remove(wpfTable);
                    RebuildOverlayTables();
                }
                _viewModel?.MarkDirty();
            }
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
            // PaperHost 기준 좌표 = OverlayImageCanvas/UnderlayImageCanvas 좌표.
            var prevMode = imageBlock.WrapMode;
            Point currentPos;
            try { currentPos = imgControl.TransformToVisual(PaperHost).Transform(new Point(0, 0)); }
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
                if (toOverlay && imageBlock.AnchorPageIndex == 0
                    && imageBlock.OverlayXMm == 0 && imageBlock.OverlayYMm == 0)
                {
                    CommitOverlayDragPosition(imageBlock, currentPos.X, currentPos.Y);
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
        // 텍스트 입력은 활성 RTB 에만 적용 — 비활성 RTB 의 교차 선택을 해제한다.
        if (_crossSelActive) ClearCrossColumnSelection();

        if (_viewModel?.IsWriteProtected != true) return;
        e.Handled = true;
        _viewModel.TryUnlockForEditing();
    }

    private void OnEditorPreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (sender is not RichTextBox rtb) return;

        // 쓰기 보호 모드 — 편집 의도 키를 잠금 해제 트리거로 사용
        if (_viewModel?.IsWriteProtected == true && IsEditingIntent(e))
        {
            e.Handled = true;
            _viewModel.TryUnlockForEditing();
            return;
        }

        // Shift 없는 탐색키 → 단 교차 선택 해제 (경계 내 이동 포함).
        // TryHandlePageBoundaryNavigation 의 비-Shift 경계 분기도 ClearCrossColumnSelection
        // 을 직접 호출하므로 중복 호출이 되지만 이중 호출은 무해하다.
        if (_crossSelActive && (Keyboard.Modifiers & ModifierKeys.Shift) == 0)
        {
            if (e.Key is Key.Left or Key.Right or Key.Up or Key.Down
                      or Key.Home or Key.End or Key.PageDown or Key.PageUp)
                ClearCrossColumnSelection();
        }

        // per-page RTB 모델에서 페이지 경계를 넘는 캐럿 이동을 직접 처리.
        TryHandlePageBoundaryNavigation(rtb, e);
    }

    /// <summary>
    /// per-page 편집기에서 RTB 간 캐럿/포커스 이동.
    /// PgUp/PgDn: 항상 인접 페이지로 이동.
    /// 위/아래 화살표: 현재 RTB 의 첫/마지막 줄에서만 인접 페이지로 이동.
    /// Ctrl+Home/End: 첫/마지막 페이지로 이동.
    /// </summary>
    private bool TryHandlePageBoundaryNavigation(RichTextBox rtb, KeyEventArgs e)
    {
        var pages = PageEditorHost.PageEditors;
        int idx = -1;
        for (int i = 0; i < pages.Count; i++)
        {
            if (ReferenceEquals(pages[i], rtb)) { idx = i; break; }
        }
        if (idx < 0) return false;

        var caret = rtb.CaretPosition;
        if (caret is null) return false;

        var mods    = Keyboard.Modifiers;
        bool shift  = (mods & ModifierKeys.Shift)   == ModifierKeys.Shift;
        bool ctrl   = (mods & ModifierKeys.Control) == ModifierKeys.Control;

        // Shift+방향키: 단 경계에서 선택 확장 (경계가 아닌 경우 WPF 기본 동작에 양보)
        if (shift)
            return TryExtendCrossColumnSelection(rtb, caret, idx, e);

        // 비-Shift 경계 이동 시 기존 단 교차 선택을 해제한다.
        ClearCrossColumnSelection();

        // 다단 모드에서 한 페이지가 여러 RTB(=단)로 쪼개져 있으므로 PgUp/PgDn 은
        // 인접 단이 아니라 인접 *물리 페이지* 의 첫 단으로 점프해야 한다.
        int colCount  = Math.Max(1, _pageGeometry?.ColumnCount ?? 1);
        int curPage   = idx / colCount;

        switch (e.Key)
        {
            case Key.PageDown:
                return MoveCaretToPage((curPage + 1) * colCount, atTop: true, e);

            case Key.PageUp:
                return MoveCaretToPage((curPage - 1) * colCount, atTop: false, e);

            case Key.Down when caret.GetLineStartPosition(1) is null:
                return MoveCaretToPage(idx + 1, atTop: true, e);

            case Key.Up when caret.GetLineStartPosition(-1) is null:
                return MoveCaretToPage(idx - 1, atTop: false, e);

            // 좌우 화살표: 현재 RTB 의 첫/마지막 삽입 위치에서 인접 RTB(다음/이전 단 또는 페이지)로 이동.
            // 다단 모드에서 단 사이 캐럿 이동을 처리한다 — 단 1 끝에서 →, 단 2 시작에서 ←.
            case Key.Right when caret.GetNextInsertionPosition(WpfDocs.LogicalDirection.Forward) is null:
                return MoveCaretToPage(idx + 1, atTop: true, e);

            case Key.Left when caret.GetNextInsertionPosition(WpfDocs.LogicalDirection.Backward) is null:
                return MoveCaretToPage(idx - 1, atTop: false, e);

            case Key.Home when ctrl:
                return MoveCaretToPage(0, atTop: true, e);

            case Key.End when ctrl:
                return MoveCaretToPage(pages.Count - 1, atTop: false, e);

            case Key.Delete when !ctrl:
                return TryDeleteAcrossPageBoundary(idx, rtb, e);

            case Key.Back when !ctrl:
                return TryBackspaceAcrossPageBoundary(idx, rtb, e);
        }
        return false;
    }

    /// <summary>
    /// Del 키: 현재 RTB 의 마지막 삽입 위치에서 다음 페이지 RTB 의 첫 블록을 끌어올린다.
    /// 빈 공간 하단에서 Del 을 눌러도 다음 페이지 내용이 올라오도록 처리.
    /// </summary>
    private bool TryDeleteAcrossPageBoundary(int pageIdx, RichTextBox rtb, KeyEventArgs e)
    {
        var pages = PageEditorHost.PageEditors;
        if (pageIdx + 1 >= pages.Count) return false;

        // 현재 RTB 에 아직 삭제할 내용이 남아 있으면 RTB 의 기본 Delete 동작에 양보.
        var nextInsert = rtb.CaretPosition
            .GetNextInsertionPosition(WpfDocs.LogicalDirection.Forward);
        if (nextInsert is not null) return false;

        var nextRtb   = pages[pageIdx + 1];
        var nextDoc   = nextRtb.Document;
        var firstBlock = nextDoc.Blocks.FirstBlock;
        if (firstBlock is null) return false;

        var curDoc    = rtb.Document;
        var lastBlock = curDoc.Blocks.LastBlock;

        if (lastBlock is WpfDocs.Paragraph lastPara && firstBlock is WpfDocs.Paragraph firstPara)
        {
            // 단락 병합: 다음 페이지 첫 단락의 인라인을 현재 페이지 마지막 단락으로 이동
            var inlines = firstPara.Inlines.ToList();
            firstPara.Inlines.Clear();
            nextDoc.Blocks.Remove(firstPara);
            foreach (var il in inlines)
                lastPara.Inlines.Add(il);
        }
        else
        {
            // 단락 종류가 다르거나 블록 단위 이동 (표·목록 등)
            nextDoc.Blocks.Remove(firstBlock);
            curDoc.Blocks.Add(firstBlock);
        }

        // 다음 페이지 RTB 가 비면 빈 단락 하나 유지 (RichTextBox 최소 요건)
        if (nextDoc.Blocks.Count == 0)
            nextDoc.Blocks.Add(new WpfDocs.Paragraph());

        e.Handled = true;
        ScheduleLivePaginationRefresh();
        return true;
    }

    /// <summary>
    /// Backspace 키: 현재 RTB 의 첫 삽입 위치에서 이전 페이지 RTB 의 마지막 블록에 병합한다.
    /// 페이지 상단에서 Backspace 를 눌러도 이전 페이지 내용과 연결되도록 처리.
    /// </summary>
    private bool TryBackspaceAcrossPageBoundary(int pageIdx, RichTextBox rtb, KeyEventArgs e)
    {
        var pages = PageEditorHost.PageEditors;
        if (pageIdx <= 0) return false;

        // 현재 RTB 에 아직 삭제할 내용이 남아 있으면 RTB 의 기본 Backspace 동작에 양보.
        var prevInsert = rtb.CaretPosition
            .GetNextInsertionPosition(WpfDocs.LogicalDirection.Backward);
        if (prevInsert is not null) return false;

        var prevRtb    = pages[pageIdx - 1];
        var prevDoc    = prevRtb.Document;
        var lastBlock  = prevDoc.Blocks.LastBlock;
        if (lastBlock is null) return false;

        var curDoc     = rtb.Document;
        var firstBlock = curDoc.Blocks.FirstBlock;

        if (lastBlock is WpfDocs.Paragraph lastPara && firstBlock is WpfDocs.Paragraph firstPara)
        {
            // 단락 병합: 현재 페이지 첫 단락의 인라인을 이전 페이지 마지막 단락으로 이동
            var inlines = firstPara.Inlines.ToList();
            firstPara.Inlines.Clear();
            curDoc.Blocks.Remove(firstPara);
            foreach (var il in inlines)
                lastPara.Inlines.Add(il);
        }
        else
        {
            curDoc.Blocks.Remove(firstBlock);
            prevDoc.Blocks.Add(firstBlock);
        }

        // 현재 페이지 RTB 가 비면 빈 단락 하나 유지
        if (curDoc.Blocks.Count == 0)
            curDoc.Blocks.Add(new WpfDocs.Paragraph());

        // 이전 페이지 RTB 로 포커스 이동, 캐럿을 끝에 배치
        var anchor = prevRtb.Document.ContentEnd
            .GetInsertionPosition(WpfDocs.LogicalDirection.Backward)
            ?? prevRtb.Document.ContentEnd;
        prevRtb.CaretPosition = anchor;
        prevRtb.Focus();

        e.Handled = true;
        ScheduleLivePaginationRefresh();
        return true;
    }

    /// <summary>지정 인덱스 페이지 RTB 로 포커스 이동, 캐럿을 시작/끝에 배치.</summary>
    private bool MoveCaretToPage(int targetIdx, bool atTop, KeyEventArgs e)
    {
        var pages = PageEditorHost.PageEditors;
        if (targetIdx < 0 || targetIdx >= pages.Count) return false;

        var target = pages[targetIdx];
        target.Focus();
        var anchor = atTop
            ? target.Document.ContentStart.GetInsertionPosition(WpfDocs.LogicalDirection.Forward)
              ?? target.Document.ContentStart
            : target.Document.ContentEnd.GetInsertionPosition(WpfDocs.LogicalDirection.Backward)
              ?? target.Document.ContentEnd;
        target.CaretPosition = anchor;
        e.Handled = true;
        return true;
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
        double paperW = PaperHost.Width;
        double viewW  = EditorScrollViewer.ViewportWidth;
        if (paperW <= 0 || viewW <= 0) return;
        _viewModel.ZoomPercent = (viewW - hMargin) / paperW * 100;
    }

    private void OnFitToPage(object sender, RoutedEventArgs e)
    {
        if (_viewModel is null) return;
        const double hMargin = 64; // StackPanel 좌우 여백
        const double vMargin = 76; // StackPanel 상(28)+하(48) 여백
        double paperW = PaperHost.Width;
        double paperH = PaperHost.MinHeight; // 콘텐츠가 길어도 한 페이지 분량 기준
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
        if (PaperHost.IsMouseCaptured) PaperHost.ReleaseMouseCapture();
        if (_viewModel is not null) _viewModel.StatusMessage = SR.StatusReady;
        // 그리기 모드 종료 후 키보드 포커스를 본문 RTB 로 복원
        BodyEditor.Focus();
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
        // 직선 자동마감 직후 ClickCount==2 이벤트 억제
        // (끝점에서 더블클릭 시 ClickCount==1 로 선이 완성되고, ClickCount==2 가 뒤따라 발생하므로 무시)
        if (_suppressNextClickAfterLineFinish)
        {
            _suppressNextClickAfterLineFinish = false;
            if (e.ClickCount >= 2)
            {
                e.Handled = true;
                return;
            }
        }

        // ── 폴리선/스플라인 클릭 입력 모드 ──────────────────────────────
        if (_drawingPolyline_active)
        {
            var pos = e.GetPosition(PaperHost);
            pos.X = Math.Clamp(pos.X, 0, PaperHost.ActualWidth);
            pos.Y = Math.Clamp(pos.Y, 0, PaperHost.ActualHeight);

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
                // 자동마감 직후 ClickCount==2 이벤트가 뒤따르므로 억제 플래그 설정.
                if (_drawingPolyline_kind == ShapeKind.Line && _drawingPolyline_points.Count >= 2)
                {
                    FinishPolylineShape();
                    _suppressNextClickAfterLineFinish = true;
                }
            }

            e.Handled = true;
            return;
        }

        // ── Ctrl+클릭 → 오버레이 개체 멀티-선택 토글 (그리기 모드 아닐 때) ──────────
        // 오버레이(이미지/도형/표/글상자) 컨트롤은 BodyEditor 의 형제(같은 Grid)이므로
        // BodyEditor.PreviewMouseLeftButtonDown 으로는 잡히지 않는다. PaperHost 의
        // tunneling preview 가 부모이므로 여기서 가로채야 한다.
        // e.Handled = true 로 마킹하면 오버레이 자신의 bubbling MouseLeftButtonDown 핸들러
        // (드래그 시작 등) 가 호출되지 않는다 — `+=` 로 등록된 핸들러 기본 동작.
        if (!_drawingTextBox && !_drawingShape_active && !_drawingPolyline_active &&
            (Keyboard.Modifiers & ModifierKeys.Control) != 0 &&
            (Keyboard.Modifiers & ModifierKeys.Alt) == 0)
        {
            var ptForHit = e.GetPosition(PageEditorHost);
            var hitOverlay = FindAnyOverlayControlAt(ptForHit);
            if (hitOverlay != null)
            {
                ToggleMultiSelectControl(hitOverlay);
                Focus();
                Keyboard.Focus(this);
                e.Handled = true;
                return;
            }
        }

        // ── 일반 클릭 — 멀티-선택된 개체를 클릭하면 유지, 그 외엔 해제 ──────────────
        if (!_drawingTextBox && !_drawingShape_active && !_drawingPolyline_active &&
            (Keyboard.Modifiers & ModifierKeys.Control) == 0 &&
            _multiSelectedControls.Count > 0)
        {
            var ptForHit = e.GetPosition(PageEditorHost);
            var hitOverlay = FindAnyOverlayControlAt(ptForHit);
            // 멀티-선택된 개체 위 클릭: 유지(드래그 등 일반 동작 양보)
            // 멀티-선택 외 영역 클릭: 멀티-선택 해제
            if (hitOverlay == null || !_multiSelectedControls.Contains(hitOverlay))
                ClearMultiSelect();
        }

        // ── 단 너비 드래그 시작 ──────────────────────────────────────────────────
        if (!_drawingTextBox && !_drawingShape_active && !_drawingPolyline_active && _colDivHovering)
        {
            var ptPaper = e.GetPosition(PaperHost);
            if (TryHitColumnDivider(ptPaper, out int divLeftIdx))
            {
                _colDivDragging       = true;
                _colDivDragLeftIdx    = divLeftIdx;
                _colDivDragStartX     = ptPaper.X;
                _colDivDragStartWidths = (double[])_pageGeometry!.ColWidthsDip.Clone();
                PaperHost.CaptureMouse();
                Mouse.OverrideCursor = Cursors.SizeWE;
                e.Handled = true;
                return;
            }
        }

        // ── 마퀴(범위 드래그) 시작 — 그리기 모드 아닐 때 ──────────────────────────
        // 마퀴가 시작되는 조건:
        //   A) 용지 여백(Padding) 안쪽이고 오버레이 개체가 없을 때 (무수정자 클릭)
        //   B) Ctrl+드래그로 오버레이 개체가 없는 곳 (Ctrl 없이 텍스트 영역 드래그는 텍스트 선택으로 양보)
        // Shift·Alt 조합은 텍스트 확장선택·BehindText 드래그에 양보.
        if (!_drawingTextBox && !_drawingShape_active && !_drawingPolyline_active)
        {
            bool alt   = (Keyboard.Modifiers & ModifierKeys.Alt)   != 0;
            bool shift = (Keyboard.Modifiers & ModifierKeys.Shift) != 0;
            bool ctrl  = (Keyboard.Modifiers & ModifierKeys.Control) != 0;

            if (!alt && !shift)
            {
                // per-page 모드: PageEditorHost 기준 절대 좌표 사용 — 오버레이 Canvas 와 동일 좌표계.
                var ptDoc    = e.GetPosition(PageEditorHost);
                var ptPaper2 = e.GetPosition(PaperHost);
                var hitCtrl2 = FindAnyOverlayControlAt(ptDoc);

                // 용지 여백 영역: 페이지 Padding 의 바깥쪽.
                // Y 는 페이지 로컬 좌표로 변환해 모든 페이지에 동일하게 적용.
                double stride2    = _pageGeometry?.PageStrideDip ?? 1.0;
                double pageLocalY = stride2 > 0 ? ptDoc.Y % stride2 : ptDoc.Y;
                bool inMargin = _pageGeometry != null
                    && (ptDoc.X < _pageGeometry.PadLeftDip
                    ||  ptDoc.X > _pageGeometry.PageWidthDip - _pageGeometry.PadRightDip
                    ||  pageLocalY < _pageGeometry.PadTopDip
                    ||  pageLocalY > _pageGeometry.PageHeightDip - _pageGeometry.PadBottomDip);

                bool startMarquee = hitCtrl2 == null && (inMargin || ctrl);

                if (startMarquee)
                {
                    if (!ctrl) ClearMultiSelect();

                    ptPaper2.X = Math.Clamp(ptPaper2.X, 0, PaperHost.ActualWidth);
                    ptPaper2.Y = Math.Clamp(ptPaper2.Y, 0, PaperHost.ActualHeight);

                    _drawStart = ptPaper2;
                    _marqueeSelecting = true;

                    Canvas.SetLeft(DrawPreviewRect, ptPaper2.X);
                    Canvas.SetTop(DrawPreviewRect, ptPaper2.Y);
                    DrawPreviewRect.Width = 0;
                    DrawPreviewRect.Height = 0;
                    DrawPreviewRect.Visibility = Visibility.Visible;

                    PaperHost.CaptureMouse();
                    e.Handled = true;
                    return;
                }
            }
        }

        if (!_drawingTextBox && !_drawingShape_active) return;

        var startPos = e.GetPosition(PaperHost);
        startPos.X = Math.Clamp(startPos.X, 0, PaperHost.ActualWidth);
        startPos.Y = Math.Clamp(startPos.Y, 0, PaperHost.ActualHeight);

        _drawStart = startPos;
        _drawingInProgress = true;

        Canvas.SetLeft(DrawPreviewRect, startPos.X);
        Canvas.SetTop(DrawPreviewRect, startPos.Y);
        DrawPreviewRect.Width = 0;
        DrawPreviewRect.Height = 0;
        DrawPreviewRect.Visibility = Visibility.Visible;

        PaperHost.CaptureMouse();
        e.Handled = true;
    }

    private void OnPaperPreviewMouseMove(object sender, MouseEventArgs e)
    {
        // ── 단 너비 드래그 중 ─────────────────────────────────────────────────
        if (_colDivDragging) { HandleColDivDragMove(e); return; }

        if (_drawingPolyline_active && _drawingPolyline_points.Count > 0)
        {
            var pos = e.GetPosition(PaperHost);
            pos.X = Math.Clamp(pos.X, 0, PaperHost.ActualWidth);
            pos.Y = Math.Clamp(pos.Y, 0, PaperHost.ActualHeight);
            UpdatePolylinePreview(pos);
            return;
        }

        var pos2 = e.GetPosition(PaperHost);
        pos2.X = Math.Clamp(pos2.X, 0, PaperHost.ActualWidth);
        pos2.Y = Math.Clamp(pos2.Y, 0, PaperHost.ActualHeight);

        double x2 = Math.Min(_drawStart.X, pos2.X);
        double y2 = Math.Min(_drawStart.Y, pos2.Y);
        double w2 = Math.Abs(pos2.X - _drawStart.X);
        double h2 = Math.Abs(pos2.Y - _drawStart.Y);

        // 마퀴 드래그 중: DrawPreviewRect 업데이트만 (개체 수집은 MouseUp 에서)
        if (_marqueeSelecting)
        {
            Canvas.SetLeft(DrawPreviewRect, x2);
            Canvas.SetTop(DrawPreviewRect, y2);
            DrawPreviewRect.Width  = w2;
            DrawPreviewRect.Height = h2;
            return;
        }

        // ── 단 구분선 hover 감지 ─────────────────────────────────────────────
        if (!_drawingInProgress && !_drawingTextBox && !_drawingShape_active && !_drawingPolyline_active
            && e.LeftButton != MouseButtonState.Pressed && !_tableColResizeActive)
        {
            bool onDiv = TryHitColumnDivider(pos2, out _);
            if (onDiv != _colDivHovering)
            {
                _colDivHovering = onDiv;
                if (!_tableColResizeHovering)
                    Mouse.OverrideCursor = onDiv ? Cursors.SizeWE : null;
            }
        }

        if (!_drawingInProgress || (!_drawingTextBox && !_drawingShape_active)) return;

        Canvas.SetLeft(DrawPreviewRect, x2);
        Canvas.SetTop(DrawPreviewRect, y2);
        DrawPreviewRect.Width  = w2;
        DrawPreviewRect.Height = h2;
    }

    private void OnPaperPreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        // ── 단 너비 드래그 완료 ───────────────────────────────────────────────
        if (_colDivDragging) { FinishColDivDrag(e); return; }

        // ── 마퀴 드래그 완료: 사각형 안의 모든 개체 선택 ──
        if (_marqueeSelecting)
        {
            _marqueeSelecting = false;
            DrawPreviewRect.Visibility = Visibility.Collapsed;
            if (PaperHost.IsMouseCaptured) PaperHost.ReleaseMouseCapture();

            var mPos = e.GetPosition(PaperHost);
            mPos.X = Math.Clamp(mPos.X, 0, PaperHost.ActualWidth);
            mPos.Y = Math.Clamp(mPos.Y, 0, PaperHost.ActualHeight);

            double mx = Math.Min(_drawStart.X, mPos.X);
            double my = Math.Min(_drawStart.Y, mPos.Y);
            double mw = Math.Abs(mPos.X - _drawStart.X);
            double mh = Math.Abs(mPos.Y - _drawStart.Y);

            if (mw >= 4 && mh >= 4)
            {
                bool addMode = (Keyboard.Modifiers & ModifierKeys.Control) != 0;
                ApplyMarqueeSelection(new Rect(mx, my, mw, mh), addMode);
            }
            e.Handled = true;
            return;
        }

        if (!_drawingInProgress) return;

        var pos = e.GetPosition(PaperHost);
        pos.X = Math.Clamp(pos.X, 0, PaperHost.ActualWidth);
        pos.Y = Math.Clamp(pos.Y, 0, PaperHost.ActualHeight);

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

        // 드래그된 PaperHost 절대 좌표를 (AnchorPageIndex, OverlayXMm/YMm) 으로 분해.
        var model = new TextBoxObject
        {
            Shape    = _drawingShape,
            WidthMm  = w / TextBoxOverlay.DipsPerMm,
            HeightMm = h / TextBoxOverlay.DipsPerMm,
            Status   = NodeStatus.Modified,
        };
        CommitOverlayDragPosition(model, x, y);
        _viewModel?.AddOverlayBlockToCurrentSection(model);
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

        var pt = e.GetPosition(PageEditorHost); // 오버레이 Canvas 와 동일 좌표계

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

        // ④ 오버레이 표 (InFrontOfText/Fixed 우선, BehindText 차선)
        var hitTable = (FindCanvasChildAt(OverlayTableCanvas, pt)
                     ?? FindCanvasChildAt(UnderlayTableCanvas, pt))?.Tag as PolyDonky.Core.Table;
        if (hitTable is not null)
        {
            OpenContextMenu(BuildOverlayTableMenu(hitTable));
            e.Handled = true;
            return;
        }

        // ⑤ BodyEditor 내 컨텐츠 — ContextMenuOpening 이 처리 (e.Handled = false)
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
                var curPos = Mouse.GetPosition(PaperHost);
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
        menu.PlacementTarget = PaperHost;
        menu.IsOpen = true;
    }

    /// <summary>Canvas 자식 중 PaperHost/BodyEditor 기준 좌표 pt 아래 첫 번째 요소를 반환.</summary>
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
            StrokeColor       = "#2C3E50",
            StrokeThicknessPt = 1.5,
            FillColor         = kind is PolyDonky.Core.ShapeKind.Polygon
                                     or PolyDonky.Core.ShapeKind.ClosedSpline
                                 ? "#7FB3D9" : null,
            FillOpacity       = 0.7,
            Status            = NodeStatus.Modified,
        };
        CommitOverlayDragPosition(shape, minX, minY);

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
            StrokeColor       = "#2C3E50",
            StrokeThicknessPt = 1.5,
            FillColor         = kind is PolyDonky.Core.ShapeKind.Line
                                     or PolyDonky.Core.ShapeKind.Polyline
                                     or PolyDonky.Core.ShapeKind.Spline
                                 ? null : "#7FB3D9",
            FillOpacity = 0.7,
            Status      = NodeStatus.Modified,
        };
        CommitOverlayDragPosition(shape, xDip, yDip);

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
        // Polyline 그리기 완료 후 그리기 모드 플래그·캡처를 정리하고 본문 RTB 로 포커스 복귀 —
        // 이전에는 EndDrawingMode 가 호출되지 않아 _drawingPolyline_active 등이 true 로 남고
        // 키보드 포커스가 BodyEditor 로 돌아오지 않아 텍스트 입력이 안 되는 문제가 있었다.
        EndDrawingMode();
    }

    private void InsertShapeBlock(PolyDonky.Core.ShapeObject shape)
    {
        // 오버레이 모드 도형은 페이지 RTB 가 아닌 모델(_viewModel.Document.Sections) 에만 등록한다.
        // 캔버스 표시는 RebuildOverlayShapes() 가 모델을 읽어 처리.
        // (예전에는 앵커 단락을 활성 페이지 RTB 에 삽입했는데, 그 결과
        //  ① TextChanged → ScheduleLivePaginationRefresh → SetupPageEditors → RestoreCaretToLastEditor
        //     로 이어져 도형 그리기 직후 커서가 마지막 페이지 끝으로 점프하고,
        //  ② ParseAllPageEditors 가 앵커도 본문 블록으로 파싱해 모델의 도형과 중복되며,
        //  ③ 활성 RTB 가 바뀌면 다른 페이지에서 그린 도형 앵커를 못 찾아 사라지는 문제가 있었다.)
        _viewModel?.AddShapeToCurrentSection(shape);
        RebuildOverlayShapes();

        // 본문이 비어 일반 Paragraph 가 없으면 커서가 갈 곳이 없어 텍스트 입력 불가 — 항상 하나 보장.
        var doc = BodyEditor.Document;
        if (!doc.Blocks.OfType<System.Windows.Documents.Paragraph>().Any())
        {
            doc.Blocks.Add(new System.Windows.Documents.Paragraph());
        }
    }

    // ── 도형 오버레이 재구축 ──────────────────────────────────────────────
    private void RebuildOverlayShapes()
    {
        // 재구축 전 선택 해제 — Children.Clear() 직후 _selectedShapeCtrl / _multiSelectedControls 이 stale 참조가 됨.
        DeselectShape();
        ClearMultiSelect();
        OverlayShapeCanvas.Children.Clear();
        UnderlayShapeCanvas.Children.Clear();

        // per-page 모드에서 오버레이 도형 앵커는 어떤 페이지 RTB 에도 포함되지 않으므로
        // (PerPageDocumentSplitter 가 BodyBlocks 만 RTB 에 넣음) 모델을 직접 순회한다.
        // 이전에는 BodyEditor.Document.Blocks 에서 읽어 저장→재불러오기 후 도형이 사라지고,
        // 활성 페이지 RTB 가 바뀌면 다른 페이지에 있던 도형이 안 보이는 문제가 있었다.
        var section = _viewModel?.Document.Sections.FirstOrDefault();
        if (section is null) return;
        foreach (var coreBlock in section.Blocks)
        {
            if (coreBlock is not PolyDonky.Core.ShapeObject shape) continue;
            if (shape.WrapMode is not (PolyDonky.Core.ImageWrapMode.InFrontOfText
                                    or PolyDonky.Core.ImageWrapMode.BehindText)) continue;

            var ctrl = Services.FlowDocumentBuilder.BuildOverlayShapeControl(shape);
            ctrl.Tag = shape;
            PlaceOverlay(ctrl, shape);
            ctrl.Cursor = Cursors.SizeAll;

            // 우클릭은 PaperHost.PreviewMouseRightButtonDown 통합 핸들러가 처리 — 개별 ContextMenu 불필요.
            ctrl.MouseLeftButtonDown += OnOverlayShapeMouseDown;

            var canvas = shape.WrapMode == PolyDonky.Core.ImageWrapMode.BehindText
                ? UnderlayShapeCanvas
                : OverlayShapeCanvas;
            canvas.Children.Add(ctrl);
        }
    }

    private void RebuildOverlayTables()
    {
        ClearMultiSelect();
        OverlayTableCanvas.Children.Clear();
        UnderlayTableCanvas.Children.Clear();

        var overlaySection = _viewModel?.Document.Sections.FirstOrDefault();
        if (overlaySection is null) return;
        foreach (var coreBlock in overlaySection.Blocks)
        {
            if (coreBlock is not PolyDonky.Core.Table table) continue;
            if (table.WrapMode == PolyDonky.Core.TableWrapMode.Block) continue;

            var ctrl = Services.FlowDocumentBuilder.BuildOverlayTableControl(table);
            if (ctrl is null) continue;

            ctrl.Tag = table;
            PlaceOverlay(ctrl, table);
            ctrl.Cursor = Cursors.SizeAll;
            ctrl.MouseLeftButtonDown += OnOverlayTableMouseDown;

            var canvas = table.WrapMode == PolyDonky.Core.TableWrapMode.BehindText
                ? UnderlayTableCanvas
                : OverlayTableCanvas;
            canvas.Children.Add(ctrl);
        }
    }

    // ── 오버레이 표 메뉴 / 이벤트 ────────────────────────────────────────

    private System.Windows.Controls.ContextMenu BuildOverlayTableMenu(PolyDonky.Core.Table table)
    {
        var menu = new System.Windows.Controls.ContextMenu();
        menu.Items.Add(MakeMenuItem("표 속성(_T)...", () =>
        {
            var dlg = new TablePropertiesWindow(table) { Owner = this };
            if (dlg.ShowDialog() == true)
            {
                // WrapMode 변경 시 FlowDocument 앵커 블록을 교체해야 할 수 있음
                var anchor = BodyEditor.Document.Blocks
                    .FirstOrDefault(b => ReferenceEquals(b.Tag, table));
                if (anchor is not null)
                {
                    var newAnchor = table.WrapMode == PolyDonky.Core.TableWrapMode.Block
                        ? (System.Windows.Documents.Block)Services.FlowDocumentBuilder.BuildTable(table)
                        : new System.Windows.Documents.Paragraph
                          {
                              Tag        = table,
                              Margin     = new Thickness(0),
                              FontSize   = 0.1,
                              Foreground = System.Windows.Media.Brushes.Transparent,
                              Background = System.Windows.Media.Brushes.Transparent,
                          };
                    BodyEditor.Document.Blocks.InsertBefore(anchor, newAnchor);
                    BodyEditor.Document.Blocks.Remove(anchor);
                }
                RebuildOverlayTables();
                _viewModel?.MarkDirty();
            }
        }));
        menu.Items.Add(new System.Windows.Controls.Separator());
        menu.Items.Add(MakeMenuItem("표 삭제(_X)", () =>
        {
            _viewModel?.RemoveOverlayBlock(table);
            var anchor = BodyEditor.Document.Blocks
                .FirstOrDefault(b => ReferenceEquals(b.Tag, table));
            if (anchor is not null)
                BodyEditor.Document.Blocks.Remove(anchor);
            RebuildOverlayTables();
        }));
        return menu;
    }

    // ── 오버레이 표 드래그 이동 상태 ────────────────────────────────────────
    private FrameworkElement? _draggingOverlayTable;
    private Point  _overlayTableDragStart;
    private double _overlayTableDragStartLeft;
    private double _overlayTableDragStartTop;
    private bool   _overlayTableDragMoved;

    private void OnOverlayTableMouseDown(object sender, MouseButtonEventArgs e)
    {
        if (sender is not FrameworkElement fe) return;

        // 더블클릭 → 속성 다이얼로그
        if (e.ClickCount == 2 && fe.Tag is PolyDonky.Core.Table tbl)
        {
            var dlg = new TablePropertiesWindow(tbl) { Owner = this };
            if (dlg.ShowDialog() == true) { RebuildOverlayTables(); _viewModel?.MarkDirty(); }
            e.Handled = true;
            return;
        }

        // 단일 클릭 → 드래그 시작 (Block 모드 표는 드래그 불가)
        if (fe.Tag is PolyDonky.Core.Table t &&
            t.WrapMode == PolyDonky.Core.TableWrapMode.Block) return;

        if (fe.Parent is not Canvas canvas) return;
        _draggingOverlayTable       = fe;
        _overlayTableDragStart      = e.GetPosition(canvas);
        _overlayTableDragStartLeft  = Canvas.GetLeft(fe);
        _overlayTableDragStartTop   = Canvas.GetTop(fe);
        if (double.IsNaN(_overlayTableDragStartLeft)) _overlayTableDragStartLeft = 0;
        if (double.IsNaN(_overlayTableDragStartTop))  _overlayTableDragStartTop  = 0;
        _overlayTableDragMoved = false;
        fe.CaptureMouse();
        fe.MouseMove         += OnOverlayTableDragMove;
        fe.MouseLeftButtonUp += OnOverlayTableDragUp;
        e.Handled = true;
    }

    private void OnOverlayTableDragMove(object sender, MouseEventArgs e)
    {
        if (sender is not FrameworkElement fe || !ReferenceEquals(_draggingOverlayTable, fe)) return;
        if (fe.Parent is not Canvas canvas) return;
        var pos = e.GetPosition(canvas);
        double dx = pos.X - _overlayTableDragStart.X;
        double dy = pos.Y - _overlayTableDragStart.Y;
        if (Math.Abs(dx) > 0.5 || Math.Abs(dy) > 0.5) _overlayTableDragMoved = true;
        Canvas.SetLeft(fe, _overlayTableDragStartLeft + dx);
        Canvas.SetTop (fe, _overlayTableDragStartTop  + dy);
    }

    private void OnOverlayTableDragUp(object sender, MouseButtonEventArgs e)
    {
        if (sender is not FrameworkElement fe || !ReferenceEquals(_draggingOverlayTable, fe)) return;
        fe.ReleaseMouseCapture();
        fe.MouseMove         -= OnOverlayTableDragMove;
        fe.MouseLeftButtonUp -= OnOverlayTableDragUp;
        _draggingOverlayTable = null;

        if (_overlayTableDragMoved && fe.Tag is PolyDonky.Core.Table t)
        {
            double left = Canvas.GetLeft(fe);
            double top  = Canvas.GetTop(fe);
            if (double.IsNaN(left)) left = 0;
            if (double.IsNaN(top))  top  = 0;
            CommitOverlayDragPosition(t, left, top);
            _viewModel?.MarkDirty();
        }
        e.Handled = true;
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
            CommitOverlayDragPosition(s, left, top);
            _viewModel?.MarkDirty();
        }
        e.Handled = true;
    }

    // ── 글상자 단독 복사/잘라내기/붙여넣기 ──────────────────────────────
    // 통합 후 글상자도 PolyDonky.FlowSelection.v1 (List<Block>) 포맷을 단독 항목으로 사용한다.
    // 별도 FloatingObject* 포맷은 폐지 (IWPF 통합으로 모든 부유 개체가 Block 트리 안에서 처리).

    /// <summary>
    /// 선택된 글상자를 복사한다. 안쪽 본문에 포커스가 있어도 텍스트 선택이 비어 있으면
    /// "글상자 자체 복사" 의도로 간주 — Word/PowerPoint 와 동일한 mental model.
    /// 안쪽 본문에 텍스트 선택이 있으면 가로채지 않고 일반 복사에 양보.
    /// </summary>
    private bool TryCopySelectedFloatingObject()
    {
        if (_selectedOverlay is null) return false;
        if (_selectedOverlay.IsEditorFocusWithin && _selectedOverlay.HasEditorTextSelection)
            return false;

        var json = System.Text.Json.JsonSerializer.Serialize(
            new List<Block> { _selectedOverlay.Model }, JsonDefaults.Options);
        var dataObj = new System.Windows.DataObject();
        dataObj.SetData(FlowSelectionClipboardFormat, json);
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
        _viewModel?.RemoveOverlayBlock(overlay.Model);
        _selectedOverlay = null;
        BodyEditor.Focus();
        return true;
    }

    /// <summary>현재 섹션의 글상자(TextBoxObject) 를 캔버스에 다시 채워 그린다 (문서 로드 시).</summary>
    private void RebuildFloatingObjects()
    {
        ClearMultiSelect();
        FloatingCanvas.Children.Clear();
        _selectedOverlay = null;
        // 옛 InnerEditor 참조는 기각 — 다시 만들 overlay 의 InnerEditor 가 GotKeyboardFocus 시 갱신.
        _lastTextEditor = null;
        var section = _viewModel?.Document.Sections.FirstOrDefault();
        if (section is null) return;
        foreach (var obj in section.Blocks.OfType<TextBoxObject>())
        {
            AddTextBoxOverlay(obj);
        }
    }

    private TextBoxOverlay AddTextBoxOverlay(TextBoxObject model)
    {
        var overlay = new TextBoxOverlay(model);
        PlaceOverlay(overlay, model);
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
                _viewModel?.NotifyOverlayChanged();
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
                _viewModel?.NotifyOverlayChanged();
            }
        };

        overlay.AppearanceChangedCommitted += (_, _) => _viewModel?.NotifyOverlayChanged();

        overlay.GeometryChangedCommitted += (_, _) =>
        {
            // Canvas 절대 DIP → (페이지 인덱스, 페이지 로컬 mm) 변환.
            double left = Canvas.GetLeft(overlay); if (double.IsNaN(left)) left = 0;
            double top  = Canvas.GetTop(overlay);  if (double.IsNaN(top))  top  = 0;
            CommitOverlayDragPosition(model, left, top);
            model.WidthMm    = overlay.ActualWidth  / TextBoxOverlay.DipsPerMm;
            model.HeightMm   = overlay.ActualHeight / TextBoxOverlay.DipsPerMm;
            model.Status     = NodeStatus.Modified;
            _viewModel?.NotifyOverlayChanged();
        };

        overlay.ContentChangedCommitted += (_, _) => _viewModel?.NotifyOverlayChanged();

        overlay.DeleteRequested += (_, _) =>
        {
            FloatingCanvas.Children.Remove(overlay);
            _viewModel?.RemoveOverlayBlock(model);
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

    /// <summary>선택된 객체(도형/글상자/멀티-선택)를 종류 무관하게 삭제.</summary>
    private bool TryDeleteSelectedObject()
    {
        // 멀티-선택: 모든 선택 개체를 삭제
        if (_multiSelectedControls.Count > 0)
        {
            DeleteMultiSelected();
            return true;
        }
        if (_selectedShape is { } shape)
        {
            DeleteOverlayShape(shape);
            DeselectShape();
            return true;
        }
        if (_selectedOverlay is { } overlay)
        {
            // 안쪽 편집기(단일 InnerEditor 또는 다단 column RTB)에 포커스가 있으면 — selection 유무
            // 와 무관하게 — Del 키를 RTB 가 텍스트 삭제로 처리하도록 양보. 그렇지 않으면 글상자에서
            // 텍스트 편집 중 Del 을 누를 때마다 글상자 자체가 통째로 삭제되는 버그 발생.
            if (overlay.IsEditorFocusWithin)
                return false;
            FloatingCanvas.Children.Remove(overlay);
            _viewModel?.RemoveOverlayBlock(overlay.Model);
            _selectedOverlay = null;
            BodyEditor.Focus();
            return true;
        }
        return false;
    }

    /// <summary>선택된 객체를 종류 무관하게 클립보드로 복사.</summary>
    private bool TryCopySelectedObject()
    {
        if (_multiSelectedControls.Count > 0) return CopyMultiSelectedToClipboard();
        if (_selectedShape is { } shape) return CopyShapeToClipboard(shape);
        return TryCopySelectedFloatingObject();
    }

    /// <summary>잘라내기 = 복사 후 삭제.</summary>
    private bool TryCutSelectedObject()
    {
        if (_multiSelectedControls.Count > 0)
        {
            if (!CopyMultiSelectedToClipboard()) return false;
            DeleteMultiSelected();
            return true;
        }
        if (_selectedShape is { } shape)
        {
            if (!CopyShapeToClipboard(shape)) return false;
            DeleteOverlayShape(shape);
            DeselectShape();
            return true;
        }
        return TryCutSelectedFloatingObject();
    }

    /// <summary>붙여넣기는 클립보드 포맷에 따라 자동 분기.</summary>
    private bool TryPasteSelectedObject()
    {
        // 글상자 InnerEditor(또는 다단 column RTB) 편집 중이면 RichTextBox 기본 붙여넣기에 양보
        if (_selectedOverlay?.IsEditorFocusWithin == true)
            return false;

        // 통합 멀티-선택 포맷 — 모든 부유 개체(글상자 포함)가 단일 Block 리스트로 직렬화됨
        if (Clipboard.ContainsData(FlowSelectionClipboardFormat))
            return TryPasteFlowSelection();
        if (Clipboard.ContainsData(BlockClipboardFormat))
            return TryPasteBlockFromClipboard();
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

    // ── 마퀴(범위 드래그) 멀티-선택 헬퍼 ──────────────────────────────────────

    /// <summary>지정 점(BodyEditor 좌표)에 위치한 오버레이 Canvas 자식 반환. 없으면 null.</summary>
    private FrameworkElement? FindAnyOverlayControlAt(Point ptEditor)
    {
        // 오버레이 Canvas 와 BodyEditor 는 같은 Grid 좌표계 사용 — 좌표 변환 불필요.
        return FindCanvasChildAt(OverlayImageCanvas,  ptEditor)
            ?? FindCanvasChildAt(OverlayShapeCanvas,  ptEditor)
            ?? FindCanvasChildAt(OverlayTableCanvas,  ptEditor)
            ?? FindFloatingCanvasChildAt(ptEditor)
            ?? FindCanvasChildAt(UnderlayImageCanvas, ptEditor)
            ?? FindCanvasChildAt(UnderlayShapeCanvas, ptEditor)
            ?? FindCanvasChildAt(UnderlayTableCanvas, ptEditor);
    }

    /// <summary>FloatingCanvas 자식(TextBoxOverlay) 중 지정 점을 포함하는 것 반환.</summary>
    private FrameworkElement? FindFloatingCanvasChildAt(Point pt)
    {
        foreach (UIElement child in FloatingCanvas.Children)
        {
            if (child is not FrameworkElement fe) continue;
            double left = Canvas.GetLeft(fe); if (double.IsNaN(left)) left = 0;
            double top  = Canvas.GetTop(fe);  if (double.IsNaN(top))  top  = 0;
            double w = fe.ActualWidth  > 0 ? fe.ActualWidth  : fe.Width;
            double h = fe.ActualHeight > 0 ? fe.ActualHeight : fe.Height;
            if (double.IsNaN(w) || w <= 0 || double.IsNaN(h) || h <= 0) continue;
            if (pt.X >= left && pt.X <= left + w && pt.Y >= top && pt.Y <= top + h) return fe;
        }
        return null;
    }

    /// <summary>Canvas 자식의 Rect (Canvas 기준 좌표)를 반환한다.</summary>
    private static Rect GetCanvasChildRect(FrameworkElement fe)
    {
        double left = Canvas.GetLeft(fe); if (double.IsNaN(left)) left = 0;
        double top  = Canvas.GetTop(fe);  if (double.IsNaN(top))  top  = 0;
        double w = fe.ActualWidth  > 0 ? fe.ActualWidth  : (double.IsNaN(fe.Width)  ? 0 : fe.Width);
        double h = fe.ActualHeight > 0 ? fe.ActualHeight : (double.IsNaN(fe.Height) ? 0 : fe.Height);
        return new Rect(left, top, Math.Max(w, 0), Math.Max(h, 0));
    }

    /// <summary>멀티-선택 하이라이트(파란 글로우 효과) on/off.</summary>
    private static void SetMultiSelectHighlight(FrameworkElement fe, bool on)
    {
        fe.Effect = on
            ? new System.Windows.Media.Effects.DropShadowEffect
              {
                  Color       = WpfMedia.Colors.DodgerBlue,
                  ShadowDepth = 0,
                  BlurRadius  = 14,
                  Opacity     = 0.9,
              }
            : null;
    }

    /// <summary>Ctrl+A — 본문 텍스트 + 모든 오버레이(이미지/도형/표/글상자) 통합 선택.</summary>
    private void SelectAllIncludingOverlays()
    {
        ClearMultiSelect();

        // 본문 텍스트 전체 선택
        BodyEditor.SelectAll();

        // 모든 오버레이 Canvas 자식을 멀티-선택에 추가
        Canvas[] canvases = { OverlayImageCanvas, OverlayShapeCanvas, OverlayTableCanvas,
                               FloatingCanvas,
                               UnderlayImageCanvas, UnderlayShapeCanvas, UnderlayTableCanvas };
        foreach (var canvas in canvases)
        {
            foreach (UIElement child in canvas.Children)
            {
                if (child is not FrameworkElement fe) continue;
                _multiSelectedControls.Add(fe);
                SetMultiSelectHighlight(fe, true);
            }
        }

        if (_multiSelectedControls.Count > 0)
        {
            // 단일 선택과 공존하지 않도록 해제
            DeselectAllOverlays();
            Focus();
            Keyboard.Focus(this);
        }
    }

    /// <summary>모든 멀티-선택 해제.</summary>
    private void ClearMultiSelect()
    {
        foreach (var fe in _multiSelectedControls)
            SetMultiSelectHighlight(fe, false);
        _multiSelectedControls.Clear();
    }

    /// <summary>오버레이 개체를 멀티-선택에 추가하거나, 이미 있으면 제거(토글).</summary>
    private void ToggleMultiSelectControl(FrameworkElement fe)
    {
        int idx = _multiSelectedControls.IndexOf(fe);
        if (idx >= 0)
        {
            SetMultiSelectHighlight(fe, false);
            _multiSelectedControls.RemoveAt(idx);
        }
        else
        {
            // 단일 선택(도형/글상자) 과 공존하지 않도록 단일 선택 먼저 해제
            DeselectAllOverlays();
            _multiSelectedControls.Add(fe);
            SetMultiSelectHighlight(fe, true);
        }
    }

    /// <summary>
    /// 마퀴 사각형(PaperHost 좌표) 안에 위치하는 모든 오버레이 개체를 선택하고
    /// 본문 텍스트 블록도 해당 범위로 선택한다.
    /// </summary>
    private void ApplyMarqueeSelection(Rect marqueePaper, bool addToExisting)
    {
        if (!addToExisting) ClearMultiSelect();

        // 모든 오버레이 Canvas 순회 — 마퀴와 겹치는 자식 수집
        Canvas[] canvases = { OverlayImageCanvas, OverlayShapeCanvas, OverlayTableCanvas,
                               FloatingCanvas,
                               UnderlayImageCanvas, UnderlayShapeCanvas, UnderlayTableCanvas };
        foreach (var canvas in canvases)
        {
            foreach (UIElement child in canvas.Children)
            {
                if (child is not FrameworkElement fe) continue;
                var feRect = GetCanvasChildRect(fe);
                if (feRect.Width <= 0 || feRect.Height <= 0) continue;
                if (!feRect.IntersectsWith(marqueePaper)) continue;
                if (_multiSelectedControls.Contains(fe)) continue;
                _multiSelectedControls.Add(fe);
                SetMultiSelectHighlight(fe, true);
            }
        }

        // 본문 텍스트: BodyEditor 좌표 = PaperHost 좌표 (같은 Grid)
        // GetPositionFromPoint 로 시작·끝 TextPointer 를 구해 Selection 을 설정한다.
        try
        {
            var tl = marqueePaper.TopLeft;
            var br = marqueePaper.BottomRight;
            var startTp = BodyEditor.GetPositionFromPoint(tl, snapToText: true);
            var endTp   = BodyEditor.GetPositionFromPoint(br, snapToText: true);
            if (startTp != null && endTp != null)
            {
                if (startTp.CompareTo(endTp) > 0)
                    (startTp, endTp) = (endTp, startTp);
                BodyEditor.Selection.Select(startTp, endTp);
            }
        }
        catch { /* GetPositionFromPoint 실패 시 무시 */ }

        // 멀티-선택이 있으면 Window 로 포커스 이동 → Delete/Ctrl+C 처리를 Window 가 받는다.
        if (_multiSelectedControls.Count > 0)
        {
            Focus();
            Keyboard.Focus(this);
        }
        else
        {
            BodyEditor.Focus();
        }
    }

    /// <summary>멀티-선택된 오버레이 개체들을 모두 삭제.</summary>
    private void DeleteMultiSelected()
    {
        var toDelete = _multiSelectedControls.ToList();
        ClearMultiSelect();

        foreach (var fe in toDelete)
        {
            if (fe is TextBoxOverlay tbo)
            {
                FloatingCanvas.Children.Remove(tbo);
                _viewModel?.RemoveOverlayBlock(tbo.Model);
            }
            else if (fe.Tag is PolyDonky.Core.ShapeObject shape)
            {
                DeleteOverlayShape(shape);
            }
            else if (fe.Tag is PolyDonky.Core.ImageBlock img)
            {
                // ImageBlock anchor paragraph 찾아서 제거
                foreach (var block in BodyEditor.Document.Blocks.ToList())
                {
                    if (ReferenceEquals(block.Tag, img))
                    {
                        BodyEditor.Document.Blocks.Remove(block);
                        break;
                    }
                }
                RebuildOverlayImages();
            }
            else if (fe.Tag is PolyDonky.Core.Table tbl)
            {
                foreach (var block in BodyEditor.Document.Blocks.ToList())
                {
                    if (ReferenceEquals(block.Tag, tbl))
                    {
                        BodyEditor.Document.Blocks.Remove(block);
                        break;
                    }
                }
                RebuildOverlayTables();
            }
        }

        // 본문 텍스트 선택도 삭제
        if (!BodyEditor.Selection.IsEmpty)
            BodyEditor.Selection.Text = string.Empty;

        _viewModel?.MarkDirty();
        BodyEditor.Focus();
    }

    /// <summary>
    /// 멀티-선택된 모든 개체(오버레이 이미지/도형/표/글상자 + 본문 텍스트)를
    /// PolyDonky.FlowSelection.v1 포맷으로 클립보드에 저장한다.
    /// </summary>
    private bool CopyMultiSelectedToClipboard()
    {
        var blocks = new List<PolyDonky.Core.Block>();
        // _multiSelectedControls 에서 이미 수집한 Core 객체 (참조 ID 기준)
        // — 같은 객체가 BodyEditor 텍스트 선택을 통해 ExtractCoreSelection 에서 다시 반환되어도 중복 추가되지 않게 한다.
        var collectedRefs = new System.Collections.Generic.HashSet<object>(
            ReferenceEqualityComparer.Instance);

        // 오버레이 개체 (도형·이미지·표·글상자) 모두 단일 Block 직렬화 경로로 추출.
        foreach (var fe in _multiSelectedControls)
        {
            PolyDonky.Core.Block? coreBlock = fe switch
            {
                TextBoxOverlay tbo                       => tbo.Model,
                _ => fe.Tag switch
                {
                    PolyDonky.Core.ImageBlock     img   => img,
                    PolyDonky.Core.ShapeObject    shape => shape,
                    PolyDonky.Core.Table          tbl   => tbl,
                    PolyDonky.Core.TextBoxObject  tb    => tb,
                    _                                   => null,
                },
            };
            if (coreBlock is null) continue;
            collectedRefs.Add(coreBlock);  // 원본 참조 기억 (clone 이 아니라 fe.Tag 자체)
            try
            {
                var jsonClone = System.Text.Json.JsonSerializer.Serialize(coreBlock, PolyDonky.Core.JsonDefaults.Options);
                var clone = System.Text.Json.JsonSerializer.Deserialize<PolyDonky.Core.Block>(jsonClone, PolyDonky.Core.JsonDefaults.Options);
                if (clone != null) { ResetCoreBlockId(clone); blocks.Add(clone); }
            }
            catch { }
        }

        // 본문 텍스트 선택도 포함.
        // ExtractCoreSelection 의 동작을 인라인으로 재현하되, **Wpf Block.Tag** 가
        // collectedRefs(이미 _multiSelectedControls 로 수집한 Core 객체)와 같으면 스킵한다.
        // → 같은 오버레이 도형/이미지/표가 두 번 직렬화되는 것을 방지하면서,
        //   Inline 모드 이미지/도형(별도 BlockUIContainer.Tag) 은 정상 포함된다.
        var sel = BodyEditor.Selection;
        if (!sel.IsEmpty)
        {
            foreach (var wpfBlock in BodyEditor.Document.Blocks)
            {
                if (PolyDonky.App.Services.PageBreakPadder.IsPagePadding(wpfBlock)) continue;
                if (wpfBlock.ContentEnd.CompareTo(sel.Start) <= 0) continue;
                if (wpfBlock.ContentStart.CompareTo(sel.End) >= 0) break;

                // 멀티-선택 컨트롤이 이미 가리키는 Core 객체와 동일하면 스킵.
                if (wpfBlock.Tag is object tag && collectedRefs.Contains(tag)) continue;

                if (wpfBlock is System.Windows.Documents.Table wpfTable)
                    EnsureCoreTable(wpfTable);

                PolyDonky.Core.Block? core;
                if (wpfBlock is System.Windows.Documents.Paragraph wpfPara
                    && wpfBlock.Tag is not PolyDonky.Core.Table
                    && wpfBlock.Tag is not PolyDonky.Core.ImageBlock
                    && wpfBlock.Tag is not PolyDonky.Core.ShapeObject)
                {
                    bool clippedAtStart = wpfBlock.ContentStart.CompareTo(sel.Start) < 0;
                    bool clippedAtEnd   = wpfBlock.ContentEnd.CompareTo(sel.End)     > 0;
                    core = (clippedAtStart || clippedAtEnd)
                        ? PolyDonky.App.Services.FlowDocumentParser.ParseParagraphClipped(wpfPara, sel.Start, sel.End)
                        : PolyDonky.App.Services.FlowDocumentParser.ParseSingleBlock(wpfBlock);
                }
                else
                {
                    core = PolyDonky.App.Services.FlowDocumentParser.ParseSingleBlock(wpfBlock);
                }
                if (core is null) continue;

                try
                {
                    var jsonClone = System.Text.Json.JsonSerializer.Serialize(core, PolyDonky.Core.JsonDefaults.Options);
                    var clone = System.Text.Json.JsonSerializer.Deserialize<PolyDonky.Core.Block>(jsonClone, PolyDonky.Core.JsonDefaults.Options);
                    if (clone != null) { ResetCoreBlockId(clone); blocks.Add(clone); }
                }
                catch { }
            }
        }

        if (blocks.Count == 0) return false;

        var json = System.Text.Json.JsonSerializer.Serialize(blocks, PolyDonky.Core.JsonDefaults.Options);
        var dataObj = new DataObject();
        dataObj.SetData(FlowSelectionClipboardFormat, json);

        // plain-text 폴백 — 단락 + 글상자 텍스트
        var plainParts = blocks.Select(b => b switch
        {
            PolyDonky.Core.Paragraph     p  => p.GetPlainText(),
            PolyDonky.Core.TextBoxObject tb => tb.GetPlainText(),
            _                                => string.Empty,
        }).Where(s => !string.IsNullOrEmpty(s));
        dataObj.SetText(string.Join('\n', plainParts));

        Clipboard.SetDataObject(dataObj, copy: true);
        return true;
    }

}

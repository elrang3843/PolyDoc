using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Media3D;
using System.Windows.Shapes;
using PolyDoc.Core;
using WpfMedia = System.Windows.Media;

namespace PolyDoc.App.Views;

/// <summary>
/// 글상자(<see cref="TextBoxObject"/>) 의 시각 + 인터랙션 컨테이너.
///
/// - 부모는 <c>FloatingCanvas</c> (MainWindow). Canvas.Left/Top + Width/Height 로 위치/크기 결정.
/// - 모델은 mm 단위, 캔버스는 DIP 단위 — <see cref="DipsPerMm"/> 로 변환.
/// - 외곽(Border/Path) 클릭 = 드래그 이동, 내부 RichTextBox 클릭 = 서식 편집.
/// - 선택 시 4코너 핸들로 리사이즈.
/// - 우클릭 컨텍스트 메뉴: 속성, 앞/뒤 순서, 삭제.
/// </summary>
public partial class TextBoxOverlay : UserControl
{
    public const double DipsPerMm = 96.0 / 25.4;

    // ── PathGeometry 문자열 (100×100 정규화 공간, Stretch=Fill 로 자동 스케일) ──

    // 말풍선: 둥근 사각형 + 하단 중앙 삼각 꼬리
    private const string PathSpeech =
        "M 6,0 L 94,0 Q 100,0 100,6 L 100,70 Q 100,76 94,76 " +
        "L 58,76 L 50,95 L 42,76 L 6,76 Q 0,76 0,70 L 0,6 Q 0,0 6,0 Z";

    // 구름풍선: 여러 개의 둥근 튀어나온 부분으로 구성
    private const string PathCloud =
        "M 20,80 " +
        "C 5,80 0,65 5,55 C 0,45 5,32 18,32 " +
        "C 14,18 28,10 40,16 C 43,5 57,2 66,10 " +
        "C 73,3 87,6 88,20 C 100,20 100,38 93,46 " +
        "C 102,55 98,72 86,76 C 86,88 74,90 64,82 " +
        "C 58,92 44,93 38,84 C 28,92 18,90 20,80 Z";

    // 가시풍선: 12각 별 모양 (뾰족한 돌기)
    private const string PathSpiky =
        "M 50,0 L 58,35 L 94,25 L 68,50 L 93,75 " +
        "L 58,65 L 50,100 L 42,65 L 7,75 L 32,50 " +
        "L 6,25 L 42,35 Z";

    // 번개상자: 번개 볼트 실루엣
    private const string PathLightning =
        "M 65,0 L 22,52 L 46,52 L 35,100 L 78,48 L 54,48 Z";

    public TextBoxObject Model { get; }

    /// <summary>이 오버레이가 선택됨 — 호출자가 다른 오버레이 선택 해제할 때 사용.</summary>
    public event EventHandler? Selected;

    /// <summary>드래그/리사이즈 종료 — 호출자가 모델 mm 좌표 갱신 + Dirty 표시.</summary>
    public event EventHandler? GeometryChangedCommitted;

    /// <summary>본문 편집 종료 — 호출자가 모델 동기화 + Dirty 표시.</summary>
    public event EventHandler? ContentChangedCommitted;

    /// <summary>오버레이 삭제 요청 (Delete 키 / 컨텍스트 메뉴). 호출자가 모델/캔버스에서 제거.</summary>
    public event EventHandler? DeleteRequested;

    /// <summary>앞으로 가져오기 요청. 호출자가 Canvas ZOrder 조정.</summary>
    public event EventHandler? BringForwardRequested;

    /// <summary>뒤로 보내기 요청. 호출자가 Canvas ZOrder 조정.</summary>
    public event EventHandler? SendBackRequested;

    /// <summary>속성 변경 확정 (속성 대화상자 OK). 호출자가 Dirty 표시.</summary>
    public event EventHandler? AppearanceChangedCommitted;

    private bool _suppressTextChanged;

    public TextBoxOverlay(TextBoxObject model)
    {
        Model = model ?? throw new ArgumentNullException(nameof(model));
        InitializeComponent();

        // 초기 텍스트 로드 (plain text → FlowDocument)
        _suppressTextChanged = true;
        LoadModelTextToEditor();
        _suppressTextChanged = false;

        ApplyShapeFromModel();
        Loaded += (_, _) => UpdateHandlePositions();
        SizeChanged += (_, _) => UpdateHandlePositions();
    }

    // ── 선택 상태 ────────────────────────────────────────────────────

    private bool _isSelected;
    public bool IsSelected
    {
        get => _isSelected;
        set
        {
            if (_isSelected == value) return;
            _isSelected = value;
            SelectionFrame.Visibility = value ? Visibility.Visible : Visibility.Collapsed;
            HandlesCanvas.Visibility  = value ? Visibility.Visible : Visibility.Collapsed;
            if (!value && InnerEditor.IsKeyboardFocusWithin)
                Keyboard.ClearFocus();
        }
    }

    /// <summary>드래그 생성 직후 호출 — 선택 + 즉시 본문 편집 모드 진입.</summary>
    public void BeginEditing()
    {
        IsSelected = true;
        InnerEditor.Focus();
        Keyboard.Focus(InnerEditor);
    }

    // ── 모양/배경/테두리 적용 ────────────────────────────────────────

    private void ApplyShapeFromModel()
    {
        if (Model.Shape == TextBoxShape.Rectangle)
        {
            ShapeBorder.Visibility = Visibility.Visible;
            ShapePath.Visibility   = Visibility.Collapsed;

            ShapeBorder.BorderThickness = new Thickness(Math.Max(0.5, Model.BorderThicknessPt));
            if (TryParseColor(Model.BorderColor, out var bc))
                ShapeBorder.BorderBrush = new SolidColorBrush(bc);
            else
                ShapeBorder.BorderBrush = Brushes.Black;

            if (TryParseColor(Model.BackgroundColor, out var fillc))
                ShapeBorder.Background = new SolidColorBrush(fillc);
            else
                ShapeBorder.Background = Brushes.White;
        }
        else
        {
            ShapeBorder.Visibility = Visibility.Collapsed;
            ShapePath.Visibility   = Visibility.Visible;

            var pathData = Model.Shape switch
            {
                TextBoxShape.Speech    => PathSpeech,
                TextBoxShape.Cloud     => PathCloud,
                TextBoxShape.Spiky     => PathSpiky,
                TextBoxShape.Lightning => PathLightning,
                _                      => PathSpeech,
            };
            ShapePath.Data = Geometry.Parse(pathData);
            ShapePath.StrokeThickness = Math.Max(0.5, Model.BorderThicknessPt);

            if (TryParseColor(Model.BorderColor, out var bc))
                ShapePath.Stroke = new SolidColorBrush(bc);
            else
                ShapePath.Stroke = Brushes.Black;

            if (TryParseColor(Model.BackgroundColor, out var fillc))
                ShapePath.Fill = new SolidColorBrush(fillc);
            else
                ShapePath.Fill = Brushes.White;
        }

        // ── 안쪽 여백 (Margin = padding) ──────────────────────────────
        InnerEditor.Margin = new Thickness(
            Model.PaddingLeftMm   * DipsPerMm,
            Model.PaddingTopMm    * DipsPerMm,
            Model.PaddingRightMm  * DipsPerMm,
            Model.PaddingBottomMm * DipsPerMm);

        // ── 가로 정렬 ─────────────────────────────────────────────────
        var ta = Model.HAlign switch
        {
            TextBoxHAlign.Center  => System.Windows.TextAlignment.Center,
            TextBoxHAlign.Right   => System.Windows.TextAlignment.Right,
            TextBoxHAlign.Justify => System.Windows.TextAlignment.Justify,
            _                     => System.Windows.TextAlignment.Left,
        };
        InnerEditor.Document.TextAlignment = ta;
        foreach (var b in InnerEditor.Document.Blocks)
            if (b is System.Windows.Documents.Paragraph wp) wp.TextAlignment = ta;
    }

    // PolyDoc.Core.Color 와 충돌하므로 WpfMedia alias 로 명시.
    private static bool TryParseColor(string? hex, out WpfMedia.Color color)
    {
        color = default;
        if (string.IsNullOrWhiteSpace(hex)) return false;
        var s = hex.Trim();
        if (!s.StartsWith('#')) s = '#' + s;
        try { color = (WpfMedia.Color)WpfMedia.ColorConverter.ConvertFromString(s)!; return true; }
        catch { return false; }
    }

    // ── 모델 ↔ RichTextBox 동기화 ────────────────────────────────────

    private void LoadModelTextToEditor()
    {
        // 모델의 plain text → FlowDocument paragraphs
        var doc = new System.Windows.Documents.FlowDocument();
        foreach (var block in Model.Content)
        {
            if (block is PolyDoc.Core.Paragraph cp)
            {
                var para = new System.Windows.Documents.Paragraph(new System.Windows.Documents.Run(cp.GetPlainText()));
                doc.Blocks.Add(para);
            }
        }
        if (!doc.Blocks.Any())
            doc.Blocks.Add(new System.Windows.Documents.Paragraph());

        InnerEditor.Document = doc;
    }

    private void SyncEditorToModel()
    {
        // RichTextBox의 plain text를 모델에 동기화 (단락 유지)
        var doc = InnerEditor.Document;
        Model.Content.Clear();
        foreach (var block in doc.Blocks)
        {
            if (block is System.Windows.Documents.Paragraph para)
            {
                var range = new System.Windows.Documents.TextRange(para.ContentStart, para.ContentEnd);
                var cp = new PolyDoc.Core.Paragraph();
                var text = range.Text.TrimEnd('\r', '\n');
                if (text.Length > 0) cp.AddText(text);
                Model.Content.Add(cp);
            }
        }
        if (Model.Content.Count == 0)
            Model.Content.Add(new PolyDoc.Core.Paragraph());
    }

    // ── 내부 편집 ────────────────────────────────────────────────────

    private void OnInnerTextChanged(object sender, TextChangedEventArgs e)
    {
        if (_suppressTextChanged) return;
        SyncEditorToModel();
        Model.Status = NodeStatus.Modified;
        ContentChangedCommitted?.Invoke(this, EventArgs.Empty);
    }

    // ── 컨텍스트 메뉴 ────────────────────────────────────────────────

    private void OnContextMenuOpened(object sender, RoutedEventArgs e)
    {
        // 선택 상태를 컨텍스트 메뉴 열릴 때 확보
        Selected?.Invoke(this, EventArgs.Empty);
        IsSelected = true;
    }

    private void OnContextMenuProperties(object sender, RoutedEventArgs e)
    {
        var dlg = new TextBoxPropertiesWindow(Model) { Owner = Window.GetWindow(this) };
        if (dlg.ShowDialog() == true)
        {
            Model.BorderColor        = dlg.ResultBorderColor;
            Model.BorderThicknessPt  = dlg.ResultBorderThicknessPt;
            Model.BackgroundColor    = dlg.ResultBackgroundColor;
            Model.PaddingTopMm       = dlg.ResultPaddingTopMm;
            Model.PaddingBottomMm    = dlg.ResultPaddingBottomMm;
            Model.PaddingLeftMm      = dlg.ResultPaddingLeftMm;
            Model.PaddingRightMm     = dlg.ResultPaddingRightMm;
            Model.HAlign             = dlg.ResultHAlign;
            Model.VAlign             = dlg.ResultVAlign;
            Model.Status             = NodeStatus.Modified;
            ApplyShapeFromModel();
            AppearanceChangedCommitted?.Invoke(this, EventArgs.Empty);
        }
    }

    private void OnContextMenuCharFormat(object sender, RoutedEventArgs e)
    {
        InnerEditor.Focus();
        Keyboard.Focus(InnerEditor);
        var dlg = new CharFormatWindow(InnerEditor) { Owner = Window.GetWindow(this) };
        if (dlg.ShowDialog() == true)
        {
            SyncEditorToModel();
            Model.Status = NodeStatus.Modified;
            ContentChangedCommitted?.Invoke(this, EventArgs.Empty);
        }
        InnerEditor.Focus();
    }

    private void OnContextMenuParaFormat(object sender, RoutedEventArgs e)
    {
        InnerEditor.Focus();
        Keyboard.Focus(InnerEditor);
        var dlg = new ParaFormatWindow(InnerEditor) { Owner = Window.GetWindow(this) };
        if (dlg.ShowDialog() == true)
        {
            SyncEditorToModel();
            Model.Status = NodeStatus.Modified;
            ContentChangedCommitted?.Invoke(this, EventArgs.Empty);
        }
        InnerEditor.Focus();
    }

    private void OnContextMenuBringForward(object sender, RoutedEventArgs e)
        => BringForwardRequested?.Invoke(this, EventArgs.Empty);

    private void OnContextMenuSendBack(object sender, RoutedEventArgs e)
        => SendBackRequested?.Invoke(this, EventArgs.Empty);

    private void OnContextMenuDelete(object sender, RoutedEventArgs e)
        => DeleteRequested?.Invoke(this, EventArgs.Empty);

    // ── 드래그 이동 ──────────────────────────────────────────────────

    private bool _dragging;
    private Point _dragStart;
    private double _dragOrigLeft, _dragOrigTop;

    private void OnRootMouseDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ChangedButton != MouseButton.Left) return;

        // 핸들 Rectangle 클릭은 OnHandleMouseDown 이 처리 — 여기서 e.Handled 를 세팅하면
        // tunneling 이 멈춰 OnHandleMouseDown 이 호출되지 않으므로 반드시 빠져나온다.
        if (e.OriginalSource is Rectangle { Tag: string rTag } &&
            rTag is "TL" or "TR" or "BL" or "BR")
            return;

        Selected?.Invoke(this, EventArgs.Empty);

        if (IsInsideEditor(e.OriginalSource as DependencyObject))
            return;

        Focus();
        if (Parent is not IInputElement parent) return;
        _dragStart    = e.GetPosition(parent);
        _dragOrigLeft = SafeGetCanvasLeft(this);
        _dragOrigTop  = SafeGetCanvasTop(this);
        _dragging     = true;
        CaptureMouse();
        e.Handled = true;
    }

    private void OnRootMouseMove(object sender, MouseEventArgs e)
    {
        if (_dragging && Parent is IInputElement parent)
        {
            var pos = e.GetPosition(parent);
            Canvas.SetLeft(this, _dragOrigLeft + (pos.X - _dragStart.X));
            Canvas.SetTop(this,  _dragOrigTop  + (pos.Y - _dragStart.Y));
            return;
        }
        if (_resizing) DoResize(e);
    }

    private void OnRootMouseUp(object sender, MouseButtonEventArgs e)
    {
        if (_dragging)
        {
            _dragging = false;
            ReleaseMouseCapture();
            GeometryChangedCommitted?.Invoke(this, EventArgs.Empty);
            e.Handled = true;
        }
        if (_resizing)
        {
            _resizing = false;
            ReleaseMouseCapture();
            GeometryChangedCommitted?.Invoke(this, EventArgs.Empty);
            e.Handled = true;
        }
    }

    private static bool IsInsideEditor(DependencyObject? source)
    {
        while (source is not null)
        {
            if (source is RichTextBox) return true;
            // Paragraph/Run 등 FrameworkContentElement 는 Visual 이 아니므로
            // VisualTreeHelper 대신 LogicalTreeHelper 를 사용해야 한다.
            source = source is Visual or Visual3D
                ? VisualTreeHelper.GetParent(source)
                : LogicalTreeHelper.GetParent(source);
        }
        return false;
    }

    // ── 4코너 리사이즈 ───────────────────────────────────────────────

    private bool _resizing;
    private string _resizeCorner = "";
    private Point _resizeStart;
    private double _resizeOrigLeft, _resizeOrigTop, _resizeOrigW, _resizeOrigH;

    private void OnHandleMouseDown(object sender, MouseButtonEventArgs e)
    {
        if (sender is not Rectangle r || r.Tag is not string corner) return;
        if (Parent is not IInputElement parent) return;

        IsSelected = true;
        Selected?.Invoke(this, EventArgs.Empty);

        _resizeCorner   = corner;
        _resizing       = true;
        _resizeStart    = e.GetPosition(parent);
        _resizeOrigLeft = SafeGetCanvasLeft(this);
        _resizeOrigTop  = SafeGetCanvasTop(this);
        _resizeOrigW    = ActualWidth;
        _resizeOrigH    = ActualHeight;
        CaptureMouse();
        e.Handled = true;
    }

    private void DoResize(MouseEventArgs e)
    {
        if (!_resizing || Parent is not IInputElement parent) return;
        var pos = e.GetPosition(parent);
        double dx = pos.X - _resizeStart.X;
        double dy = pos.Y - _resizeStart.Y;
        const double minSize = 20;

        double newL = _resizeOrigLeft, newT = _resizeOrigTop;
        double newW = _resizeOrigW,    newH = _resizeOrigH;

        switch (_resizeCorner)
        {
            case "BR":
                newW = Math.Max(minSize, _resizeOrigW + dx);
                newH = Math.Max(minSize, _resizeOrigH + dy);
                break;
            case "TR":
                newW = Math.Max(minSize, _resizeOrigW + dx);
                newH = Math.Max(minSize, _resizeOrigH - dy);
                newT = _resizeOrigTop + (_resizeOrigH - newH);
                break;
            case "BL":
                newW = Math.Max(minSize, _resizeOrigW - dx);
                newH = Math.Max(minSize, _resizeOrigH + dy);
                newL = _resizeOrigLeft + (_resizeOrigW - newW);
                break;
            case "TL":
                newW = Math.Max(minSize, _resizeOrigW - dx);
                newH = Math.Max(minSize, _resizeOrigH - dy);
                newL = _resizeOrigLeft + (_resizeOrigW - newW);
                newT = _resizeOrigTop  + (_resizeOrigH - newH);
                break;
        }

        Canvas.SetLeft(this, newL);
        Canvas.SetTop(this,  newT);
        Width  = newW;
        Height = newH;
    }

    // ── 키 입력 ─────────────────────────────────────────────────────

    private void OnRootKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Delete && !InnerEditor.IsKeyboardFocusWithin)
        {
            DeleteRequested?.Invoke(this, EventArgs.Empty);
            e.Handled = true;
        }
    }

    // ── 핸들 위치 계산 ───────────────────────────────────────────────

    private void UpdateHandlePositions()
    {
        const double half = 4;
        double w = ActualWidth;
        double h = ActualHeight;
        Canvas.SetLeft(HandleTL, -half);     Canvas.SetTop(HandleTL, -half);
        Canvas.SetLeft(HandleTR, w - half);  Canvas.SetTop(HandleTR, -half);
        Canvas.SetLeft(HandleBL, -half);     Canvas.SetTop(HandleBL, h - half);
        Canvas.SetLeft(HandleBR, w - half);  Canvas.SetTop(HandleBR, h - half);
    }

    // ── 유틸 ────────────────────────────────────────────────────────

    private static double SafeGetCanvasLeft(UIElement el)
    {
        var v = Canvas.GetLeft(el);
        return double.IsNaN(v) ? 0 : v;
    }

    private static double SafeGetCanvasTop(UIElement el)
    {
        var v = Canvas.GetTop(el);
        return double.IsNaN(v) ? 0 : v;
    }
}

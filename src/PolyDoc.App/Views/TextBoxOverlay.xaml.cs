using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;
using PolyDoc.Core;

namespace PolyDoc.App.Views;

/// <summary>
/// 글상자(<see cref="TextBoxObject"/>) 의 시각 + 인터랙션 컨테이너.
///
/// - 부모는 <c>FloatingCanvas</c> (MainWindow). Canvas.Left/Top + Width/Height 로 위치/크기 결정.
/// - 모델은 mm 단위, 캔버스는 DIP 단위 — <see cref="DipsPerMm"/> 로 변환.
/// - 외곽(Border) 클릭 = 드래그 이동 시작, 내부 TextBox 클릭 = 텍스트 편집.
/// - 선택 시 4코너 핸들로 리사이즈 (드래그 중 모델 동기화는 호출자가
///   <see cref="GeometryChangedCommitted"/> 이벤트 받아서 처리).
///
/// 이번 사이클은 사각형만. Shape enum 은 모델에 보존되지만 Rectangle 외에는 동일 렌더.
/// </summary>
public partial class TextBoxOverlay : UserControl
{
    public const double DipsPerMm = 96.0 / 25.4;

    public TextBoxObject Model { get; }

    /// <summary>이 오버레이가 선택됨 — 호출자가 다른 오버레이 선택 해제할 때 사용.</summary>
    public event EventHandler? Selected;

    /// <summary>드래그/리사이즈 종료 — 호출자가 모델 mm 좌표 갱신 + Dirty 표시.</summary>
    public event EventHandler? GeometryChangedCommitted;

    /// <summary>본문 편집 종료 — 호출자가 모델 동기화 + Dirty 표시.</summary>
    public event EventHandler? ContentChangedCommitted;

    /// <summary>오버레이 삭제 요청 (Delete 키). 호출자가 모델/캔버스에서 제거.</summary>
    public event EventHandler? DeleteRequested;

    private bool _suppressTextChanged;

    public TextBoxOverlay(TextBoxObject model)
    {
        Model = model ?? throw new ArgumentNullException(nameof(model));
        InitializeComponent();

        // 초기 텍스트 로드.
        _suppressTextChanged = true;
        InnerEditor.Text = model.GetPlainText();
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
            // 선택 해제 시 inner editor 도 포커스 풀어 캐럿이 사라지도록.
            if (!value && InnerEditor.IsKeyboardFocusWithin)
            {
                Keyboard.ClearFocus();
            }
        }
    }

    /// <summary>드래그 생성 직후 호출 — 선택 + 즉시 본문 편집 모드 진입.</summary>
    public void BeginEditing()
    {
        IsSelected = true;
        InnerEditor.Focus();
        Keyboard.Focus(InnerEditor);
    }

    // ── 모양/배경/테두리 적용 ───────────────────────────────────────

    private void ApplyShapeFromModel()
    {
        // Rectangle 만 렌더 (다른 모양은 다음 사이클).
        ShapeBorder.BorderThickness = new Thickness(Math.Max(0.5, Model.BorderThicknessPt));

        if (TryParseColor(Model.BorderColor, out var bc))
            ShapeBorder.BorderBrush = new SolidColorBrush(bc);
        // null/빈 = 검정 기본 유지

        if (TryParseColor(Model.BackgroundColor, out var fillc))
            ShapeBorder.Background = new SolidColorBrush(fillc);
        // null/빈 = 흰색 기본 유지
    }

    private static bool TryParseColor(string? hex, out Color color)
    {
        color = default;
        if (string.IsNullOrWhiteSpace(hex)) return false;
        var s = hex.Trim();
        if (!s.StartsWith('#')) s = '#' + s;
        try { color = (Color)ColorConverter.ConvertFromString(s)!; return true; }
        catch { return false; }
    }

    // ── 내부 편집 ───────────────────────────────────────────────────

    private void OnInnerTextChanged(object sender, TextChangedEventArgs e)
    {
        if (_suppressTextChanged) return;
        Model.SetPlainText(InnerEditor.Text);
        Model.Status = NodeStatus.Modified;
        ContentChangedCommitted?.Invoke(this, EventArgs.Empty);
    }

    // ── 드래그 이동 ─────────────────────────────────────────────────

    private bool _dragging;
    private Point _dragStart;
    private double _dragOrigLeft, _dragOrigTop;

    private void OnRootMouseDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ChangedButton != MouseButton.Left) return;

        // 핸들 클릭은 OnHandleMouseDown 이 먼저 처리하고 e.Handled=true 로 막음.
        // 여기까지 왔다면 핸들 외 클릭 — 본문 영역 또는 외곽 보더.

        Selected?.Invoke(this, EventArgs.Empty);

        // 클릭이 inner TextBox 내부에 떨어졌으면 텍스트 편집 모드로 — 드래그 안 함.
        if (IsInsideEditor(e.OriginalSource as DependencyObject))
        {
            return; // TextBox 가 자체적으로 focus/caret 처리
        }

        // 외곽 클릭 — 드래그 이동 시작.
        Focus();
        if (Parent is not IInputElement parent) return;
        _dragStart      = e.GetPosition(parent);
        _dragOrigLeft   = SafeGetCanvasLeft(this);
        _dragOrigTop    = SafeGetCanvasTop(this);
        _dragging       = true;
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
            if (source is TextBox) return true;
            source = VisualTreeHelper.GetParent(source);
        }
        return false;
    }

    // ── 4코너 리사이즈 ──────────────────────────────────────────────

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
        // 캡처는 UserControl 에 — 핸들 자체에 캡처하면 마우스가 핸들 밖으로 나가는 동안 이벤트가 끊김.
        CaptureMouse();
        e.Handled = true;
    }

    private void DoResize(MouseEventArgs e)
    {
        if (!_resizing || Parent is not IInputElement parent) return;
        var pos = e.GetPosition(parent);
        double dx = pos.X - _resizeStart.X;
        double dy = pos.Y - _resizeStart.Y;
        const double minSize = 20; // 최소 20 DIP

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

    // ── 키 입력 (Delete) ────────────────────────────────────────────

    private void OnRootKeyDown(object sender, KeyEventArgs e)
    {
        // 본문 편집 중에는 Delete 가 텍스트 한 글자 삭제 — 오버레이 삭제 X
        if (e.Key == Key.Delete && !InnerEditor.IsKeyboardFocusWithin)
        {
            DeleteRequested?.Invoke(this, EventArgs.Empty);
            e.Handled = true;
        }
    }

    // ── 핸들 위치 계산 ──────────────────────────────────────────────

    private void UpdateHandlePositions()
    {
        const double half = 4;
        double w = ActualWidth;
        double h = ActualHeight;
        Canvas.SetLeft(HandleTL, -half);  Canvas.SetTop(HandleTL, -half);
        Canvas.SetLeft(HandleTR, w - half); Canvas.SetTop(HandleTR, -half);
        Canvas.SetLeft(HandleBL, -half);  Canvas.SetTop(HandleBL, h - half);
        Canvas.SetLeft(HandleBR, w - half); Canvas.SetTop(HandleBR, h - half);
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

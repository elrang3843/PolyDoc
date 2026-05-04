using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Effects;
using System.Windows.Shapes;
using PolyDonky.App.Services;
using PolyDonky.Core;

namespace PolyDonky.App.Views;

/// <summary>
/// 선택된 도형의 크기/모양 편집 핸들 — 코너 8개(bbox 도형) 또는 정점 핸들(점-기반 도형).
/// 핸들은 도형 컨트롤과 같은 부모 Canvas 의 자식으로 추가된다 (절대 좌표).
/// </summary>
public partial class MainWindow
{
    // ── 핸들 상수 ─────────────────────────────────────────────────────────
    private const double ShapeHandleSize       = 9.0;
    private const double ShapeHandleHitPad     = 0.0;   // 추가 hit 영역 (현재 미사용)
    private const double ShapeMinSizeMm        = 2.0;
    private const string VertexHandleTagPrefix = "V";

    // ── 핸들 상태 ─────────────────────────────────────────────────────────
    private readonly List<Rectangle> _shapeEditHandles = new();
    private Rectangle? _draggingShapeHandle;
    private string     _draggingHandleTag = string.Empty;
    private Point      _handleDragStartCanvasDip;
    private double     _handleDragOrigXMm;
    private double     _handleDragOrigYMm;
    private double     _handleDragOrigWMm;
    private double     _handleDragOrigHMm;
    private List<(double X, double Y)> _handleDragOrigPoints = new();
    private bool       _handleDragMoved;

    private static readonly HashSet<ShapeKind> _pointBasedKinds = new()
    {
        ShapeKind.Line,
        ShapeKind.Polyline,
        ShapeKind.Spline,
        ShapeKind.Polygon,
        ShapeKind.ClosedSpline,
    };

    // ── Public hooks (SelectShape/DeselectShape 에서 호출) ───────────────────

    private void ShowShapeEditHandles(FrameworkElement shapeCtrl, ShapeObject shape)
    {
        HideShapeEditHandles();
        if (shapeCtrl.Parent is not Canvas parent) return;

        if (_pointBasedKinds.Contains(shape.Kind))
        {
            EnsureLineEndpointsForEditing(shape);

            var pts = shape.Points;
            for (int i = 0; i < pts.Count; i++)
            {
                var h = CreateHandle($"{VertexHandleTagPrefix}{i}", isVertex: true);
                PositionVertexHandle(h, shapeCtrl, shape, pts[i]);
                parent.Children.Add(h);
                _shapeEditHandles.Add(h);
            }
        }
        else
        {
            foreach (var tag in new[] { "TL", "T", "TR", "L", "R", "BL", "B", "BR" })
            {
                var h = CreateHandle(tag, isVertex: false);
                PositionBBoxHandle(h, shapeCtrl, shape, tag);
                parent.Children.Add(h);
                _shapeEditHandles.Add(h);
            }
        }
    }

    private void HideShapeEditHandles()
    {
        foreach (var h in _shapeEditHandles)
        {
            if (h.Parent is Canvas c) c.Children.Remove(h);
            h.MouseLeftButtonDown -= OnShapeHandleMouseDown;
        }
        _shapeEditHandles.Clear();
    }

    private void RefreshShapeEditHandlePositions()
    {
        if (_selectedShape is null || _selectedShapeCtrl is null) return;
        foreach (var h in _shapeEditHandles)
        {
            if (h.Tag is not string tag) continue;
            if (tag.StartsWith(VertexHandleTagPrefix, StringComparison.Ordinal))
            {
                if (!int.TryParse(tag.AsSpan(1), out int idx)) continue;
                if (idx < 0 || idx >= _selectedShape.Points.Count) continue;
                PositionVertexHandle(h, _selectedShapeCtrl, _selectedShape, _selectedShape.Points[idx]);
            }
            else
            {
                PositionBBoxHandle(h, _selectedShapeCtrl, _selectedShape, tag);
            }
        }
    }

    // ── 핸들 생성 / 위치 ─────────────────────────────────────────────────

    private Rectangle CreateHandle(string tag, bool isVertex)
    {
        var fill = isVertex
            ? (Brush)new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xFF, 0xC0, 0x40))
            : Brushes.White;
        var stroke = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0x29, 0x6B, 0xC9));

        var h = new Rectangle
        {
            Width            = ShapeHandleSize,
            Height           = ShapeHandleSize,
            Fill             = fill,
            Stroke           = stroke,
            StrokeThickness  = 1.0,
            Tag              = tag,
            Cursor           = GetHandleCursor(tag, isVertex),
            IsHitTestVisible = true,
        };
        h.MouseLeftButtonDown += OnShapeHandleMouseDown;
        Panel.SetZIndex(h, 1000);
        return h;
    }

    private static Cursor GetHandleCursor(string tag, bool isVertex)
    {
        if (isVertex) return Cursors.Cross;
        return tag switch
        {
            "TL" or "BR" => Cursors.SizeNWSE,
            "TR" or "BL" => Cursors.SizeNESW,
            "T"  or "B"  => Cursors.SizeNS,
            "L"  or "R"  => Cursors.SizeWE,
            _            => Cursors.SizeAll,
        };
    }

    private static void PositionVertexHandle(Rectangle h, FrameworkElement ctrl, ShapeObject shape, ShapePoint pt)
    {
        double left = Canvas.GetLeft(ctrl); if (double.IsNaN(left)) left = 0;
        double top  = Canvas.GetTop (ctrl); if (double.IsNaN(top))  top  = 0;
        double cx = left + FlowDocumentBuilder.MmToDip(pt.X);
        double cy = top  + FlowDocumentBuilder.MmToDip(pt.Y);

        ApplyShapeRotation(ref cx, ref cy, ctrl, shape);

        Canvas.SetLeft(h, cx - ShapeHandleSize / 2);
        Canvas.SetTop (h, cy - ShapeHandleSize / 2);
    }

    private static void PositionBBoxHandle(Rectangle h, FrameworkElement ctrl, ShapeObject shape, string tag)
    {
        double left = Canvas.GetLeft(ctrl); if (double.IsNaN(left)) left = 0;
        double top  = Canvas.GetTop (ctrl); if (double.IsNaN(top))  top  = 0;
        double w    = FlowDocumentBuilder.MmToDip(shape.WidthMm);
        double hd   = FlowDocumentBuilder.MmToDip(shape.HeightMm);

        double cx = tag.Contains('L') ? left
                  : tag.Contains('R') ? left + w
                  : left + w / 2;
        double cy = tag.Contains('T') ? top
                  : tag.Contains('B') ? top + hd
                  : top + hd / 2;

        ApplyShapeRotation(ref cx, ref cy, ctrl, shape);

        Canvas.SetLeft(h, cx - ShapeHandleSize / 2);
        Canvas.SetTop (h, cy - ShapeHandleSize / 2);
    }

    /// <summary>
    /// 도형의 회전 각도를 적용해 (cx, cy) 핸들 위치를 도형 중심 기준으로 회전시킨다.
    /// FlowDocumentBuilder.BuildShapeVisual 가 canvas 전체에 RenderTransform 으로 회전을
    /// 적용하므로(원점 0.5,0.5), 핸들도 같은 변환을 따라가야 도형 위에 정확히 얹힌다.
    /// </summary>
    private static void ApplyShapeRotation(ref double cx, ref double cy, FrameworkElement ctrl, ShapeObject shape)
    {
        double angle = shape.RotationAngleDeg;
        if (Math.Abs(angle) < 0.01) return;

        double left = Canvas.GetLeft(ctrl); if (double.IsNaN(left)) left = 0;
        double top  = Canvas.GetTop (ctrl); if (double.IsNaN(top))  top  = 0;
        double centerX = left + FlowDocumentBuilder.MmToDip(shape.WidthMm)  / 2.0;
        double centerY = top  + FlowDocumentBuilder.MmToDip(shape.HeightMm) / 2.0;

        double rad = angle * Math.PI / 180.0;
        double cos = Math.Cos(rad), sin = Math.Sin(rad);
        double dx = cx - centerX;
        double dy = cy - centerY;
        cx = centerX + dx * cos - dy * sin;
        cy = centerY + dx * sin + dy * cos;
    }

    /// <summary>캔버스 좌표 델타를 도형 로컬 좌표 델타로 역회전.</summary>
    private static (double dxLocal, double dyLocal) InverseRotateDelta(double dxCanvas, double dyCanvas, double angleDeg)
    {
        if (Math.Abs(angleDeg) < 0.01) return (dxCanvas, dyCanvas);
        double rad = -angleDeg * Math.PI / 180.0;
        double cos = Math.Cos(rad), sin = Math.Sin(rad);
        return (dxCanvas * cos - dyCanvas * sin,
                dxCanvas * sin + dyCanvas * cos);
    }

    // ── 마우스 핸들러 ───────────────────────────────────────────────────

    private void OnShapeHandleMouseDown(object sender, MouseButtonEventArgs e)
    {
        if (sender is not Rectangle handle) return;
        if (_selectedShape is null || _selectedShapeCtrl is null) return;
        if (handle.Parent is not Canvas canvas) return;

        e.Handled = true;

        _draggingShapeHandle      = handle;
        _draggingHandleTag        = handle.Tag as string ?? string.Empty;
        _handleDragStartCanvasDip = e.GetPosition(canvas);
        _handleDragOrigXMm        = _selectedShape.OverlayXMm;
        _handleDragOrigYMm        = _selectedShape.OverlayYMm;
        _handleDragOrigWMm        = _selectedShape.WidthMm;
        _handleDragOrigHMm        = _selectedShape.HeightMm;
        _handleDragOrigPoints     = _selectedShape.Points
            .Select(p => (p.X, p.Y))
            .ToList();
        _handleDragMoved          = false;

        handle.CaptureMouse();
        handle.MouseMove         += OnShapeHandleMouseMove;
        handle.MouseLeftButtonUp += OnShapeHandleMouseUp;
    }

    private void OnShapeHandleMouseMove(object sender, MouseEventArgs e)
    {
        if (sender is not Rectangle handle || !ReferenceEquals(_draggingShapeHandle, handle)) return;
        if (_selectedShape is null) return;
        if (handle.Parent is not Canvas canvas) return;

        var pos = e.GetPosition(canvas);
        double dxCanvasMm = (pos.X - _handleDragStartCanvasDip.X) / FlowDocumentBuilder.MmToDip(1.0);
        double dyCanvasMm = (pos.Y - _handleDragStartCanvasDip.Y) / FlowDocumentBuilder.MmToDip(1.0);

        // 회전된 도형은 캔버스 좌표 델타를 도형 로컬 좌표로 역회전해서 처리해야
        // 핸들 드래그 방향과 변형 결과가 일치한다.
        var (dxMm, dyMm) = InverseRotateDelta(dxCanvasMm, dyCanvasMm, _selectedShape.RotationAngleDeg);

        if (Math.Abs(dxMm) > 0.01 || Math.Abs(dyMm) > 0.01) _handleDragMoved = true;

        if (_draggingHandleTag.StartsWith(VertexHandleTagPrefix, StringComparison.Ordinal))
            ApplyVertexDrag(dxMm, dyMm);
        else
            ApplyBBoxResize(_draggingHandleTag, dxMm, dyMm);

        _selectedShape.Status = NodeStatus.Modified;
        RefreshSelectedShapeVisual();
    }

    private void OnShapeHandleMouseUp(object sender, MouseButtonEventArgs e)
    {
        if (sender is not Rectangle handle || !ReferenceEquals(_draggingShapeHandle, handle)) return;
        handle.ReleaseMouseCapture();
        handle.MouseMove         -= OnShapeHandleMouseMove;
        handle.MouseLeftButtonUp -= OnShapeHandleMouseUp;
        _draggingShapeHandle = null;

        if (_handleDragMoved && _selectedShape is not null)
        {
            // 정점 편집 후 bbox 정규화 — 음수 좌표나 bbox 밖 정점이 있으면 내부 좌표로 옮긴다.
            if (_pointBasedKinds.Contains(_selectedShape.Kind))
                NormalizeShapeBoundingBox(_selectedShape);
            _viewModel?.MarkDirty();
            RefreshSelectedShapeVisual();
        }

        e.Handled = true;
    }

    // ── 변환 로직 ───────────────────────────────────────────────────────

    private void ApplyVertexDrag(double dxMm, double dyMm)
    {
        if (_selectedShape is null) return;
        if (!int.TryParse(_draggingHandleTag.AsSpan(1), out int idx)) return;
        if (idx < 0 || idx >= _handleDragOrigPoints.Count) return;

        var (origX, origY) = _handleDragOrigPoints[idx];
        if (idx >= _selectedShape.Points.Count) return;
        _selectedShape.Points[idx] = new ShapePoint
        {
            X = origX + dxMm,
            Y = origY + dyMm,
        };
    }

    private void ApplyBBoxResize(string tag, double dxMm, double dyMm)
    {
        if (_selectedShape is null) return;
        double newX = _handleDragOrigXMm;
        double newY = _handleDragOrigYMm;
        double newW = _handleDragOrigWMm;
        double newH = _handleDragOrigHMm;

        bool moveLeft   = tag.Contains('L');
        bool moveRight  = tag.Contains('R');
        bool moveTop    = tag.Contains('T');
        bool moveBottom = tag.Contains('B');

        if (moveLeft)
        {
            newX = _handleDragOrigXMm + dxMm;
            newW = _handleDragOrigWMm - dxMm;
        }
        if (moveRight)
        {
            newW = _handleDragOrigWMm + dxMm;
        }
        if (moveTop)
        {
            newY = _handleDragOrigYMm + dyMm;
            newH = _handleDragOrigHMm - dyMm;
        }
        if (moveBottom)
        {
            newH = _handleDragOrigHMm + dyMm;
        }

        if (newW < ShapeMinSizeMm)
        {
            if (moveLeft) newX = _handleDragOrigXMm + (_handleDragOrigWMm - ShapeMinSizeMm);
            newW = ShapeMinSizeMm;
        }
        if (newH < ShapeMinSizeMm)
        {
            if (moveTop) newY = _handleDragOrigYMm + (_handleDragOrigHMm - ShapeMinSizeMm);
            newH = ShapeMinSizeMm;
        }

        // 회전된 도형은 회전 중심(원본 bbox 중심)을 그대로 유지해야 시각적 점프가 없다.
        // 로컬 공간 리사이즈를 그대로 적용하면 RenderTransformOrigin(0.5,0.5)이 이동해
        // 화면상 도형이 갑자기 튄다. 따라서 newX/newY 를 원본 중심에 다시 정렬한다.
        if (Math.Abs(_selectedShape.RotationAngleDeg) > 0.01)
        {
            double origCenterX = _handleDragOrigXMm + _handleDragOrigWMm / 2.0;
            double origCenterY = _handleDragOrigYMm + _handleDragOrigHMm / 2.0;
            newX = origCenterX - newW / 2.0;
            newY = origCenterY - newH / 2.0;
        }

        _selectedShape.OverlayXMm = newX;
        _selectedShape.OverlayYMm = newY;
        _selectedShape.WidthMm    = newW;
        _selectedShape.HeightMm   = newH;
    }

    private static void EnsureLineEndpointsForEditing(ShapeObject shape)
    {
        if (shape.Kind == ShapeKind.Line && shape.Points.Count == 0)
        {
            shape.Points.Add(new ShapePoint { X = 0, Y = 0 });
            shape.Points.Add(new ShapePoint { X = shape.WidthMm, Y = shape.HeightMm });
        }
    }

    /// <summary>
    /// 정점 편집 후 bbox 정규화 — 모든 점이 [0, Width] × [0, Height] 안에 들어오도록
    /// OverlayXMm/YMm 와 Width/Height 를 조정하고 점 좌표를 평행이동한다.
    /// </summary>
    private static void NormalizeShapeBoundingBox(ShapeObject shape)
    {
        if (shape.Points.Count == 0) return;

        double minX = shape.Points.Min(p => p.X);
        double minY = shape.Points.Min(p => p.Y);
        double maxX = shape.Points.Max(p => p.X);
        double maxY = shape.Points.Max(p => p.Y);

        double shiftX = Math.Min(minX, 0);
        double shiftY = Math.Min(minY, 0);

        if (shiftX != 0 || shiftY != 0)
        {
            for (int i = 0; i < shape.Points.Count; i++)
            {
                shape.Points[i] = new ShapePoint
                {
                    X = shape.Points[i].X - shiftX,
                    Y = shape.Points[i].Y - shiftY,
                };
            }
            shape.OverlayXMm += shiftX;
            shape.OverlayYMm += shiftY;
            maxX -= shiftX;
            maxY -= shiftY;
        }

        if (maxX > shape.WidthMm)  shape.WidthMm  = maxX;
        if (maxY > shape.HeightMm) shape.HeightMm = maxY;
    }

    // ── 선택 도형 시각 갱신 ───────────────────────────────────────────────

    /// <summary>
    /// 선택된 도형의 Canvas 컨트롤을 새로 빌드해 부모 Canvas 에 교체 삽입한다.
    /// 핸들은 동일한 부모 Canvas 의 자식이므로 영향받지 않으며, 위치만 새로 계산된다.
    /// </summary>
    private void RefreshSelectedShapeVisual()
    {
        if (_selectedShape is null || _selectedShapeCtrl is null) return;
        var oldCtrl = _selectedShapeCtrl;
        if (oldCtrl.Parent is not Canvas parent) return;

        int idx = parent.Children.IndexOf(oldCtrl);

        var newCtrl = FlowDocumentBuilder.BuildOverlayShapeControl(_selectedShape);
        newCtrl.Tag = _selectedShape;
        PlaceOverlay(newCtrl, _selectedShape);
        newCtrl.Cursor = Cursors.SizeAll;
        newCtrl.MouseLeftButtonDown += OnOverlayShapeMouseDown;
        newCtrl.Effect = new DropShadowEffect
        {
            Color       = Colors.DodgerBlue,
            ShadowDepth = 0,
            BlurRadius  = 14,
            Opacity     = 0.9,
        };

        parent.Children.Remove(oldCtrl);
        if (idx >= 0 && idx <= parent.Children.Count)
            parent.Children.Insert(idx, newCtrl);
        else
            parent.Children.Add(newCtrl);

        _selectedShapeCtrl = newCtrl;

        RefreshShapeEditHandlePositions();
    }
}

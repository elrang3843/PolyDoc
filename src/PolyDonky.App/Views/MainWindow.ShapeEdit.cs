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
    private const string VertexHandleTagPrefix  = "V";
    private const string SegmentHandleTagPrefix = "S";  // 세그먼트 중간 핸들 (포인트 삽입)

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
                AttachVertexContextMenu(h, shape);
                parent.Children.Add(h);
                _shapeEditHandles.Add(h);
            }

            // 세그먼트 중간 핸들 — 클릭하면 해당 위치에 새 포인트 삽입.
            bool closedShape = shape.Kind == ShapeKind.ClosedSpline || shape.Kind == ShapeKind.Polygon;
            int segCount = closedShape ? pts.Count : pts.Count - 1;
            for (int i = 0; i < segCount; i++)
            {
                var s = CreateSegmentHandle($"{SegmentHandleTagPrefix}{i}");
                PositionSegmentHandle(s, shapeCtrl, shape, i, closedShape);
                parent.Children.Add(s);
                _shapeEditHandles.Add(s);
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
            h.MouseLeftButtonDown -= OnSegmentHandleClick;
        }
        _shapeEditHandles.Clear();
    }

    private void RefreshShapeEditHandlePositions()
    {
        if (_selectedShape is null || _selectedShapeCtrl is null) return;
        bool closedShape = _selectedShape.Kind == ShapeKind.ClosedSpline || _selectedShape.Kind == ShapeKind.Polygon;
        foreach (var h in _shapeEditHandles)
        {
            if (h.Tag is not string tag) continue;
            if (tag.StartsWith(VertexHandleTagPrefix, StringComparison.Ordinal))
            {
                if (!int.TryParse(tag.AsSpan(1), out int idx)) continue;
                if (idx < 0 || idx >= _selectedShape.Points.Count) continue;
                PositionVertexHandle(h, _selectedShapeCtrl, _selectedShape, _selectedShape.Points[idx]);
            }
            else if (tag.StartsWith(SegmentHandleTagPrefix, StringComparison.Ordinal))
            {
                if (!int.TryParse(tag.AsSpan(1), out int segIdx)) continue;
                PositionSegmentHandle(h, _selectedShapeCtrl, _selectedShape, segIdx, closedShape);
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

    // ── 통합 우클릭 위탁 처리 (MainWindow.xaml.cs 의 OnPaperPreviewMouseRightButtonDown 에서 호출) ──

    /// <summary>
    /// 도형 편집 핸들이 우클릭됐을 때 호출된다.
    /// 핸들 종류에 따라 적절한 동작을 수행한다.
    /// </summary>
    internal void OnShapeEditHandleRightClicked(Rectangle handle)
    {
        if (handle.Tag is not string tag) return;

        if (tag.StartsWith(VertexHandleTagPrefix, StringComparison.Ordinal))
        {
            // 정점 핸들: 이미 ContextMenu 가 연결되어 있으므로 직접 연다.
            if (handle.ContextMenu is { } cm)
            {
                cm.PlacementTarget = handle;
                cm.IsOpen = true;
            }
        }
        // 세그먼트 핸들: 우클릭 무시 (좌클릭만 사용)
    }

    /// <summary>정점 핸들에 우클릭 컨텍스트 메뉴("이 점 삭제")를 연결한다.</summary>
    private void AttachVertexContextMenu(Rectangle h, ShapeObject shape)
    {
        var deleteItem = new MenuItem { Header = "이 점 삭제" };
        deleteItem.Click += (_, _) =>
        {
            if (h.Tag is not string tag) return;
            if (!tag.StartsWith(VertexHandleTagPrefix, StringComparison.Ordinal)) return;
            if (!int.TryParse(tag.AsSpan(1), out int idx)) return;
            if (_selectedShape is null || _selectedShapeCtrl is null) return;

            var pts    = _selectedShape.Points;
            bool closed = _selectedShape.Kind == ShapeKind.ClosedSpline
                       || _selectedShape.Kind == ShapeKind.Polygon;
            int minPts = closed ? 3 : 2;
            if (pts.Count <= minPts || idx < 0 || idx >= pts.Count) return;

            pts.RemoveAt(idx);
            _selectedShape.Status = NodeStatus.Modified;
            _viewModel?.MarkDirty();

            HideShapeEditHandles();
            ShowShapeEditHandles(_selectedShapeCtrl, _selectedShape);
            RefreshSelectedShapeVisual();
        };

        h.ContextMenu = new ContextMenu { Items = { deleteItem } };
    }

    private Rectangle CreateSegmentHandle(string tag)
    {
        var h = new Rectangle
        {
            Width            = 7.0,
            Height           = 7.0,
            Fill             = new SolidColorBrush(System.Windows.Media.Color.FromArgb(180, 0x44, 0xAA, 0xFF)),
            Stroke           = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0x29, 0x6B, 0xC9)),
            StrokeThickness  = 1.0,
            Tag              = tag,
            Cursor           = Cursors.Cross,
            IsHitTestVisible = true,
            RenderTransform  = new RotateTransform(45, 3.5, 3.5), // 다이아몬드 모양
        };
        h.MouseLeftButtonDown += OnSegmentHandleClick;
        Panel.SetZIndex(h, 999);
        return h;
    }

    private static void PositionSegmentHandle(
        Rectangle h, FrameworkElement ctrl, ShapeObject shape, int segIdx, bool closed)
    {
        var pts = shape.Points;
        int n   = pts.Count;
        if (n == 0) return;
        int nextIdx = closed ? (segIdx + 1) % n : Math.Min(segIdx + 1, n - 1);

        // 세그먼트 앵커 중간점 (선형 보간).
        double midX = (pts[segIdx].X + pts[nextIdx].X) / 2.0;
        double midY = (pts[segIdx].Y + pts[nextIdx].Y) / 2.0;

        double left = Canvas.GetLeft(ctrl); if (double.IsNaN(left)) left = 0;
        double top  = Canvas.GetTop(ctrl);  if (double.IsNaN(top))  top  = 0;
        double cx   = left + FlowDocumentBuilder.MmToDip(midX);
        double cy   = top  + FlowDocumentBuilder.MmToDip(midY);

        ApplyShapeRotation(ref cx, ref cy, ctrl, shape);

        Canvas.SetLeft(h, cx - 3.5);
        Canvas.SetTop (h, cy - 3.5);
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

    // ── 세그먼트 핸들 / 정점 삭제 ─────────────────────────────────────────

    /// <summary>세그먼트 중간 핸들 클릭 — 해당 세그먼트에 새 앵커 포인트를 삽입한다.</summary>
    private void OnSegmentHandleClick(object sender, MouseButtonEventArgs e)
    {
        if (sender is not Rectangle h) return;
        if (h.Tag is not string tag) return;
        if (!tag.StartsWith(SegmentHandleTagPrefix, StringComparison.Ordinal)) return;
        if (!int.TryParse(tag.AsSpan(1), out int segIdx)) return;
        if (_selectedShape is null || _selectedShapeCtrl is null) return;

        var pts     = _selectedShape.Points;
        int n       = pts.Count;
        bool closed = _selectedShape.Kind == ShapeKind.ClosedSpline || _selectedShape.Kind == ShapeKind.Polygon;
        int nextIdx = closed ? (segIdx + 1) % n : Math.Min(segIdx + 1, n - 1);

        // 스플라인은 실제 곡선 위의 t=0.5 지점, 그 외는 선형 중간점.
        double newX, newY;
        if (_selectedShape.Kind == ShapeKind.Spline || _selectedShape.Kind == ShapeKind.ClosedSpline)
        {
            (newX, newY) = ComputeSplineMidpoint(_selectedShape, segIdx, closed);
        }
        else
        {
            newX = (pts[segIdx].X + pts[nextIdx].X) / 2.0;
            newY = (pts[segIdx].Y + pts[nextIdx].Y) / 2.0;
        }

        // 새 점을 삽입하고 명시적 제어점은 지워 Catmull-Rom 자동 계산으로 되돌린다.
        int insertAt = closed && segIdx == n - 1 ? n : nextIdx;
        pts.Insert(insertAt, new ShapePoint { X = newX, Y = newY });

        // 기존 앵커 제어점도 초기화 (인접 점들의 명시적 제어점 → Catmull-Rom 재계산).
        ClearExplicitControls(pts, segIdx);
        ClearExplicitControls(pts, insertAt);
        ClearExplicitControls(pts, insertAt < pts.Count - 1 ? insertAt + 1 : 0);

        _selectedShape.Status = NodeStatus.Modified;
        _viewModel?.MarkDirty();

        // 핸들 재구성 (인덱스가 바뀌어 기존 핸들은 무효).
        HideShapeEditHandles();
        ShowShapeEditHandles(_selectedShapeCtrl, _selectedShape);
        RefreshSelectedShapeVisual();
        e.Handled = true;
    }

    /// <summary>Catmull-Rom 스플라인의 세그먼트 중간점 (t=0.5)을 mm 단위로 계산.</summary>
    private static (double X, double Y) ComputeSplineMidpoint(ShapeObject shape, int segIdx, bool closed)
    {
        var pts = shape.Points;
        int n   = pts.Count;
        int i1  = segIdx;
        int i2  = closed ? (segIdx + 1) % n : Math.Min(segIdx + 1, n - 1);
        int i0  = closed ? (segIdx - 1 + n) % n : Math.Max(segIdx - 1, 0);
        int i3  = closed ? (segIdx + 2) % n     : Math.Min(segIdx + 2, n - 1);

        double p0x = pts[i0].X, p0y = pts[i0].Y;
        double p1x = pts[i1].X, p1y = pts[i1].Y;
        double p2x = pts[i2].X, p2y = pts[i2].Y;
        double p3x = pts[i3].X, p3y = pts[i3].Y;

        // 명시적 제어점이 있으면 그것을 사용, 없으면 Catmull-Rom.
        double c1x, c1y, c2x, c2y;
        var from = pts[i1];
        var to   = pts[i2];
        if (from.OutCtrlX.HasValue && from.OutCtrlY.HasValue
            && to.InCtrlX.HasValue && to.InCtrlY.HasValue)
        {
            c1x = from.OutCtrlX.Value; c1y = from.OutCtrlY.Value;
            c2x = to.InCtrlX.Value;    c2y = to.InCtrlY.Value;
        }
        else
        {
            c1x = p1x + (p2x - p0x) / 6.0; c1y = p1y + (p2y - p0y) / 6.0;
            c2x = p2x - (p3x - p1x) / 6.0; c2y = p2y - (p3y - p1y) / 6.0;
        }

        // De Casteljau t=0.5
        const double t = 0.5;
        double oneT = 1.0 - t;
        double x = oneT*oneT*oneT*p1x + 3*oneT*oneT*t*c1x + 3*oneT*t*t*c2x + t*t*t*p2x;
        double y = oneT*oneT*oneT*p1y + 3*oneT*oneT*t*c1y + 3*oneT*t*t*c2y + t*t*t*p2y;
        return (x, y);
    }

    private static void ClearExplicitControls(IList<ShapePoint> pts, int idx)
    {
        if (idx < 0 || idx >= pts.Count) return;
        var pt = pts[idx];
        if (pt.OutCtrlX.HasValue || pt.InCtrlX.HasValue)
        {
            pts[idx] = new ShapePoint { X = pt.X, Y = pt.Y };
        }
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
                var pt = shape.Points[i];
                shape.Points[i] = new ShapePoint
                {
                    X        = pt.X - shiftX,
                    Y        = pt.Y - shiftY,
                    OutCtrlX = pt.OutCtrlX.HasValue ? pt.OutCtrlX.Value - shiftX : null,
                    OutCtrlY = pt.OutCtrlY.HasValue ? pt.OutCtrlY.Value - shiftY : null,
                    InCtrlX  = pt.InCtrlX.HasValue  ? pt.InCtrlX.Value  - shiftX : null,
                    InCtrlY  = pt.InCtrlY.HasValue  ? pt.InCtrlY.Value  - shiftY : null,
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

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Xml.Linq;
using WpfShapes = System.Windows.Shapes;
using Path  = System.Windows.Shapes.Path;
using Shape = System.Windows.Shapes.Shape;
using Line  = System.Windows.Shapes.Line;
using Polygon  = System.Windows.Shapes.Polygon;
using Polyline = System.Windows.Shapes.Polyline;
using Rectangle = System.Windows.Shapes.Rectangle;
using Ellipse   = System.Windows.Shapes.Ellipse;

namespace PolyDonky.App.Services;

/// <summary>
/// 인라인 SVG 를 WPF Canvas + Shape 로 렌더링한다 — 기본 도형(rect/circle/ellipse/line/
/// polygon/polyline/path/text + g 그룹/transform) 만 지원한다. 미지원 요소(marker, clipPath,
/// filter, gradient 등) 는 무시되어 결과적으로 placeholder 보다는 정확하게 보여주되
/// 100% 충실 렌더는 아님. 라운드트립용 원본 바이트는 <see cref="PolyDonky.Core.ImageBlock.Data"/>
/// 에 그대로 보존된다 — 본 렌더러는 view-only.
/// </summary>
internal static class SvgRenderer
{
    /// <summary>SVG 바이트를 받아 WPF UIElement(Viewbox 안의 Canvas) 를 만든다. 실패 시 null.</summary>
    public static UIElement? TryRender(byte[] svgUtf8, double targetWidthDip, double targetHeightDip)
    {
        if (svgUtf8 is null || svgUtf8.Length == 0) return null;

        XDocument doc;
        try
        {
            using var ms = new MemoryStream(svgUtf8);
            var settings = new System.Xml.XmlReaderSettings
            {
                DtdProcessing = System.Xml.DtdProcessing.Ignore,
                XmlResolver   = null,
            };
            using var reader = System.Xml.XmlReader.Create(ms, settings);
            doc = XDocument.Load(reader);
        }
        catch { return null; }

        var root = doc.Root;
        if (root is null || root.Name.LocalName != "svg") return null;

        // ── viewport (viewBox 우선, 없으면 width/height) ─────────────
        double vbX = 0, vbY = 0, vbW = 0, vbH = 0;
        if (TryParseViewBox(root.Attribute("viewBox")?.Value, out var pvb))
        {
            vbX = pvb.x; vbY = pvb.y; vbW = pvb.w; vbH = pvb.h;
        }
        if (vbW <= 0) vbW = ParseLengthPx(root.Attribute("width")?.Value)  ?? 300;
        if (vbH <= 0) vbH = ParseLengthPx(root.Attribute("height")?.Value) ?? 150;

        var canvas = new Canvas { Width = vbW, Height = vbH };
        if (vbX != 0 || vbY != 0)
            canvas.RenderTransform = new TranslateTransform(-vbX, -vbY);

        // <defs><marker> 추출 — 화살표 머리 등 line/polyline 의 marker-* 참조용.
        var markers = new Dictionary<string, XElement>(StringComparer.OrdinalIgnoreCase);
        foreach (var defs in root.Descendants().Where(e => e.Name.LocalName == "defs"))
            foreach (var mk in defs.Elements().Where(e => e.Name.LocalName == "marker"))
            {
                var mid = mk.Attribute("id")?.Value;
                if (!string.IsNullOrEmpty(mid)) markers[mid] = mk;
            }

        // 루트 svg 의 style="background:..." / 인라인 fill / border 등 흡수
        var rootCtx = SvgContext.Default.MergedWith(root);
        rootCtx.Markers = markers;
        if (rootCtx.RootBackground is { } bg) canvas.Background = bg;

        foreach (var el in root.Elements())
            RenderElement(el, canvas, rootCtx);

        // 루트 svg 의 style="border:..." — SVG 사양은 아니지만 HTML 인라인 SVG 에서 흔히 쓰임.
        // 주의: Border 를 Viewbox 안에 넣으면 Viewbox 의 ScaleTransform 이 border 픽셀까지 함께
        // 축소해 border 가 RTB/Viewbox 경계에 닿거나 잘리는 문제가 발생한다.
        // 해결: border 를 Viewbox 바깥의 래퍼 Border 에 그리고, Viewbox 에는 테두리 두께만큼
        // 안쪽으로 줄인 크기를 지정해 총 외곽 크기가 targetWidthDip × targetHeightDip 을 유지하도록 함.
        double borderThk  = rootCtx.RootBorderBrush is not null && rootCtx.RootBorderThickness > 0
                            ? rootCtx.RootBorderThickness : 0.0;
        double vbTargetW  = targetWidthDip  > 0 ? Math.Max(1.0, targetWidthDip  - 2 * borderThk) : 0;
        double vbTargetH  = targetHeightDip > 0 ? Math.Max(1.0, targetHeightDip - 2 * borderThk) : 0;

        var viewbox = new Viewbox
        {
            Stretch              = Stretch.Uniform,
            HorizontalAlignment  = HorizontalAlignment.Center,
            VerticalAlignment    = VerticalAlignment.Center,
            Child                = canvas,
        };
        if (vbTargetW > 0) viewbox.Width  = vbTargetW;
        if (vbTargetH > 0) viewbox.Height = vbTargetH;

        if (borderThk > 0 && rootCtx.RootBorderBrush is { } bb)
        {
            return new System.Windows.Controls.Border
            {
                Child           = viewbox,
                BorderBrush     = bb,
                BorderThickness = new Thickness(borderThk),
                Background      = rootCtx.RootBackground,
            };
        }
        return viewbox;
    }

    // ─────────────────────────────────────────────────────────────
    // 스타일 컨텍스트 (상속)
    // ─────────────────────────────────────────────────────────────
    private sealed class SvgContext
    {
        public Brush?     Fill           { get; set; }
        public bool       FillNone       { get; set; }
        public Brush?     Stroke         { get; set; }
        public bool       StrokeNone     { get; set; } = true;     // SVG 기본은 stroke=none
        public double     StrokeWidth    { get; set; } = 1.0;
        public double?    StrokeOpacity  { get; set; }
        public double?    FillOpacity    { get; set; }
        public DoubleCollection? StrokeDashArray { get; set; }
        public string     FontFamily     { get; set; } = "Segoe UI";
        public double     FontSize       { get; set; } = 12;
        public FontWeight FontWeight     { get; set; } = FontWeights.Normal;
        public FontStyle  FontStyle      { get; set; } = FontStyles.Normal;
        public string     TextAnchor     { get; set; } = "start";
        public Brush?     RootBackground { get; set; }
        public Brush?     RootBorderBrush { get; set; }
        public double     RootBorderThickness { get; set; }
        public Dictionary<string, XElement> Markers { get; set; } = new(StringComparer.OrdinalIgnoreCase);

        public static SvgContext Default => new()
        {
            Fill       = Brushes.Black,
            FillNone   = false,
            Stroke     = null,
            StrokeNone = true,
        };

        public SvgContext Clone() => (SvgContext)MemberwiseClone();

        /// <summary>상속한 컨텍스트에 엘리먼트의 attribute / style 을 덮어 새 컨텍스트를 만든다.</summary>
        public SvgContext MergedWith(XElement el)
        {
            var c = Clone();
            var styleProps = ParseInlineStyle(el.Attribute("style")?.Value);

            string? Get(string name)
            {
                if (styleProps.TryGetValue(name, out var v)) return v;
                var a = el.Attribute(name);
                return a?.Value;
            }

            // fill
            var fill = Get("fill");
            if (fill is not null)
            {
                if (string.Equals(fill, "none", StringComparison.OrdinalIgnoreCase))
                { c.FillNone = true; c.Fill = null; }
                else if (TryParseBrush(fill) is { } fb) { c.Fill = fb; c.FillNone = false; }
            }

            // stroke
            var stroke = Get("stroke");
            if (stroke is not null)
            {
                if (string.Equals(stroke, "none", StringComparison.OrdinalIgnoreCase))
                { c.StrokeNone = true; c.Stroke = null; }
                else if (TryParseBrush(stroke) is { } sb) { c.Stroke = sb; c.StrokeNone = false; }
            }

            if (TryParseDouble(Get("stroke-width"), out var sw) && sw >= 0) c.StrokeWidth = sw;
            if (TryParseDouble(Get("stroke-opacity"), out var so)) c.StrokeOpacity = so;
            if (TryParseDouble(Get("fill-opacity"),   out var fo)) c.FillOpacity   = fo;
            var dash = Get("stroke-dasharray");
            if (!string.IsNullOrWhiteSpace(dash) && !string.Equals(dash, "none", StringComparison.OrdinalIgnoreCase))
                c.StrokeDashArray = ParseDashArray(dash!);

            if (Get("font-family") is { } ff && !string.IsNullOrWhiteSpace(ff))
                c.FontFamily = ff.Split(',')[0].Trim().Trim('"', '\'');
            if (TryParseLengthPx(Get("font-size"), out var fs) && fs > 0) c.FontSize = fs;
            var fw = Get("font-weight");
            if (fw is not null) c.FontWeight = ParseFontWeight(fw);
            var fst = Get("font-style");
            if (string.Equals(fst, "italic", StringComparison.OrdinalIgnoreCase)) c.FontStyle = FontStyles.Italic;
            var ta = Get("text-anchor");
            if (!string.IsNullOrWhiteSpace(ta)) c.TextAnchor = ta!;

            // 루트 svg 의 background (style 에서만 — SVG 사양은 아니지만 HTML 인라인 SVG 에서 흔함)
            if (styleProps.TryGetValue("background", out var bg) || styleProps.TryGetValue("background-color", out bg))
            {
                if (TryParseBrush(SplitFirstToken(bg)) is { } bgBrush) c.RootBackground = bgBrush;
            }

            // 루트 svg 의 border (style="border: 1px solid #ddd" 같은 인라인 CSS — SVG 사양은 아니지만 동일).
            // shorthand 또는 border-(width|style|color) 분리 모두 지원.
            if (styleProps.TryGetValue("border", out var bd))
            {
                if (TryParseBorderShorthand(bd, out var bw, out var bbrush))
                { c.RootBorderThickness = bw; c.RootBorderBrush = bbrush; }
            }
            if (styleProps.TryGetValue("border-width", out var bdw) &&
                TryParseLengthPx(bdw, out var bwv) && bwv > 0)
                c.RootBorderThickness = bwv;
            if (styleProps.TryGetValue("border-color", out var bdc) &&
                TryParseBrush(bdc) is { } bdcb)
                c.RootBorderBrush = bdcb;

            return c;
        }
    }

    /// <summary>"1px solid #ddd" 같은 border shorthand 에서 width 와 color 를 추출한다.</summary>
    private static bool TryParseBorderShorthand(string value, out double widthPx, out Brush brush)
    {
        widthPx = 1; brush = Brushes.Black;
        if (string.IsNullOrWhiteSpace(value)) return false;
        var tokens = value.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
        bool gotWidth = false, gotColor = false;
        foreach (var t in tokens)
        {
            if (!gotWidth && TryParseLengthPx(t, out var w) && w >= 0)  { widthPx = w; gotWidth = true; continue; }
            if (string.Equals(t, "solid", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(t, "dashed", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(t, "dotted", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(t, "double", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(t, "none", StringComparison.OrdinalIgnoreCase))
                continue;
            if (TryParseBrush(t) is { } b) { brush = b; gotColor = true; }
        }
        return gotWidth || gotColor;
    }

    // ─────────────────────────────────────────────────────────────
    // 엘리먼트 렌더링
    // ─────────────────────────────────────────────────────────────
    private static void RenderElement(XElement el, Canvas parent, SvgContext inherited)
    {
        var name = el.Name.LocalName;
        var ctx  = inherited.MergedWith(el);

        switch (name)
        {
            case "g":
            {
                var sub = new Canvas();
                ApplyTransform(sub, el.Attribute("transform")?.Value);
                foreach (var c in el.Elements()) RenderElement(c, sub, ctx);
                parent.Children.Add(sub);
                break;
            }
            case "rect":
                RenderRect(el, parent, ctx); break;
            case "circle":
                RenderCircle(el, parent, ctx); break;
            case "ellipse":
                RenderEllipse(el, parent, ctx); break;
            case "line":
                RenderLine(el, parent, ctx); break;
            case "polygon":
                RenderPolygon(el, parent, ctx, closed: true); break;
            case "polyline":
                RenderPolygon(el, parent, ctx, closed: false); break;
            case "path":
                RenderPath(el, parent, ctx); break;
            case "text":
                RenderText(el, parent, ctx); break;
            case "defs":
            case "title":
            case "desc":
            case "metadata":
                break; // 의도적으로 무시
            default:
                // 미지원 요소는 자식만 재귀
                foreach (var c in el.Elements()) RenderElement(c, parent, ctx);
                break;
        }
    }

    private static void RenderRect(XElement el, Canvas parent, SvgContext ctx)
    {
        var x = AttrDouble(el, "x");
        var y = AttrDouble(el, "y");
        var w = AttrDouble(el, "width");
        var h = AttrDouble(el, "height");
        if (w <= 0 || h <= 0) return;
        var rx = AttrDouble(el, "rx");
        var ry = AttrDouble(el, "ry");
        if (rx > 0 && ry <= 0) ry = rx;
        if (ry > 0 && rx <= 0) rx = ry;

        var rect = new Rectangle
        {
            Width  = w,
            Height = h,
            RadiusX = rx,
            RadiusY = ry,
        };
        ApplyPaint(rect, ctx);
        ApplyTransform(rect, el.Attribute("transform")?.Value);
        Canvas.SetLeft(rect, x);
        Canvas.SetTop(rect,  y);
        parent.Children.Add(rect);
    }

    private static void RenderCircle(XElement el, Canvas parent, SvgContext ctx)
    {
        var cx = AttrDouble(el, "cx");
        var cy = AttrDouble(el, "cy");
        var r  = AttrDouble(el, "r");
        if (r <= 0) return;
        var ell = new Ellipse { Width = 2 * r, Height = 2 * r };
        ApplyPaint(ell, ctx);
        ApplyTransform(ell, el.Attribute("transform")?.Value);
        Canvas.SetLeft(ell, cx - r);
        Canvas.SetTop(ell,  cy - r);
        parent.Children.Add(ell);
    }

    private static void RenderEllipse(XElement el, Canvas parent, SvgContext ctx)
    {
        var cx = AttrDouble(el, "cx");
        var cy = AttrDouble(el, "cy");
        var rx = AttrDouble(el, "rx");
        var ry = AttrDouble(el, "ry");
        if (rx <= 0 || ry <= 0) return;
        var ell = new Ellipse { Width = 2 * rx, Height = 2 * ry };
        ApplyPaint(ell, ctx);
        ApplyTransform(ell, el.Attribute("transform")?.Value);
        Canvas.SetLeft(ell, cx - rx);
        Canvas.SetTop(ell,  cy - ry);
        parent.Children.Add(ell);
    }

    private static void RenderLine(XElement el, Canvas parent, SvgContext ctx)
    {
        var x1 = AttrDouble(el, "x1");
        var y1 = AttrDouble(el, "y1");
        var x2 = AttrDouble(el, "x2");
        var y2 = AttrDouble(el, "y2");
        var line = new Line { X1 = x1, Y1 = y1, X2 = x2, Y2 = y2 };
        // Line 은 fill 무시 — stroke 만 적용.
        ApplyStroke(line, ctx);
        ApplyTransform(line, el.Attribute("transform")?.Value);
        parent.Children.Add(line);

        // 화살표 머리 등 marker-* 참조 처리.
        var markerEnd   = ResolveMarkerRef(GetAttrOrStyle(el, "marker-end"),   ctx);
        var markerStart = ResolveMarkerRef(GetAttrOrStyle(el, "marker-start"), ctx);
        if (markerEnd is not null)
        {
            double angleEnd = Math.Atan2(y2 - y1, x2 - x1) * 180.0 / Math.PI;
            RenderMarker(parent, markerEnd, x2, y2, angleEnd, ctx);
        }
        if (markerStart is not null)
        {
            double angleStart = Math.Atan2(y1 - y2, x1 - x2) * 180.0 / Math.PI;
            RenderMarker(parent, markerStart, x1, y1, angleStart, ctx);
        }
    }

    private static string? GetAttrOrStyle(XElement el, string name)
    {
        var v = el.Attribute(name)?.Value;
        if (!string.IsNullOrEmpty(v)) return v;
        var style = el.Attribute("style")?.Value;
        if (string.IsNullOrEmpty(style)) return null;
        var props = ParseInlineStyle(style);
        return props.TryGetValue(name, out var sv) ? sv : null;
    }

    /// <summary>"url(#id)" 형식 의 marker 참조를 해소해 marker XElement 를 돌려준다. 미해소 시 null.</summary>
    private static XElement? ResolveMarkerRef(string? markerRef, SvgContext ctx)
    {
        if (string.IsNullOrWhiteSpace(markerRef)) return null;
        markerRef = markerRef.Trim();
        if (!markerRef.StartsWith("url(", StringComparison.OrdinalIgnoreCase)) return null;
        var inside = markerRef.Substring(4, markerRef.IndexOf(')') - 4).Trim().Trim('"', '\'');
        if (inside.StartsWith('#')) inside = inside.Substring(1);
        return ctx.Markers.TryGetValue(inside, out var mk) ? mk : null;
    }

    private static void RenderMarker(Canvas parent, XElement marker, double atX, double atY, double angleDeg, SvgContext ctx)
    {
        // Marker 좌표계 (viewBox) — 없으면 markerWidth/Height 와 동일.
        double mw = double.TryParse(marker.Attribute("markerWidth")?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var mwv) ? mwv : 3;
        double mh = double.TryParse(marker.Attribute("markerHeight")?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var mhv) ? mhv : 3;
        double vbW = mw, vbH = mh, vbX = 0, vbY = 0;
        if (TryParseViewBox(marker.Attribute("viewBox")?.Value, out var mvb))
        { vbX = mvb.x; vbY = mvb.y; vbW = mvb.w; vbH = mvb.h; }
        double refX = double.TryParse(marker.Attribute("refX")?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var rxv) ? rxv : 0;
        double refY = double.TryParse(marker.Attribute("refY")?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var ryv) ? ryv : 0;
        bool orientAuto = string.Equals(marker.Attribute("orient")?.Value, "auto", StringComparison.OrdinalIgnoreCase);
        if (!orientAuto && double.TryParse(marker.Attribute("orient")?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var fixedAngle))
            angleDeg = fixedAngle;

        // marker 자식을 marker 좌표계 그대로 그려주는 내부 Canvas.
        var inner = new Canvas { Width = vbW, Height = vbH };
        if (vbX != 0 || vbY != 0)
            inner.RenderTransform = new TranslateTransform(-vbX, -vbY);
        var markerCtx = SvgContext.Default;
        markerCtx.Markers = ctx.Markers; // 중첩 marker 참조 가능
        foreach (var c in marker.Elements())
            RenderElement(c, inner, markerCtx);

        // 변환: 1) refX,refY 를 원점으로, 2) markerWidth/Height 비율로 스케일, 3) 라인 방향으로 회전, 4) 끝점으로 평행이동.
        var tg = new TransformGroup();
        tg.Children.Add(new TranslateTransform(-refX, -refY));
        tg.Children.Add(new ScaleTransform(mw / vbW, mh / vbH));
        tg.Children.Add(new RotateTransform(angleDeg));
        tg.Children.Add(new TranslateTransform(atX, atY));

        var wrapper = new Canvas { Width = mw, Height = mh, RenderTransform = tg };
        wrapper.Children.Add(inner);
        parent.Children.Add(wrapper);
    }

    private static void RenderPolygon(XElement el, Canvas parent, SvgContext ctx, bool closed)
    {
        var pts = ParsePoints(el.Attribute("points")?.Value);
        if (pts.Count < 2) return;
        Shape shape = closed
            ? new Polygon  { Points = new PointCollection(pts) }
            : (Shape)new Polyline { Points = new PointCollection(pts) };
        if (closed) ApplyPaint(shape, ctx);
        else        ApplyPaintNoFillByDefault(shape, ctx);
        ApplyTransform(shape, el.Attribute("transform")?.Value);
        parent.Children.Add(shape);

        if (!closed)
        {
            // polyline 끝/시작 marker.
            var markerEnd   = ResolveMarkerRef(GetAttrOrStyle(el, "marker-end"),   ctx);
            var markerStart = ResolveMarkerRef(GetAttrOrStyle(el, "marker-start"), ctx);
            if (markerEnd is not null && pts.Count >= 2)
            {
                var p1 = pts[^2]; var p2 = pts[^1];
                double a = Math.Atan2(p2.Y - p1.Y, p2.X - p1.X) * 180.0 / Math.PI;
                RenderMarker(parent, markerEnd, p2.X, p2.Y, a, ctx);
            }
            if (markerStart is not null && pts.Count >= 2)
            {
                var p1 = pts[1]; var p2 = pts[0];
                double a = Math.Atan2(p2.Y - p1.Y, p2.X - p1.X) * 180.0 / Math.PI;
                RenderMarker(parent, markerStart, p2.X, p2.Y, a, ctx);
            }
        }
    }

    private static void RenderPath(XElement el, Canvas parent, SvgContext ctx)
    {
        var d = el.Attribute("d")?.Value;
        if (string.IsNullOrWhiteSpace(d)) return;
        Geometry geom;
        try { geom = Geometry.Parse(d); }
        catch { return; }
        var path = new Path { Data = geom };
        ApplyPaintNoFillByDefault(path, ctx);
        ApplyTransform(path, el.Attribute("transform")?.Value);
        parent.Children.Add(path);

        // path 끝 marker — d 문자열에서 마지막 두 좌표 쌍을 추출해 끝점 + 접선 방향 근사.
        var markerEnd   = ResolveMarkerRef(GetAttrOrStyle(el, "marker-end"),   ctx);
        var markerStart = ResolveMarkerRef(GetAttrOrStyle(el, "marker-start"), ctx);
        if (markerEnd is not null || markerStart is not null)
        {
            var matches = System.Text.RegularExpressions.Regex.Matches(d!, @"-?\d+(?:\.\d+)?");
            if (matches.Count >= 4)
            {
                var nums = matches
                    .Select(m => double.TryParse(m.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? v : 0)
                    .ToList();
                if (markerEnd is not null)
                {
                    double endX = nums[^2], endY = nums[^1];
                    double prevX = nums[^4], prevY = nums[^3];
                    double a = Math.Atan2(endY - prevY, endX - prevX) * 180.0 / Math.PI;
                    RenderMarker(parent, markerEnd, endX, endY, a, ctx);
                }
                if (markerStart is not null)
                {
                    double startX = nums[0], startY = nums[1];
                    double nextX = nums[2], nextY = nums[3];
                    double a = Math.Atan2(startY - nextY, startX - nextX) * 180.0 / Math.PI;
                    RenderMarker(parent, markerStart, startX, startY, a, ctx);
                }
            }
        }
    }

    private static void RenderText(XElement el, Canvas parent, SvgContext ctx)
    {
        var x = AttrDouble(el, "x");
        var y = AttrDouble(el, "y");

        // 자식 tspan 을 평탄화한 단순 텍스트
        var text = string.Concat(el.DescendantNodes().OfType<XText>().Select(t => t.Value)).Trim();
        if (string.IsNullOrEmpty(text)) text = (el.Value ?? "").Trim();
        if (string.IsNullOrEmpty(text)) return;

        var typeface = new Typeface(new FontFamily(ctx.FontFamily), ctx.FontStyle, ctx.FontWeight, FontStretches.Normal);
        var ft = new FormattedText(
            text,
            CultureInfo.CurrentCulture,
            FlowDirection.LeftToRight,
            typeface,
            ctx.FontSize,
            Brushes.Black,
            1.0);

        // text-anchor 처리
        double left = x;
        if (ctx.TextAnchor.Equals("middle", StringComparison.OrdinalIgnoreCase)) left = x - ft.Width / 2.0;
        else if (ctx.TextAnchor.Equals("end", StringComparison.OrdinalIgnoreCase)) left = x - ft.Width;

        // SVG y 는 baseline. WPF Canvas.Top 은 상단. baseline ≒ ft.Baseline.
        double top = y - ft.Baseline;

        var tb = new TextBlock
        {
            Text       = text,
            FontFamily = new FontFamily(ctx.FontFamily),
            FontSize   = ctx.FontSize,
            FontWeight = ctx.FontWeight,
            FontStyle  = ctx.FontStyle,
            Foreground = ResolveFillBrush(ctx) ?? Brushes.Black,
        };
        Canvas.SetLeft(tb, left);
        Canvas.SetTop(tb,  top);
        ApplyTransform(tb, el.Attribute("transform")?.Value);
        parent.Children.Add(tb);
    }

    // ─────────────────────────────────────────────────────────────
    // Paint / Stroke
    // ─────────────────────────────────────────────────────────────
    private static void ApplyPaint(Shape s, SvgContext ctx)
    {
        s.Fill = ResolveFillBrush(ctx);
        ApplyStroke(s, ctx);
    }

    /// <summary>polyline / path 처럼 SVG 기본 fill 이 black 이면 어색한 요소용.</summary>
    private static void ApplyPaintNoFillByDefault(Shape s, SvgContext ctx)
    {
        // fill 속성을 명시하지 않은 polyline/path 는 fill="none" 처리하지 않으면
        // 검정으로 채워져 직선·곡선이 안 보인다 — SVG 표준은 black fill 이지만
        // 본 렌더러에서는 "fill 속성을 명시했을 때만 fill 적용" 규칙을 쓴다.
        if (ctx.FillNone) s.Fill = null;
        else              s.Fill = ResolveFillBrush(ctx);
        ApplyStroke(s, ctx);
    }

    private static Brush? ResolveFillBrush(SvgContext ctx)
    {
        if (ctx.FillNone) return null;
        var b = ctx.Fill;
        if (b is null) return null;
        if (ctx.FillOpacity is { } op && op < 1.0)
        {
            var clone = b.Clone();
            clone.Opacity = Math.Clamp(op, 0, 1);
            return clone;
        }
        return b;
    }

    private static void ApplyStroke(Shape s, SvgContext ctx)
    {
        if (ctx.StrokeNone || ctx.Stroke is null) { s.Stroke = null; return; }
        var b = ctx.Stroke;
        if (ctx.StrokeOpacity is { } op && op < 1.0)
        {
            b = b.Clone();
            b.Opacity = Math.Clamp(op, 0, 1);
        }
        s.Stroke          = b;
        s.StrokeThickness = ctx.StrokeWidth;
        if (ctx.StrokeDashArray is { } da)
            s.StrokeDashArray = da;
    }

    // ─────────────────────────────────────────────────────────────
    // transform
    // ─────────────────────────────────────────────────────────────
    private static void ApplyTransform(UIElement el, string? transformStr)
    {
        if (string.IsNullOrWhiteSpace(transformStr)) return;
        var t = ParseTransform(transformStr);
        if (t is not null) el.RenderTransform = t;
    }

    private static Transform? ParseTransform(string s)
    {
        var group = new TransformGroup();
        int i = 0;
        while (i < s.Length)
        {
            while (i < s.Length && (char.IsWhiteSpace(s[i]) || s[i] == ',')) i++;
            if (i >= s.Length) break;
            int nameStart = i;
            while (i < s.Length && (char.IsLetter(s[i]) || s[i] == '-')) i++;
            var fname = s.Substring(nameStart, i - nameStart);
            while (i < s.Length && char.IsWhiteSpace(s[i])) i++;
            if (i >= s.Length || s[i] != '(') break;
            i++; // (
            int argsStart = i;
            while (i < s.Length && s[i] != ')') i++;
            if (i >= s.Length) break;
            var argsStr = s.Substring(argsStart, i - argsStart);
            i++; // )

            var nums = argsStr.Split(new[] { ' ', ',', '\t', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries)
                              .Select(p => double.TryParse(p, NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? v : 0.0)
                              .ToArray();

            switch (fname)
            {
                case "translate":
                    group.Children.Add(new TranslateTransform(
                        nums.Length > 0 ? nums[0] : 0,
                        nums.Length > 1 ? nums[1] : 0));
                    break;
                case "scale":
                    group.Children.Add(new ScaleTransform(
                        nums.Length > 0 ? nums[0] : 1,
                        nums.Length > 1 ? nums[1] : (nums.Length > 0 ? nums[0] : 1)));
                    break;
                case "rotate":
                    if (nums.Length >= 3) group.Children.Add(new RotateTransform(nums[0], nums[1], nums[2]));
                    else if (nums.Length >= 1) group.Children.Add(new RotateTransform(nums[0]));
                    break;
                case "matrix":
                    if (nums.Length >= 6)
                        group.Children.Add(new MatrixTransform(new Matrix(nums[0], nums[1], nums[2], nums[3], nums[4], nums[5])));
                    break;
                case "skewX":
                    if (nums.Length >= 1) group.Children.Add(new SkewTransform(nums[0], 0));
                    break;
                case "skewY":
                    if (nums.Length >= 1) group.Children.Add(new SkewTransform(0, nums[0]));
                    break;
            }
        }
        if (group.Children.Count == 0) return null;
        if (group.Children.Count == 1) return group.Children[0];
        return group;
    }

    // ─────────────────────────────────────────────────────────────
    // 파싱 헬퍼
    // ─────────────────────────────────────────────────────────────
    private static double AttrDouble(XElement el, string name)
    {
        var v = el.Attribute(name)?.Value;
        return TryParseLengthPx(v, out var d) ? d : 0;
    }

    private static bool TryParseDouble(string? s, out double v)
    {
        v = 0;
        if (string.IsNullOrWhiteSpace(s)) return false;
        return double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out v);
    }

    /// <summary>"12", "12px", "1.5em"(em → 1em=16px 추정) 등을 px(=DIP) 로 변환.</summary>
    private static bool TryParseLengthPx(string? s, out double px)
    {
        px = 0;
        if (string.IsNullOrWhiteSpace(s)) return false;
        s = s!.Trim();
        if (s.EndsWith("px", StringComparison.OrdinalIgnoreCase))
            return double.TryParse(s.AsSpan(0, s.Length - 2), NumberStyles.Float, CultureInfo.InvariantCulture, out px);
        if (s.EndsWith("pt", StringComparison.OrdinalIgnoreCase) &&
            double.TryParse(s.AsSpan(0, s.Length - 2), NumberStyles.Float, CultureInfo.InvariantCulture, out var pt))
        { px = pt * 96.0 / 72.0; return true; }
        if (s.EndsWith("em", StringComparison.OrdinalIgnoreCase) &&
            double.TryParse(s.AsSpan(0, s.Length - 2), NumberStyles.Float, CultureInfo.InvariantCulture, out var em))
        { px = em * 16.0; return true; }
        if (s.EndsWith("%", StringComparison.Ordinal)) return false; // 컨텍스트 불명
        return double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out px);
    }

    private static double? ParseLengthPx(string? s)
        => TryParseLengthPx(s, out var v) ? v : null;

    private static bool TryParseViewBox(string? s, out (double x, double y, double w, double h) vb)
    {
        vb = (0, 0, 0, 0);
        if (string.IsNullOrWhiteSpace(s)) return false;
        var parts = s!.Split(new[] { ' ', ',', '\t' }, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length != 4) return false;
        if (!double.TryParse(parts[0], NumberStyles.Float, CultureInfo.InvariantCulture, out var x)) return false;
        if (!double.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture, out var y)) return false;
        if (!double.TryParse(parts[2], NumberStyles.Float, CultureInfo.InvariantCulture, out var w)) return false;
        if (!double.TryParse(parts[3], NumberStyles.Float, CultureInfo.InvariantCulture, out var h)) return false;
        vb = (x, y, w, h);
        return true;
    }

    private static List<Point> ParsePoints(string? s)
    {
        var list = new List<Point>();
        if (string.IsNullOrWhiteSpace(s)) return list;
        var nums = s!.Split(new[] { ' ', ',', '\t', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
        for (int i = 0; i + 1 < nums.Length; i += 2)
        {
            if (double.TryParse(nums[i],   NumberStyles.Float, CultureInfo.InvariantCulture, out var x) &&
                double.TryParse(nums[i+1], NumberStyles.Float, CultureInfo.InvariantCulture, out var y))
                list.Add(new Point(x, y));
        }
        return list;
    }

    private static DoubleCollection ParseDashArray(string s)
    {
        var dc = new DoubleCollection();
        var parts = s.Split(new[] { ' ', ',' }, StringSplitOptions.RemoveEmptyEntries);
        foreach (var p in parts)
            if (double.TryParse(p, NumberStyles.Float, CultureInfo.InvariantCulture, out var v) && v >= 0)
                dc.Add(v);
        return dc;
    }

    private static Brush? TryParseBrush(string? s)
    {
        if (string.IsNullOrWhiteSpace(s)) return null;
        s = s!.Trim();
        if (string.Equals(s, "none", StringComparison.OrdinalIgnoreCase)) return null;
        if (string.Equals(s, "currentColor", StringComparison.OrdinalIgnoreCase)) return Brushes.Black;
        // url(#...) gradient/pattern 은 미지원 — 무시하고 기본 검정
        if (s.StartsWith("url(", StringComparison.OrdinalIgnoreCase)) return Brushes.Black;
        try
        {
            var c = (Color)ColorConverter.ConvertFromString(s)!;
            return new SolidColorBrush(c);
        }
        catch { return null; }
    }

    private static Dictionary<string, string> ParseInlineStyle(string? style)
    {
        var d = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (string.IsNullOrWhiteSpace(style)) return d;
        foreach (var decl in style!.Split(';', StringSplitOptions.RemoveEmptyEntries))
        {
            var idx = decl.IndexOf(':');
            if (idx <= 0) continue;
            var k = decl.Substring(0, idx).Trim();
            var v = decl.Substring(idx + 1).Trim();
            if (k.Length > 0) d[k] = v;
        }
        return d;
    }

    private static FontWeight ParseFontWeight(string s)
    {
        s = s.Trim();
        if (string.Equals(s, "bold", StringComparison.OrdinalIgnoreCase))   return FontWeights.Bold;
        if (string.Equals(s, "normal", StringComparison.OrdinalIgnoreCase)) return FontWeights.Normal;
        if (int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n))
        {
            if (n >= 700) return FontWeights.Bold;
            if (n >= 600) return FontWeights.SemiBold;
            if (n >= 500) return FontWeights.Medium;
            if (n <= 300) return FontWeights.Light;
        }
        return FontWeights.Normal;
    }

    private static string SplitFirstToken(string? s)
    {
        if (string.IsNullOrWhiteSpace(s)) return "";
        var sp = s!.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
        return sp.Length > 0 ? sp[0] : "";
    }
}

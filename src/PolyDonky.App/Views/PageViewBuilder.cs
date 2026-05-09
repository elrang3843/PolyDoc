using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using PolyDonky.App.Pagination;
using PolyDonky.App.Services;
using PolyDonky.Core;
using WpfColor = System.Windows.Media.Color;
using WpfDoc = System.Windows.Documents;

namespace PolyDonky.App.Views;

/// <summary>
/// 편집창과 미리보기가 공유하는 페이지 캔버스 렌더링 유틸리티.
/// </summary>
public static class PageViewBuilder
{
    /// <summary>
    /// Canvas 에 페이지 배경(흰 Border + 선택적 그림자·여백 가이드·페이지 라벨)을 채운다.
    /// </summary>
    public static void BuildPageFrames(
        Canvas        target,
        PageGeometry  geo,
        int           pageCount,
        PageSettings? pageSettings = null,
        Brush?        pageBg       = null,
        bool          showShadow   = true,
        bool          showGuides   = true,
        bool          showLabels   = true)
    {
        target.Children.Clear();
        pageBg ??= Brushes.White;

        for (int i = 0; i < pageCount; i++)
        {
            double topY = i * geo.PageStrideDip;

            var border = new Border
            {
                Width            = geo.PageWidthDip,
                Height           = geo.PageHeightDip,
                Background       = pageBg,
                BorderBrush      = new SolidColorBrush(WpfColor.FromRgb(0xCC, 0xCC, 0xCC)),
                BorderThickness  = new Thickness(0.5),
                IsHitTestVisible = false,
            };
            if (showShadow)
                border.Effect = new System.Windows.Media.Effects.DropShadowEffect
                {
                    BlurRadius  = 18,
                    ShadowDepth = 6,
                    Opacity     = 0.38,
                    Direction   = 270,
                    Color       = Colors.Black,
                };
            Canvas.SetLeft(border, 0);
            Canvas.SetTop (border, topY);
            target.Children.Add(border);

            if (showGuides && (pageSettings?.ShowMarginGuides ?? true))
            {
                var guide = new System.Windows.Shapes.Rectangle
                {
                    Width           = geo.PageWidthDip  - geo.PadLeftDip - geo.PadRightDip,
                    Height          = geo.PageHeightDip - geo.PadTopDip  - geo.PadBottomDip,
                    Stroke          = new SolidColorBrush(WpfColor.FromArgb(0x66, 0x00, 0x78, 0xD4)),
                    StrokeThickness = 0.7,
                    StrokeDashArray = new DoubleCollection { 4, 3 },
                    Fill            = Brushes.Transparent,
                    IsHitTestVisible = false,
                };
                Canvas.SetLeft(guide, geo.PadLeftDip);
                Canvas.SetTop (guide, topY + geo.PadTopDip);
                target.Children.Add(guide);
            }

            if (showLabels)
            {
                var label = new System.Windows.Controls.TextBlock
                {
                    Text             = $"{i + 1}페이지",
                    FontSize         = 10,
                    Foreground       = new SolidColorBrush(WpfColor.FromArgb(0xA0, 0x50, 0x50, 0xB4)),
                    IsHitTestVisible = false,
                };
                Canvas.SetLeft(label, 6);
                Canvas.SetTop (label, topY + 2);
                target.Children.Add(label);

                // 디버그 정보 — 페이지 높이/본문 가용 높이/여백 (육안 검증용).
                // 좌측 여백 안쪽 상단에 세로로 stacked. 페이지네이션 측정값과 실제 렌더 높이를
                // 비교해 클리핑·공백 원인을 추적할 때 사용한다.
                double bodyH = geo.PageHeightDip - geo.PadTopDip - geo.PadBottomDip;
                double pageHmm = FlowDocumentBuilder.DipToMm(geo.PageHeightDip);
                double bodyHmm = FlowDocumentBuilder.DipToMm(bodyH);
                double padTopMm    = FlowDocumentBuilder.DipToMm(geo.PadTopDip);
                double padBottomMm = FlowDocumentBuilder.DipToMm(geo.PadBottomDip);
                var debugLabel = new System.Windows.Controls.TextBlock
                {
                    Text =
                        $"page H: {geo.PageHeightDip:F1} DIP ({pageHmm:F1} mm)\n" +
                        $"body H: {bodyH:F1} DIP ({bodyHmm:F1} mm)\n" +
                        $"pad ↑: {geo.PadTopDip:F1} DIP ({padTopMm:F1} mm)\n" +
                        $"pad ↓: {geo.PadBottomDip:F1} DIP ({padBottomMm:F1} mm)",
                    FontSize         = 9,
                    FontFamily       = new FontFamily("Consolas, Cascadia Mono, monospace"),
                    Foreground       = new SolidColorBrush(WpfColor.FromArgb(0xC0, 0xB4, 0x40, 0x40)),
                    Background       = new SolidColorBrush(WpfColor.FromArgb(0x30, 0xFF, 0xFF, 0x80)),
                    Padding          = new Thickness(3, 1, 3, 1),
                    IsHitTestVisible = false,
                };
                Canvas.SetLeft(debugLabel, 6);
                Canvas.SetTop (debugLabel, topY + 18);
                target.Children.Add(debugLabel);
            }
        }
    }

    /// <summary>
    /// 오버레이 Canvas 에 적용할 페이지 영역 클립 Geometry 를 반환한다.
    /// 페이지 간 갭(InterPageGapDip)은 클립에서 제외 — 오버레이가 갭 영역으로 넘어가지 않는다.
    /// </summary>
    public static Geometry BuildPageClipGeometry(PageGeometry geo, int pageCount)
    {
        var group = new GeometryGroup();
        for (int i = 0; i < pageCount; i++)
        {
            double y = i * geo.PageStrideDip;
            group.Children.Add(new RectangleGeometry(new Rect(0, y, geo.PageWidthDip, geo.PageHeightDip)));
        }
        group.Freeze();
        return group;
    }

    /// <summary>
    /// <paramref name="paginated"/> 의 OverlayBlocks 로 오버레이·언더레이 캔버스를 채운다.
    /// </summary>
    public static void PopulateOverlayCanvases(
        PaginatedDocument paginated,
        PageGeometry      geo,
        Canvas overlayShape,  Canvas underlayShape,
        Canvas overlayImage,  Canvas underlayImage,
        Canvas overlayTable,  Canvas underlayTable,
        Canvas floatingCanvas)
    {
        overlayShape.Children.Clear();  underlayShape.Children.Clear();
        overlayImage.Children.Clear();  underlayImage.Children.Clear();
        overlayTable.Children.Clear();  underlayTable.Children.Clear();
        floatingCanvas.Children.Clear();

        foreach (var page in paginated.Pages)
        {
            foreach (var oop in page.OverlayBlocks)
            {
                switch (oop.Source)
                {
                    case TextBoxObject tb:
                    {
                        var ov = new TextBoxOverlay(tb)
                        {
                            Width            = FlowDocumentBuilder.MmToDip(tb.WidthMm),
                            Height           = FlowDocumentBuilder.MmToDip(tb.HeightMm),
                            IsHitTestVisible = false,
                        };
                        ov.Tag = tb;
                        PlaceAt(ov, geo, oop.AnchorPageIndex, oop.XMm, oop.YMm);
                        floatingCanvas.Children.Add(ov);
                        break;
                    }

                    case ShapeObject shp:
                    {
                        var ctrl = FlowDocumentBuilder.BuildOverlayShapeControl(shp);
                        ctrl.Tag = shp;
                        PlaceAt(ctrl, geo, oop.AnchorPageIndex, oop.XMm, oop.YMm);
                        var behind = shp.WrapMode == ImageWrapMode.BehindText;
                        (behind ? underlayShape : overlayShape).Children.Add(ctrl);
                        break;
                    }

                    case ImageBlock img:
                    {
                        var ctrl = FlowDocumentBuilder.BuildOverlayImageControl(img);
                        if (ctrl is null) break;
                        ctrl.Tag = img;
                        PlaceAt(ctrl, geo, oop.AnchorPageIndex, oop.XMm, oop.YMm);
                        var behind = img.WrapMode == ImageWrapMode.BehindText;
                        (behind ? underlayImage : overlayImage).Children.Add(ctrl);
                        break;
                    }

                    case Table tbl:
                    {
                        var ctrl = FlowDocumentBuilder.BuildOverlayTableControl(tbl);
                        if (ctrl is null) break;
                        ctrl.Tag = tbl;
                        PlaceAt(ctrl, geo, oop.AnchorPageIndex, oop.XMm, oop.YMm);
                        var behind = tbl.WrapMode == TableWrapMode.BehindText;
                        (behind ? underlayTable : overlayTable).Children.Add(ctrl);
                        break;
                    }
                }
            }
        }
    }

    private static void PlaceAt(FrameworkElement ctrl, PageGeometry geo, int pageIndex, double xMm, double yMm)
    {
        var (xDip, yDip) = geo.ToAbsoluteDip(pageIndex, xMm, yMm);
        Canvas.SetLeft(ctrl, xDip);
        Canvas.SetTop (ctrl, yDip);
    }

    /// <summary>
    /// 워터마크를 지정된 Canvas 에 렌더링한다.
    /// </summary>
    public static void BuildWatermarkLayer(
        Canvas               target,
        WatermarkSettings?   watermark,
        PageGeometry         geo,
        int                  pageCount)
    {
        target.Children.Clear();
        if (watermark is null || !watermark.Enabled || string.IsNullOrWhiteSpace(watermark.Text))
            return;

        for (int i = 0; i < pageCount; i++)
        {
            double pageCenterY = i * geo.PageStrideDip + geo.PageHeightDip / 2;
            double pageCenterX = geo.PageWidthDip / 2;

            var tb = new TextBlock
            {
                Text                = watermark.Text,
                FontSize            = watermark.FontSize,
                FontWeight          = FontWeights.Bold,
                Foreground          = ParseColorBrush(watermark.Color, watermark.Opacity),
                TextAlignment       = TextAlignment.Center,
                IsHitTestVisible    = false,
                LayoutTransform     = new RotateTransform(watermark.Rotation),
            };
            tb.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
            var width = tb.DesiredSize.Width;
            var height = tb.DesiredSize.Height;

            Canvas.SetLeft(tb, pageCenterX - width / 2);
            Canvas.SetTop(tb, pageCenterY - height / 2);
            target.Children.Add(tb);
        }
    }

    private static Brush ParseColorBrush(string colorHex, double opacity)
    {
        try
        {
            var color = (WpfColor)System.Windows.Media.ColorConverter.ConvertFromString(colorHex)!;
            return new SolidColorBrush(color) { Opacity = opacity };
        }
        catch { }
        return new SolidColorBrush(Colors.Gray) { Opacity = opacity };
    }

    // ── 머리말·꼬리말 렌더링 ─────────────────────────────────────────────────
    /// <summary>
    /// 모든 페이지에 머리말·꼬리말을 렌더링한다. 좌/가운데/우 3분할.
    /// 토큰(<c>{PAGE}</c>, <c>{NUMPAGES}</c>, <c>{TITLE}</c> …) 은 페이지마다 치환된다.
    /// 편집 차단(<see cref="UIElement.IsHitTestVisible"/> = false) — 편집은 PageFormatWindow 가 담당.
    /// 1차 사이클: <c>DifferentFirstPage</c>/<c>DifferentOddEven</c> 모델 확장 전이라 모든 페이지 동일.
    /// </summary>
    public static void BuildHeaderFooterLayer(
        Canvas         target,
        PageSettings   page,
        PageGeometry   geo,
        int            pageCount,
        DocumentMetadata? metadata = null,
        string?        fileNameWithoutExt = null,
        DateTime?      now = null)
    {
        ArgumentNullException.ThrowIfNull(target);
        ArgumentNullException.ThrowIfNull(page);
        ArgumentNullException.ThrowIfNull(geo);

        target.Children.Clear();
        if ((page.Header is null || page.Header.IsEmpty) && (page.Footer is null || page.Footer.IsEmpty)) return;

        var resolvedNow = now ?? DateTime.Now;
        double headerYDip = FlowDocumentBuilder.MmToDip(page.MarginHeaderMm);
        double footerYDip = FlowDocumentBuilder.MmToDip(page.MarginFooterMm);
        double bodyWidthDip = Math.Max(1.0, geo.PageWidthDip - geo.PadLeftDip - geo.PadRightDip);

        var foreground = new SolidColorBrush(WpfColor.FromArgb(0xCC, 0x33, 0x33, 0x33));
        foreground.Freeze();

        for (int i = 0; i < pageCount; i++)
        {
            double topY = i * geo.PageStrideDip;
            int pageNumber = page.PageNumberStart + i;
            var ctx = new HeaderFooterTokens.Context
            {
                PageNumber = pageNumber,
                TotalPages = pageCount,
                Now        = resolvedNow,
                Title      = metadata?.Title,
                Author     = metadata?.Author,
                FileName   = fileNameWithoutExt,
            };

            // 머리말: 페이지 상단 ~ 본문 시작 사이. baseline = topY + MarginHeaderMm
            if (page.Header is { } header)
            {
                AddSlot(target, header.Left,   ctx, topY + headerYDip, geo.PadLeftDip, bodyWidthDip,
                        TextAlignment.Left,   foreground);
                AddSlot(target, header.Center, ctx, topY + headerYDip, geo.PadLeftDip, bodyWidthDip,
                        TextAlignment.Center, foreground);
                AddSlot(target, header.Right,  ctx, topY + headerYDip, geo.PadLeftDip, bodyWidthDip,
                        TextAlignment.Right,  foreground);
            }

            // 꼬리말: 본문 끝 ~ 페이지 하단 사이. baseline = topY + pageHeight - MarginFooterMm - 행높이
            if (page.Footer is { } footer)
            {
                // 한 줄 텍스트 가정 — 폰트 행높이 ≈ 14 DIP. 정확한 측정은 첫 자식 Measure 후 보정.
                double approxLineHeight = 14.0;
                double footerY = topY + geo.PageHeightDip - footerYDip - approxLineHeight;
                AddSlot(target, footer.Left,   ctx, footerY, geo.PadLeftDip, bodyWidthDip,
                        TextAlignment.Left,   foreground);
                AddSlot(target, footer.Center, ctx, footerY, geo.PadLeftDip, bodyWidthDip,
                        TextAlignment.Center, foreground);
                AddSlot(target, footer.Right,  ctx, footerY, geo.PadLeftDip, bodyWidthDip,
                        TextAlignment.Right,  foreground);
            }
        }
    }

    private static void AddSlot(
        Canvas                     target,
        HeaderFooterSlot?          slot,
        HeaderFooterTokens.Context ctx,
        double                     topDip,
        double                     leftDip,
        double                     widthDip,
        TextAlignment              alignment,
        Brush                      defaultForeground)
    {
        if (slot is null || slot.IsEmpty) return;

        double y = topDip;
        foreach (var para in slot.Paragraphs)
        {
            if (para.Runs.Count == 0) continue;

            var tb = new TextBlock
            {
                Width            = widthDip,
                TextAlignment    = alignment,
                TextTrimming     = TextTrimming.CharacterEllipsis,
                IsHitTestVisible = false,
                TextWrapping     = TextWrapping.NoWrap,
            };

            foreach (var run in para.Runs)
            {
                var resolvedText = HeaderFooterTokens.Resolve(run.Text, ctx);
                if (string.IsNullOrEmpty(resolvedText)) continue;

                var wpfRun = new WpfDoc.Run(resolvedText);

                // Apply RunStyle
                var s = run.Style;
                if (s.FontSizePt > 0)
                    wpfRun.FontSize = s.FontSizePt * 96.0 / 72.0;
                if (!string.IsNullOrEmpty(s.FontFamily))
                    wpfRun.FontFamily = new FontFamily(s.FontFamily);
                wpfRun.FontWeight  = s.Bold   ? FontWeights.Bold   : FontWeights.Normal;
                wpfRun.FontStyle   = s.Italic ? FontStyles.Italic  : FontStyles.Normal;

                if (s.Foreground is { } fg)
                    wpfRun.Foreground = new SolidColorBrush(WpfColor.FromArgb(fg.A, fg.R, fg.G, fg.B));
                else
                    wpfRun.Foreground = defaultForeground;

                var decorations = new TextDecorationCollection();
                if (s.Underline)     decorations.Add(TextDecorations.Underline[0]);
                if (s.Strikethrough) decorations.Add(TextDecorations.Strikethrough[0]);
                if (decorations.Count > 0)
                    wpfRun.TextDecorations = decorations;

                tb.Inlines.Add(wpfRun);
            }

            if (!tb.Inlines.Any()) continue;

            Canvas.SetLeft(tb, leftDip);
            Canvas.SetTop(tb, y);
            target.Children.Add(tb);

            // Advance y by approximate line height for multi-paragraph slots
            y += (para.Runs.Max(r => r.Style.FontSizePt) * 96.0 / 72.0) * 1.2;
        }
    }
}

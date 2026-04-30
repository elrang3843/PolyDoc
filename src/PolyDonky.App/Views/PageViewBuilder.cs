using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using PolyDonky.App.Pagination;
using PolyDonky.App.Services;
using PolyDonky.Core;
using WpfColor = System.Windows.Media.Color;

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
            if (WpfColor.TryParse(colorHex, out var c))
                return new SolidColorBrush(c) { Opacity = opacity };
        }
        catch { }
        return new SolidColorBrush(Colors.Gray) { Opacity = opacity };
    }
}

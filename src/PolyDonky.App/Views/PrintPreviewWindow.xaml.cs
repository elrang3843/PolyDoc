using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Threading;
using System.Windows.Xps.Packaging;
using PolyDonky.Core;
using Fdb = PolyDonky.App.Services.FlowDocumentBuilder;

namespace PolyDonky.App.Views;

public partial class PrintPreviewWindow : Window
{
    private XpsDocument? _xpsDoc;
    private string? _tmpXpsPath;
    private double _pageWidthDip;
    private double _pageHeightDip;
    private bool _initialZoomApplied;

    public PrintPreviewWindow(PolyDonkyument doc)
    {
        InitializeComponent();
        Loaded += (_, _) => BuildPreview(doc);
        SizeChanged += OnViewerSizeChanged;
    }

    private void BuildPreview(PolyDonkyument doc)
    {
        try
        {
            var page = doc.Sections.FirstOrDefault()?.Page ?? new PageSettings();

            _pageWidthDip  = Fdb.MmToDip(page.EffectiveWidthMm);
            _pageHeightDip = Fdb.MmToDip(page.EffectiveHeightMm);
            double padL = Fdb.MmToDip(page.MarginLeftMm);
            double padT = Fdb.MmToDip(page.MarginTopMm);
            double padR = Fdb.MmToDip(page.MarginRightMm);
            double padB = Fdb.MmToDip(page.MarginBottomMm);

            var fd = Fdb.Build(doc);
            fd.PageWidth   = _pageWidthDip;
            fd.PageHeight  = _pageHeightDip;
            fd.PagePadding = new Thickness(padL, padT, padR, padB);
            fd.ColumnWidth = double.MaxValue;

            _tmpXpsPath = Path.Combine(Path.GetTempPath(), $"pdpreview_{Guid.NewGuid():N}.xps");

            var innerPaginator = ((System.Windows.Documents.IDocumentPaginatorSource)fd).DocumentPaginator;
            innerPaginator.PageSize = new Size(_pageWidthDip, _pageHeightDip);

            // 글상자·오버레이 도형/그림/표는 FlowDocument 가 표현 못 하므로 첫 페이지에 합성한다.
            var overlays = BuildOverlays(doc);
            var paginator = overlays.Count > 0
                ? new OverlayCompositingPaginator(innerPaginator, overlays, new Size(_pageWidthDip, _pageHeightDip))
                : innerPaginator;

            using (var xpsWrite = new XpsDocument(_tmpXpsPath, FileAccess.ReadWrite))
                XpsDocument.CreateXpsDocumentWriter(xpsWrite).Write(paginator);

            _xpsDoc = new XpsDocument(_tmpXpsPath, FileAccess.Read);

            var fixedSeq = _xpsDoc.GetFixedDocumentSequence();
            PreviewViewer.Document = fixedSeq;

            int pageCount = fixedSeq.References.Count > 0
                ? fixedSeq.References[0].GetDocument(false)?.Pages.Count ?? 0
                : 0;
            PageInfoText.Text = pageCount > 0 ? $"총 {pageCount}페이지" : string.Empty;

            // 레이아웃 완료 후 페이지 전체 맞춤 적용 (Loaded 우선순위로 디스패치)
            Dispatcher.BeginInvoke(new Action(ApplyFitPage), DispatcherPriority.Loaded);
        }
        catch (Exception ex)
        {
            PageInfoText.Text = $"미리보기 생성 실패: {ex.Message}";
        }
        finally
        {
            LoadingOverlay.Visibility = Visibility.Collapsed;
        }
    }

    /// <summary>DocumentViewer 가용 영역에 페이지 한 장이 통째로 들어가도록 줌 계산.</summary>
    private void ApplyFitPage()
    {
        if (_pageWidthDip <= 0 || _pageHeightDip <= 0) return;
        if (PreviewViewer.ActualWidth < 50 || PreviewViewer.ActualHeight < 50)
        {
            // 레이아웃이 아직 준비되지 않았으면 한 번 더 디스패치 (단발).
            if (!_initialZoomApplied)
            {
                _initialZoomApplied = true;
                Dispatcher.BeginInvoke(new Action(() => { _initialZoomApplied = false; ApplyFitPage(); }),
                                       DispatcherPriority.Loaded);
            }
            return;
        }

        const double padding = 32;  // 스크롤바·여백 고려
        double availW = PreviewViewer.ActualWidth - padding;
        double availH = PreviewViewer.ActualHeight - padding;

        double zoom = Math.Min(availW / _pageWidthDip, availH / _pageHeightDip) * 100;
        zoom = Math.Max(10, Math.Min(500, zoom));
        SetZoom(zoom);
        _initialZoomApplied = true;
    }

    private void ApplyFitWidth()
    {
        if (_pageWidthDip <= 0) return;
        if (PreviewViewer.ActualWidth < 50) return;
        const double padding = 32;
        double availW = PreviewViewer.ActualWidth - padding;
        double zoom = (availW / _pageWidthDip) * 100;
        zoom = Math.Max(10, Math.Min(500, zoom));
        SetZoom(zoom);
    }

    private void SetZoom(double zoomPercent)
    {
        PreviewViewer.Zoom = zoomPercent;
        ZoomText.Text = $"{Math.Round(zoomPercent)}%";
    }

    private void OnViewerSizeChanged(object sender, SizeChangedEventArgs e)
    {
        // 사용자가 창 크기를 변경했을 때, 초기 fit-page 가 한 번 적용된 이후에는 자동 재계산하지 않는다
        // (사용자가 직접 줌을 조절했을 수 있으므로). 단 초기 적용이 아직 안 됐으면 다시 시도.
        if (!_initialZoomApplied) ApplyFitPage();
    }

    // ── 버튼 핸들러 ──────────────────────────────────────────────

    private void OnZoomInClick(object sender, RoutedEventArgs e)
        => SetZoom(Math.Min(500, PreviewViewer.Zoom * 1.25));

    private void OnZoomOutClick(object sender, RoutedEventArgs e)
        => SetZoom(Math.Max(10, PreviewViewer.Zoom / 1.25));

    private void OnFitPageClick(object sender, RoutedEventArgs e)
        => ApplyFitPage();

    private void OnFitWidthClick(object sender, RoutedEventArgs e)
        => ApplyFitWidth();

    private void OnPrintClick(object sender, RoutedEventArgs e)
        => PreviewViewer.Print();

    private void OnCloseClick(object sender, RoutedEventArgs e)
        => Close();

    // ── 오버레이 합성 ────────────────────────────────────────────────────

    private readonly record struct OverlayItem(UIElement Element, double X, double Y, bool Behind);

    /// <summary>
    /// 모델에서 글상자·오버레이 그림/도형/표를 추출해 페이지 좌표(DIP)와 함께 반환한다.
    /// (XMm/YMm, OverlayXMm/OverlayYMm 모두 페이지 좌상단 원점 기준이므로 단순 단위 변환만 한다.)
    /// </summary>
    private static List<OverlayItem> BuildOverlays(PolyDonkyument doc)
    {
        var list = new List<OverlayItem>();
        var section = doc.Sections.FirstOrDefault();
        if (section is null) return list;

        // 글상자
        foreach (var tb in section.FloatingObjects.OfType<TextBoxObject>())
        {
            var ctrl = new TextBoxOverlay(tb)
            {
                Width  = Fdb.MmToDip(tb.WidthMm),
                Height = Fdb.MmToDip(tb.HeightMm),
            };
            list.Add(new OverlayItem(ctrl, Fdb.MmToDip(tb.XMm), Fdb.MmToDip(tb.YMm), Behind: false));
        }

        // 본문 블록 중 오버레이 모드 항목
        foreach (var block in section.Blocks)
        {
            switch (block)
            {
                case ImageBlock img when img.WrapMode is ImageWrapMode.InFrontOfText
                                                       or ImageWrapMode.BehindText:
                {
                    var ctrl = Fdb.BuildOverlayImageControl(img);
                    if (ctrl is null) break;
                    list.Add(new OverlayItem(ctrl,
                        Fdb.MmToDip(img.OverlayXMm), Fdb.MmToDip(img.OverlayYMm),
                        Behind: img.WrapMode == ImageWrapMode.BehindText));
                    break;
                }
                case ShapeObject shape when shape.WrapMode is ImageWrapMode.InFrontOfText
                                                            or ImageWrapMode.BehindText:
                {
                    var ctrl = Fdb.BuildOverlayShapeControl(shape);
                    list.Add(new OverlayItem(ctrl,
                        Fdb.MmToDip(shape.OverlayXMm), Fdb.MmToDip(shape.OverlayYMm),
                        Behind: shape.WrapMode == ImageWrapMode.BehindText));
                    break;
                }
                case Table tbl when tbl.WrapMode != TableWrapMode.Block:
                {
                    var ctrl = Fdb.BuildOverlayTableControl(tbl);
                    if (ctrl is null) break;
                    list.Add(new OverlayItem(ctrl,
                        Fdb.MmToDip(tbl.OverlayXMm), Fdb.MmToDip(tbl.OverlayYMm),
                        Behind: tbl.WrapMode == TableWrapMode.BehindText));
                    break;
                }
            }
        }
        return list;
    }

    /// <summary>
    /// FlowDocument paginator 를 래핑해 첫 페이지에 오버레이 비주얼을 합성한다.
    /// BehindText 항목은 본문 뒤, InFrontOfText 항목은 본문 위에 그린다.
    /// (현재 모델은 오버레이의 페이지 인덱스를 보관하지 않아 전부 첫 페이지에 합성)
    /// </summary>
    private sealed class OverlayCompositingPaginator : DocumentPaginator
    {
        private readonly DocumentPaginator _inner;
        private readonly IReadOnlyList<OverlayItem> _overlays;
        private readonly Size _pageSize;

        public OverlayCompositingPaginator(DocumentPaginator inner,
            IReadOnlyList<OverlayItem> overlays, Size pageSize)
        {
            _inner    = inner;
            _overlays = overlays;
            _pageSize = pageSize;
        }

        public override DocumentPage GetPage(int pageNumber)
        {
            var inner = _inner.GetPage(pageNumber);
            if (pageNumber != 0 || _overlays.Count == 0) return inner;

            var container = new ContainerVisual();

            // BehindText 먼저
            foreach (var ov in _overlays)
                if (ov.Behind) AddElementAt(container, ov);

            // 본문 페이지 visual — 다른 visual 부모를 가지지 않도록 ContainerVisual 로 한 번 더 감싼다.
            var bodyHost = new ContainerVisual();
            bodyHost.Children.Add(inner.Visual);
            container.Children.Add(bodyHost);

            // InFrontOfText 마지막
            foreach (var ov in _overlays)
                if (!ov.Behind) AddElementAt(container, ov);

            return new DocumentPage(container, _pageSize, inner.BleedBox, inner.ContentBox);
        }

        private static void AddElementAt(ContainerVisual host, OverlayItem ov)
        {
            var e = ov.Element;
            double w = ov.Element is FrameworkElement fe && !double.IsNaN(fe.Width)  && fe.Width  > 0 ? fe.Width  : double.PositiveInfinity;
            double h = ov.Element is FrameworkElement fe2 && !double.IsNaN(fe2.Height) && fe2.Height > 0 ? fe2.Height : double.PositiveInfinity;
            e.Measure(new Size(w, h));
            var sz = new Size(
                double.IsInfinity(w) ? e.DesiredSize.Width  : w,
                double.IsInfinity(h) ? e.DesiredSize.Height : h);
            e.Arrange(new Rect(ov.X, ov.Y, sz.Width, sz.Height));
            host.Children.Add(e);
        }

        public override bool                       IsPageCountValid    => _inner.IsPageCountValid;
        public override int                        PageCount           => _inner.PageCount;
        public override Size                       PageSize            { get => _inner.PageSize; set => _inner.PageSize = value; }
        public override IDocumentPaginatorSource   Source              => _inner.Source;
        public override void                       ComputePageCount()  => _inner.ComputePageCount();
    }

    // ── Window lifecycle ─────────────────────────────────────────────────

    private void OnWindowClosed(object? sender, EventArgs e)
    {
        PreviewViewer.Document = null;
        _xpsDoc?.Close();
        _xpsDoc = null;
        if (_tmpXpsPath is not null)
        {
            try { File.Delete(_tmpXpsPath); } catch { /* best-effort */ }
            _tmpXpsPath = null;
        }
    }
}

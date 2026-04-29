using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using System.Windows.Xps.Packaging;
using PolyDonky.Core;
using Fdb = PolyDonky.App.Services.FlowDocumentBuilder;

namespace PolyDonky.App.Views;

public partial class PrintPreviewWindow : Window
{
    // ── 상태 ─────────────────────────────────────────────────────────────

    private readonly PolyDonkyument _doc;
    private PageSettings _printPage;       // 패널에서 수정된 용지 설정 (원본 불변)
    private bool _monochrome;
    private int  _copies = 1;

    private XpsDocument? _xpsDoc;
    private string?      _tmpXpsPath;
    private double _pageWidthDip;
    private double _pageHeightDip;
    private bool   _initialZoomApplied;
    private bool   _suppressSettingsEvents; // 초기화 중 이벤트 억제

    private readonly DispatcherTimer _rebuildTimer;

    // ── 용지·여백 프리셋 테이블 ─────────────────────────────────────────

    private static readonly Dictionary<string, PaperSizeKind> PaperSizeMap = new()
    {
        ["A4"]     = PaperSizeKind.A4,
        ["A3"]     = PaperSizeKind.A3,
        ["A5"]     = PaperSizeKind.A5,
        ["B5_JIS"] = PaperSizeKind.B5_JIS,
        ["B4_JIS"] = PaperSizeKind.B4_JIS,
        ["Letter"] = PaperSizeKind.Letter,
        ["Legal"]  = PaperSizeKind.Legal,
    };

    // (상, 하, 좌, 우) mm
    private static readonly Dictionary<string, (double T, double B, double L, double R)> MarginPresets = new()
    {
        ["Normal"] = (25,   25,   25,   25),
        ["Narrow"] = (12.7, 12.7, 12.7, 12.7),
        ["Wide"]   = (38.1, 38.1, 38.1, 38.1),
        ["None"]   = (0,    0,    0,    0),
    };

    // ── 생성자 ────────────────────────────────────────────────────────────

    public PrintPreviewWindow(PolyDonkyument doc)
    {
        InitializeComponent();

        _doc       = doc;
        _printPage = ClonePage(doc.Sections.FirstOrDefault()?.Page ?? new PageSettings());

        _rebuildTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(350),
        };
        _rebuildTimer.Tick += (_, _) => { _rebuildTimer.Stop(); RebuildPreview(); };

        Loaded      += OnWindowLoaded;
        SizeChanged += OnViewerSizeChanged;
    }

    private void OnWindowLoaded(object sender, RoutedEventArgs e)
    {
        InitSettingsPanel();
        BuildPreview();
    }

    // ── 설정 패널 초기화 ─────────────────────────────────────────────────

    private void InitSettingsPanel()
    {
        _suppressSettingsEvents = true;

        // 용지 크기: 문서 기본값 선택
        PaperSizeCombo.SelectedIndex = 0;

        // 방향
        PortraitRadio.IsChecked  = _printPage.Orientation == PageOrientation.Portrait;
        LandscapeRadio.IsChecked = _printPage.Orientation == PageOrientation.Landscape;

        // 여백: 문서 기본값 선택
        MarginPresetCombo.SelectedIndex = 0;

        // 색상: 컬러
        ColorRadio.IsChecked = true;

        // 매수
        CopiesBox.Text = "1";

        _suppressSettingsEvents = false;
    }

    // ── 설정 이벤트 핸들러 ───────────────────────────────────────────────

    private void OnPaperSizeChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_suppressSettingsEvents) return;
        var tag = (PaperSizeCombo.SelectedItem as ComboBoxItem)?.Tag as string ?? "Default";
        if (tag == "Default")
        {
            _printPage = ClonePage(_doc.Sections.FirstOrDefault()?.Page ?? new PageSettings());
            // 방향은 그대로 유지
            _printPage.Orientation = PortraitRadio.IsChecked == true
                ? PageOrientation.Portrait : PageOrientation.Landscape;
        }
        else if (PaperSizeMap.TryGetValue(tag, out var kind))
        {
            _printPage.ApplySizeKind(kind);
        }
        QueueRebuild();
    }

    private void OnOrientationChanged(object sender, RoutedEventArgs e)
    {
        if (_suppressSettingsEvents) return;
        _printPage.Orientation = PortraitRadio.IsChecked == true
            ? PageOrientation.Portrait : PageOrientation.Landscape;
        QueueRebuild();
    }

    private void OnMarginPresetChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_suppressSettingsEvents) return;
        var tag = (MarginPresetCombo.SelectedItem as ComboBoxItem)?.Tag as string ?? "Default";
        if (tag == "Default")
        {
            var orig = _doc.Sections.FirstOrDefault()?.Page ?? new PageSettings();
            _printPage.MarginTopMm    = orig.MarginTopMm;
            _printPage.MarginBottomMm = orig.MarginBottomMm;
            _printPage.MarginLeftMm   = orig.MarginLeftMm;
            _printPage.MarginRightMm  = orig.MarginRightMm;
        }
        else if (MarginPresets.TryGetValue(tag, out var m))
        {
            _printPage.MarginTopMm    = m.T;
            _printPage.MarginBottomMm = m.B;
            _printPage.MarginLeftMm   = m.L;
            _printPage.MarginRightMm  = m.R;
        }
        QueueRebuild();
    }

    private void OnColorModeChanged(object sender, RoutedEventArgs e)
    {
        if (_suppressSettingsEvents) return;
        _monochrome = MonoRadio.IsChecked == true;
        QueueRebuild();
    }

    private void OnCopiesPreviewInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        => e.Handled = !e.Text.All(char.IsDigit);

    private void OnCopiesUp(object sender, RoutedEventArgs e)
        => SetCopies(_copies + 1);

    private void OnCopiesDown(object sender, RoutedEventArgs e)
        => SetCopies(_copies - 1);

    private void SetCopies(int v)
    {
        _copies = Math.Clamp(v, 1, 99);
        CopiesBox.Text = _copies.ToString();
    }

    private void QueueRebuild()
    {
        LoadingOverlay.Visibility = Visibility.Visible;
        LoadingText.Text = "미리보기 재생성 중...";
        _rebuildTimer.Stop();
        _rebuildTimer.Start();
    }

    // ── 미리보기 빌드 ────────────────────────────────────────────────────

    private void BuildPreview()
    {
        _initialZoomApplied = false;
        try
        {
            _pageWidthDip  = Fdb.MmToDip(_printPage.EffectiveWidthMm);
            _pageHeightDip = Fdb.MmToDip(_printPage.EffectiveHeightMm);
            double padL = Fdb.MmToDip(_printPage.MarginLeftMm);
            double padT = Fdb.MmToDip(_printPage.MarginTopMm);
            double padR = Fdb.MmToDip(_printPage.MarginRightMm);
            double padB = Fdb.MmToDip(_printPage.MarginBottomMm);

            // FlowDocument 는 문서 원본(_doc) 기반으로 빌드하되
            // 용지 설정은 패널에서 선택한 _printPage 를 적용한다.
            var docForBuild = BuildDocWithPage(_doc, _printPage);
            var fd = Fdb.Build(docForBuild);
            fd.PageWidth   = _pageWidthDip;
            fd.PageHeight  = _pageHeightDip;
            fd.PagePadding = new Thickness(padL, padT, padR, padB);
            fd.ColumnWidth = double.MaxValue;

            // 이전 XPS 정리
            CleanupXps();

            _tmpXpsPath = Path.Combine(Path.GetTempPath(), $"pdpreview_{Guid.NewGuid():N}.xps");

            var innerPaginator = ((IDocumentPaginatorSource)fd).DocumentPaginator;
            innerPaginator.PageSize = new Size(_pageWidthDip, _pageHeightDip);

            var overlays = BuildOverlays(docForBuild);
            DocumentPaginator paginator = overlays.Count > 0
                ? new OverlayCompositingPaginator(innerPaginator, overlays, new Size(_pageWidthDip, _pageHeightDip), _monochrome, padT)
                : _monochrome
                    ? new GrayscalePaginator(innerPaginator, new Size(_pageWidthDip, _pageHeightDip))
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

    private void RebuildPreview()
    {
        PreviewViewer.Document = null;
        CleanupXps();
        BuildPreview();
    }

    /// <summary>_doc 의 블록/플로팅 객체는 유지하고 용지 설정만 교체한 임시 문서를 반환한다.</summary>
    private static PolyDonkyument BuildDocWithPage(PolyDonkyument src, PageSettings page)
    {
        var clone = new PolyDonkyument();
        foreach (var sec in src.Sections)
        {
            var newSec = new Section
            {
                Id     = sec.Id,
                Page   = page,
                Blocks = sec.Blocks,
            };
            clone.Sections.Add(newSec);
        }
        clone.OutlineStyles = src.OutlineStyles;
        return clone;
    }

    /// <summary>원본 PageSettings 의 얕은 복사본 반환. WidthMm/HeightMm/Margin 등 값 타입만 복사.</summary>
    private static PageSettings ClonePage(PageSettings src) => new()
    {
        SizeKind        = src.SizeKind,
        WidthMm         = src.WidthMm,
        HeightMm        = src.HeightMm,
        Orientation     = src.Orientation,
        MarginTopMm     = src.MarginTopMm,
        MarginBottomMm  = src.MarginBottomMm,
        MarginLeftMm    = src.MarginLeftMm,
        MarginRightMm   = src.MarginRightMm,
        MarginHeaderMm  = src.MarginHeaderMm,
        MarginFooterMm  = src.MarginFooterMm,
        PaperColor      = src.PaperColor,
        ColumnCount     = src.ColumnCount,
        ColumnGapMm     = src.ColumnGapMm,
        PageNumberStart = src.PageNumberStart,
    };

    // ── 줌 ───────────────────────────────────────────────────────────────

    private void ApplyFitPage()
    {
        if (_pageWidthDip <= 0 || _pageHeightDip <= 0) return;
        if (PreviewViewer.ActualWidth < 50 || PreviewViewer.ActualHeight < 50)
        {
            if (!_initialZoomApplied)
            {
                _initialZoomApplied = true;
                Dispatcher.BeginInvoke(new Action(() => { _initialZoomApplied = false; ApplyFitPage(); }),
                                       DispatcherPriority.Loaded);
            }
            return;
        }
        const double padding = 32;
        double availW = PreviewViewer.ActualWidth  - padding;
        double availH = PreviewViewer.ActualHeight - padding;
        double zoom   = Math.Min(availW / _pageWidthDip, availH / _pageHeightDip) * 100;
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
        double zoom   = (availW / _pageWidthDip) * 100;
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
        if (!_initialZoomApplied) ApplyFitPage();
    }

    // ── 버튼 핸들러 ──────────────────────────────────────────────────────

    private void OnZoomInClick(object sender, RoutedEventArgs e)
        => SetZoom(Math.Min(500, PreviewViewer.Zoom * 1.25));

    private void OnZoomOutClick(object sender, RoutedEventArgs e)
        => SetZoom(Math.Max(10, PreviewViewer.Zoom / 1.25));

    private void OnFitPageClick(object sender, RoutedEventArgs e)
        => ApplyFitPage();

    private void OnFitWidthClick(object sender, RoutedEventArgs e)
        => ApplyFitWidth();

    private void OnPrintClick(object sender, RoutedEventArgs e)
    {
        if (int.TryParse(CopiesBox.Text, out int c) && c >= 1) _copies = c;

        var dlg = new System.Windows.Controls.PrintDialog();
        try
        {
            if (_monochrome)
                dlg.PrintTicket.OutputColor = System.Printing.OutputColor.Monochrome;
            dlg.PrintTicket.CopyCount = (ushort)Math.Clamp(_copies, 1, 99);
        }
        catch { /* PrintTicket 미지원 드라이버 무시 */ }

        if (dlg.ShowDialog() != true) return;

        if (PreviewViewer.Document is IDocumentPaginatorSource src)
            dlg.PrintDocument(src.DocumentPaginator, "PolyDonky 문서");
    }

    private void OnCloseClick(object sender, RoutedEventArgs e)
        => Close();

    // ── 오버레이 합성 ────────────────────────────────────────────────────

    private readonly record struct OverlayItem(UIElement Element, double X, double Y, bool Behind);

    private static List<OverlayItem> BuildOverlays(PolyDonkyument doc)
    {
        var list    = new List<OverlayItem>();
        var section = doc.Sections.FirstOrDefault();
        if (section is null) return list;

        // 통합 모델 — 도형·이미지·표·글상자 모두 section.Blocks 안에서 함께 순회.
        foreach (var block in section.Blocks)
        {
            switch (block)
            {
                case TextBoxObject tb:
                {
                    var ctrl = new TextBoxOverlay(tb)
                    {
                        Width  = Fdb.MmToDip(tb.WidthMm),
                        Height = Fdb.MmToDip(tb.HeightMm),
                    };
                    list.Add(new OverlayItem(ctrl,
                        Fdb.MmToDip(tb.OverlayXMm), Fdb.MmToDip(tb.OverlayYMm),
                        Behind: tb.WrapMode == ImageWrapMode.BehindText));
                    break;
                }
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
                case PolyDonky.Core.Table tbl when tbl.WrapMode != TableWrapMode.Block:
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

    // ── Paginator ────────────────────────────────────────────────────────

    /// <summary>
    /// 오버레이를 해당 페이지 Y 범위 기준으로 합성하고, monochrome 이면 추가로 회색 변환.
    /// </summary>
    private sealed class OverlayCompositingPaginator : DocumentPaginator
    {
        private readonly DocumentPaginator       _inner;
        private readonly IReadOnlyList<OverlayItem> _overlays;
        private readonly Size   _pageSize;
        private readonly bool   _monochrome;
        private readonly double _padTopDip;   // 미리보기 페이지별 상단 여백 (DIP)

        public OverlayCompositingPaginator(DocumentPaginator inner,
            IReadOnlyList<OverlayItem> overlays, Size pageSize, bool monochrome,
            double padTopDip)
        {
            _inner      = inner;
            _overlays   = overlays;
            _pageSize   = pageSize;
            _monochrome = monochrome;
            _padTopDip  = padTopDip;
        }

        public override DocumentPage GetPage(int pageNumber)
        {
            var inner = _inner.GetPage(pageNumber);

            double pageTop    = pageNumber       * _pageSize.Height;
            double pageBottom = (pageNumber + 1) * _pageSize.Height;

            // MainWindow 연속 보기에는 페이지별 상단 여백이 없지만,
            // 미리보기 paginator 는 각 페이지에 padT 를 추가한다.
            // 2페이지 이후 오버레이가 padT 만큼 위로 올라가는 것을 보정한다.
            double yCorrection = pageNumber > 0 ? _padTopDip : 0;

            var pageOverlays = _overlays
                .Where(ov => ov.Y >= pageTop && ov.Y < pageBottom)
                .Select(ov => ov with { Y = ov.Y - pageTop + yCorrection })
                .ToList();

            DocumentPage result;
            if (pageOverlays.Count == 0)
            {
                result = inner;
            }
            else
            {
                var container = new ContainerVisual();

                foreach (var ov in pageOverlays)
                    if (ov.Behind) AddElementAt(container, ov);

                var bodyHost = new ContainerVisual();
                bodyHost.Children.Add(inner.Visual);
                container.Children.Add(bodyHost);

                foreach (var ov in pageOverlays)
                    if (!ov.Behind) AddElementAt(container, ov);

                result = new DocumentPage(container, _pageSize, inner.BleedBox, inner.ContentBox);
            }

            return _monochrome ? ToGrayscale(result, _pageSize) : result;
        }

        private static void AddElementAt(ContainerVisual host, OverlayItem ov)
        {
            var e = ov.Element;
            double w = e is FrameworkElement fe  && !double.IsNaN(fe.Width)   && fe.Width  > 0 ? fe.Width  : double.PositiveInfinity;
            double h = e is FrameworkElement fe2 && !double.IsNaN(fe2.Height) && fe2.Height > 0 ? fe2.Height : double.PositiveInfinity;
            e.Measure(new Size(w, h));
            var sz = new Size(double.IsInfinity(w) ? e.DesiredSize.Width : w,
                              double.IsInfinity(h) ? e.DesiredSize.Height : h);
            e.Arrange(new Rect(ov.X, ov.Y, sz.Width, sz.Height));
            host.Children.Add(e);
        }

        public override bool                     IsPageCountValid => _inner.IsPageCountValid;
        public override int                      PageCount        => _inner.PageCount;
        public override Size                     PageSize         { get => _inner.PageSize; set => _inner.PageSize = value; }
        public override IDocumentPaginatorSource Source           => _inner.Source;
        public override void                     ComputePageCount() => _inner.ComputePageCount();
    }

    /// <summary>오버레이 없이 흑백만 필요한 경우 사용하는 단순 래퍼.</summary>
    private sealed class GrayscalePaginator : DocumentPaginator
    {
        private readonly DocumentPaginator _inner;
        private readonly Size _pageSize;

        public GrayscalePaginator(DocumentPaginator inner, Size pageSize)
        {
            _inner    = inner;
            _pageSize = pageSize;
        }

        public override DocumentPage GetPage(int pageNumber)
            => ToGrayscale(_inner.GetPage(pageNumber), _pageSize);

        public override bool                     IsPageCountValid => _inner.IsPageCountValid;
        public override int                      PageCount        => _inner.PageCount;
        public override Size                     PageSize         { get => _inner.PageSize; set => _inner.PageSize = value; }
        public override IDocumentPaginatorSource Source           => _inner.Source;
        public override void                     ComputePageCount() => _inner.ComputePageCount();
    }

    /// <summary>
    /// DocumentPage.Visual 을 150 DPI 로 래스터화 후 Gray8 로 변환해 흑백 페이지를 반환.
    /// (화면 미리보기용. 실제 인쇄는 PrintTicket.OutputColor = Monochrome 사용.)
    /// </summary>
    private static DocumentPage ToGrayscale(DocumentPage page, Size size)
    {
        const double dpi = 150;
        int pw = Math.Max(1, (int)(size.Width  * dpi / 96));
        int ph = Math.Max(1, (int)(size.Height * dpi / 96));

        var rtb = new RenderTargetBitmap(pw, ph, dpi, dpi, PixelFormats.Pbgra32);
        rtb.Render(page.Visual);
        rtb.Freeze();

        var gray = new FormatConvertedBitmap(rtb, PixelFormats.Gray8, null, 0);
        gray.Freeze();

        var dv = new DrawingVisual();
        using (var dc = dv.RenderOpen())
            dc.DrawImage(gray, new Rect(0, 0, size.Width, size.Height));

        return new DocumentPage(dv, size, page.BleedBox, page.ContentBox);
    }

    // ── Window lifecycle ─────────────────────────────────────────────────

    private void CleanupXps()
    {
        _xpsDoc?.Close();
        _xpsDoc = null;
        if (_tmpXpsPath is not null)
        {
            try { File.Delete(_tmpXpsPath); } catch { }
            _tmpXpsPath = null;
        }
    }

    private void OnWindowClosed(object? sender, EventArgs e)
    {
        _rebuildTimer.Stop();
        PreviewViewer.Document = null;
        CleanupXps();
    }
}

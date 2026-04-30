using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using PolyDonky.App.Pagination;
using PolyDonky.App.Services;
using PolyDonky.Core;
using EnumeratedPrintQueueTypes = System.Printing.EnumeratedPrintQueueTypes;
using LocalPrintServer          = System.Printing.LocalPrintServer;
using OutputColor               = System.Printing.OutputColor;
using PrintQueue                = System.Printing.PrintQueue;

namespace PolyDonky.App.Views;

public partial class PrintPreviewWindow : Window
{
    // ── 상태 ─────────────────────────────────────────────────────────────

    private readonly PolyDonkyument _doc;
    private PageSettings _printPage;
    private bool _monochrome;
    private int  _copies = 1;

    private PageGeometry? _geo;
    private int           _pageCount;
    private bool          _initialZoomApplied;

    // InitializeComponent() 가 ComboBox 초기 SelectedIndex 적용 시 이벤트 억제.
    private bool _suppressSettingsEvents = true;

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

        _rebuildTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(350) };
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

        PaperSizeCombo.SelectedIndex   = 0;
        PortraitRadio.IsChecked        = _printPage.Orientation == PageOrientation.Portrait;
        LandscapeRadio.IsChecked       = _printPage.Orientation == PageOrientation.Landscape;
        MarginPresetCombo.SelectedIndex = 0;
        ColorRadio.IsChecked           = true;
        CopiesBox.Text                 = "1";

        PopulatePrinters();

        _suppressSettingsEvents = false;
    }

    /// <summary>
    /// 시스템에 설치된 프린터 목록을 PrinterCombo 에 채우고 기본 프린터를 선택한다.
    /// 인쇄 시 시스템 PrintDialog 를 띄우지 않고 여기서 선택한 PrintQueue 로 직접 전송한다.
    /// </summary>
    private void PopulatePrinters()
    {
        PrinterCombo.Items.Clear();
        try
        {
            using var server = new LocalPrintServer();
            var queues = server.GetPrintQueues(new[]
            {
                EnumeratedPrintQueueTypes.Local,
                EnumeratedPrintQueueTypes.Connections,
            });

            string? defaultName = null;
            try { defaultName = LocalPrintServer.GetDefaultPrintQueue()?.FullName; } catch { }

            ComboBoxItem? defaultItem = null;
            foreach (var pq in queues)
            {
                var item = new ComboBoxItem
                {
                    Content = pq.FullName,
                    Tag     = pq,
                };
                PrinterCombo.Items.Add(item);
                if (defaultName != null && pq.FullName == defaultName)
                    defaultItem = item;
            }

            if (defaultItem != null) defaultItem.IsSelected = true;
            else if (PrinterCombo.Items.Count > 0)
                ((ComboBoxItem)PrinterCombo.Items[0]).IsSelected = true;
        }
        catch
        {
            // 프린터 열거 실패 — ComboBox 비어 있으면 OnPrintClick 이 안내 메시지 표시.
        }
    }

    // ── 설정 이벤트 핸들러 ───────────────────────────────────────────────

    private void OnPaperSizeChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_suppressSettingsEvents) return;
        var tag = (PaperSizeCombo.SelectedItem as ComboBoxItem)?.Tag as string ?? "Default";
        if (tag == "Default")
        {
            _printPage = ClonePage(_doc.Sections.FirstOrDefault()?.Page ?? new PageSettings());
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

    private void OnCopiesUp(object sender, RoutedEventArgs e)   => SetCopies(_copies + 1);
    private void OnCopiesDown(object sender, RoutedEventArgs e) => SetCopies(_copies - 1);

    private void SetCopies(int v)
    {
        _copies        = Math.Clamp(v, 1, 99);
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
        GrayscaleOverlay.Visibility = Visibility.Collapsed;
        GrayscaleOverlay.Source     = null;
        try
        {
            // 1. 페이지 설정이 적용된 임시 문서 생성
            var docForBuild = BuildDocWithPage(_doc, _printPage);

            // 2. 편집창과 동일한 파이프라인으로 페이지 분할
            var paginated = FlowDocumentPaginationAdapter.Paginate(docForBuild);
            var slices    = PerPageDocumentSplitter.Split(paginated);
            var geo       = new PageGeometry(_printPage);

            _geo       = geo;
            _pageCount = paginated.PageCount;

            // 3. 읽기 전용 per-page RTB 구성 (편집창 ConfigurePageRtb 와 동일 구조, 편집 잠금)
            PreviewPageHost.SetupPages(slices, geo, rtb =>
            {
                rtb.IsReadOnly  = true;
                rtb.Focusable   = false;
                rtb.Background  = Brushes.Transparent;
                rtb.BorderThickness = new Thickness(0);
                // Disabled: 스크롤 자체를 차단 (Hidden 은 스크롤바만 숨길 뿐 스크롤은 허용)
                rtb.VerticalScrollBarVisibility   = ScrollBarVisibility.Disabled;
                rtb.HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled;
            });

            // 4. 페이지 배경 프레임
            var pageSettings = docForBuild.Sections.FirstOrDefault()?.Page;
            PageViewBuilder.BuildPageFrames(
                PreviewPageBgCanvas, geo, _pageCount, pageSettings,
                pageBg: Brushes.White, showShadow: true, showGuides: false, showLabels: true);

            // 5. 오버레이 캔버스 채우기
            PageViewBuilder.PopulateOverlayCanvases(
                paginated, geo,
                PreviewOverlayShapeCanvas,  PreviewUnderlayShapeCanvas,
                PreviewOverlayImageCanvas,  PreviewUnderlayImageCanvas,
                PreviewOverlayTableCanvas,  PreviewUnderlayTableCanvas,
                PreviewFloatingCanvas);

            // 5.5. 워터마크 렌더링
            PageViewBuilder.BuildWatermarkLayer(
                PreviewWatermarkCanvas, docForBuild.Watermark, geo, _pageCount);

            // 6. 오버레이 캔버스 클립 — 본문 텍스트가 per-page RTB 에서 페이지마다 잘려 보이는 것과
            // 동일하게, 모든 부유 객체(글상자·도형·이미지·표) 도 페이지 경계 밖(특히 페이지 간 갭) 에서
            // 잘려 보이도록 한다. 편집창과 동일한 시각 결과를 보장한다.
            var clip = PageViewBuilder.BuildPageClipGeometry(geo, _pageCount);
            PreviewOverlayShapeCanvas.Clip  = clip;
            PreviewUnderlayShapeCanvas.Clip = clip;
            PreviewOverlayImageCanvas.Clip  = clip;
            PreviewUnderlayImageCanvas.Clip = clip;
            PreviewOverlayTableCanvas.Clip  = clip;
            PreviewUnderlayTableCanvas.Clip = clip;
            PreviewFloatingCanvas.Clip      = clip;
            PreviewWatermarkCanvas.Clip     = clip;

            // 7. PaperHost 크기 설정
            PreviewPaperHost.Width     = geo.PageWidthDip;
            PreviewPaperHost.MinHeight = geo.TotalHeightDip(_pageCount);

            PageInfoText.Text = $"총 {_pageCount}페이지";

            Dispatcher.BeginInvoke(new Action(ApplyFitPage), DispatcherPriority.Loaded);
            if (_monochrome)
                Dispatcher.BeginInvoke(new Action(ApplyMonochromeOverlay), DispatcherPriority.Render);
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
        // per-page RTB 초기화
        PreviewPageHost.SetupPages([], new PageGeometry(new PageSettings()), null);
        BuildPreview();
    }

    // ── 흑백 미리보기 오버레이 ───────────────────────────────────────────

    /// <summary>
    /// PreviewPaperHost 전체를 RenderTargetBitmap 으로 캡처해 Gray8 로 변환한 뒤
    /// GrayscaleOverlay Image 로 덮어 흑백 미리보기를 구현한다.
    /// DispatcherPriority.Render 로 호출돼 레이아웃이 완료된 시점에 실행된다.
    /// </summary>
    private void ApplyMonochromeOverlay()
    {
        if (!_monochrome) return;
        try
        {
            double w = PreviewPaperHost.ActualWidth;
            double h = PreviewPaperHost.ActualHeight;
            if (w < 1 || h < 1) return;

            int pw = Math.Max(1, (int)Math.Ceiling(w));
            int ph = Math.Max(1, (int)Math.Ceiling(h));

            var rtb = new RenderTargetBitmap(pw, ph, 96, 96, PixelFormats.Pbgra32);
            rtb.Render(PreviewPaperHost);
            rtb.Freeze();

            var gray = new FormatConvertedBitmap(rtb, PixelFormats.Gray8, null, 0);
            gray.Freeze();

            GrayscaleOverlay.Width      = w;
            GrayscaleOverlay.Height     = h;
            GrayscaleOverlay.Source     = gray;
            GrayscaleOverlay.Visibility = Visibility.Visible;
        }
        catch { /* 레이아웃 미완료 등 무시 */ }
    }

    // ── 줌 ───────────────────────────────────────────────────────────────

    private void ApplyFitPage()
    {
        if (_geo is null) return;
        if (PreviewScroll.ActualWidth < 50 || PreviewScroll.ActualHeight < 50)
        {
            if (!_initialZoomApplied)
            {
                _initialZoomApplied = true;
                Dispatcher.BeginInvoke(new Action(() => { _initialZoomApplied = false; ApplyFitPage(); }),
                                       DispatcherPriority.Loaded);
            }
            return;
        }
        const double padding = 64;
        double zoomW = (PreviewScroll.ActualWidth  - padding) / _geo.PageWidthDip;
        double zoomH = (PreviewScroll.ActualHeight - padding) / _geo.PageHeightDip;
        SetZoom(Math.Clamp(Math.Min(zoomW, zoomH) * 100, 10, 500));
        _initialZoomApplied = true;
    }

    private void ApplyFitWidth()
    {
        if (_geo is null) return;
        const double padding = 64;
        SetZoom(Math.Clamp((PreviewScroll.ActualWidth - padding) / _geo.PageWidthDip * 100, 10, 500));
    }

    private void SetZoom(double zoomPercent)
    {
        double scale      = zoomPercent / 100.0;
        PreviewScale.ScaleX = scale;
        PreviewScale.ScaleY = scale;
        ZoomText.Text     = $"{Math.Round(zoomPercent)}%";
    }

    private void OnViewerSizeChanged(object sender, SizeChangedEventArgs e)
    {
        if (!_initialZoomApplied) ApplyFitPage();
    }

    // ── 버튼 핸들러 ──────────────────────────────────────────────────────

    private void OnZoomInClick(object sender, RoutedEventArgs e)
        => SetZoom(Math.Min(500, (PreviewScale.ScaleX * 100) * 1.25));

    private void OnZoomOutClick(object sender, RoutedEventArgs e)
        => SetZoom(Math.Max(10, (PreviewScale.ScaleX * 100) / 1.25));

    private void OnFitPageClick(object sender, RoutedEventArgs e)  => ApplyFitPage();
    private void OnFitWidthClick(object sender, RoutedEventArgs e) => ApplyFitWidth();

    private void OnPrintClick(object sender, RoutedEventArgs e)
    {
        if (_geo is null) return;
        if (int.TryParse(CopiesBox.Text, out int c) && c >= 1) _copies = c;

        // 좌측 패널에서 선택된 프린터 — 시스템 PrintDialog 를 띄우지 않고 곧바로 전송.
        // (좌측 패널이 이미 미리보기 + 용지/색상/매수 옵션을 갖고 있어 시스템 다이얼로그가 중복.)
        var pq = (PrinterCombo.SelectedItem as ComboBoxItem)?.Tag as PrintQueue;
        if (pq is null)
        {
            MessageBox.Show(this,
                "프린터를 선택하세요.",
                "PolyDonky",
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var dlg = new System.Windows.Controls.PrintDialog { PrintQueue = pq };
        try
        {
            if (_monochrome)
                dlg.PrintTicket.OutputColor = OutputColor.Monochrome;
            dlg.PrintTicket.CopyCount = (ushort)Math.Clamp(_copies, 1, 99);
        }
        catch { /* PrintTicket 미지원 드라이버 무시 */ }

        var paginator = new LiveViewPaginator(PreviewPaperHost, _geo, _pageCount, _monochrome);
        dlg.PrintDocument(paginator, "PolyDonky 문서");
    }

    private void OnCloseClick(object sender, RoutedEventArgs e) => Close();

    // ── 인쇄용 Paginator ─────────────────────────────────────────────────

    /// <summary>
    /// PreviewPaperHost 의 각 페이지 영역을 VisualBrush 로 캡처해 DocumentPage 로 반환한다.
    /// 미리보기와 인쇄가 완전히 동일한 렌더링 경로를 사용한다.
    /// </summary>
    private sealed class LiveViewPaginator : DocumentPaginator
    {
        private readonly UIElement   _source;
        private readonly PageGeometry _geo;
        private readonly int         _pageCount;
        private readonly bool        _monochrome;

        public LiveViewPaginator(UIElement source, PageGeometry geo, int pageCount, bool monochrome)
        {
            _source     = source;
            _geo        = geo;
            _pageCount  = pageCount;
            _monochrome = monochrome;
        }

        public override DocumentPage GetPage(int pageNumber)
        {
            var sz = new Size(_geo.PageWidthDip, _geo.PageHeightDip);

            // PreviewPaperHost 에서 해당 페이지 영역만 뽑아 (0,0) 에 렌더링
            var brush = new VisualBrush(_source)
            {
                Viewbox      = new Rect(0, pageNumber * _geo.PageStrideDip, sz.Width, sz.Height),
                ViewboxUnits = BrushMappingMode.Absolute,
                Stretch      = Stretch.None,
            };

            var dv = new DrawingVisual();
            using (var dc = dv.RenderOpen())
                dc.DrawRectangle(brush, null, new Rect(0, 0, sz.Width, sz.Height));

            var page = new DocumentPage(dv, sz, new Rect(sz), new Rect(sz));
            return _monochrome ? ToGrayscale(page, sz) : page;
        }

        public override bool IsPageCountValid => true;
        public override int  PageCount        => _pageCount;
        public override Size PageSize         { get; set; }
        public override IDocumentPaginatorSource Source => null!;
        public override void ComputePageCount() { }
    }

    // ── 흑백 변환 ────────────────────────────────────────────────────────

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

    // ── 유틸 ─────────────────────────────────────────────────────────────

    private static PolyDonkyument BuildDocWithPage(PolyDonkyument src, PageSettings page)
    {
        var clone = new PolyDonkyument();
        foreach (var sec in src.Sections)
        {
            clone.Sections.Add(new PolyDonky.Core.Section
            {
                Id     = sec.Id,
                Page   = page,
                Blocks = sec.Blocks,
            });
        }
        clone.OutlineStyles = src.OutlineStyles;
        return clone;
    }

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

    // ── Window lifecycle ─────────────────────────────────────────────────

    private void OnWindowClosed(object? sender, EventArgs e)
    {
        _rebuildTimer.Stop();
        PreviewPageHost.SetupPages([], new PageGeometry(new PageSettings()), null);
    }
}

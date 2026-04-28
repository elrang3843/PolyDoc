using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Windows;
using System.Windows.Xps.Packaging;
using PolyDonky.Core;
using Fdb = PolyDonky.App.Services.FlowDocumentBuilder;

namespace PolyDonky.App.Views;

public partial class PrintPreviewWindow : Window
{
    private XpsDocument? _xpsDoc;
    private Uri? _packUri;
    private MemoryStream? _xpsStream;

    public PrintPreviewWindow(PolyDonkyument doc)
    {
        InitializeComponent();
        Loaded += (_, _) => BuildPreviewAsync(doc);
    }

    private void BuildPreviewAsync(PolyDonkyument doc)
    {
        try
        {
            var page = doc.Sections.FirstOrDefault()?.Page ?? new PageSettings();

            double pageW = Fdb.MmToDip(page.EffectiveWidthMm);
            double pageH = Fdb.MmToDip(page.EffectiveHeightMm);
            double padL  = Fdb.MmToDip(page.MarginLeftMm);
            double padT  = Fdb.MmToDip(page.MarginTopMm);
            double padR  = Fdb.MmToDip(page.MarginRightMm);
            double padB  = Fdb.MmToDip(page.MarginBottomMm);

            var fd = Fdb.Build(doc);
            fd.PageWidth   = pageW;
            fd.PageHeight  = pageH;
            fd.PagePadding = new Thickness(padL, padT, padR, padB);
            fd.ColumnWidth = double.MaxValue;

            _xpsStream = new MemoryStream();
            var pkg = Package.Open(_xpsStream, FileMode.Create, FileAccess.ReadWrite);
            _packUri = new Uri($"pack://preview{Guid.NewGuid():N}.xps");
            PackageStore.AddPackage(_packUri, pkg);
            _xpsDoc = new XpsDocument(pkg, CompressionOption.NotCompressed, _packUri.AbsoluteUri);

            var paginator = ((System.Windows.Documents.IDocumentPaginatorSource)fd).DocumentPaginator;
            paginator.PageSize = new Size(pageW, pageH);

            XpsDocument.CreateXpsDocumentWriter(_xpsDoc).Write(paginator);

            var fixedSeq = _xpsDoc.GetFixedDocumentSequence();
            PreviewViewer.Document = fixedSeq;

            int pageCount = fixedSeq.References.Count > 0
                ? fixedSeq.References[0].GetDocument(false)?.Pages.Count ?? 0
                : 0;
            PageInfoText.Text = pageCount > 0 ? $"총 {pageCount}페이지" : string.Empty;
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

    private void OnPrintClick(object sender, RoutedEventArgs e)
        => PreviewViewer.Print();

    private void OnCloseClick(object sender, RoutedEventArgs e)
        => Close();

    private void OnWindowClosed(object? sender, EventArgs e)
    {
        PreviewViewer.Document = null;
        _xpsDoc?.Close();
        if (_packUri is not null)
        {
            PackageStore.RemovePackage(_packUri);
            _packUri = null;
        }
        _xpsStream?.Dispose();
    }
}

using System;
using System.IO;
using System.Security.Cryptography;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using PolyDonky.App.Services;
using PolyDonky.Core;
using WpfDocs = System.Windows.Documents;
using WpfMedia = System.Windows.Media;

namespace PolyDonky.App.Views;

/// <summary>
/// 그림 삽입 다이얼로그. 파일을 선택하고 크기(mm)를 지정한 뒤 캐럿 위치 블록 뒤에 ImageBlock 으로 삽입.
/// </summary>
public partial class ImageWindow : Window
{
    private readonly RichTextBox _editor;
    private byte[]? _imageData;
    private string  _mediaType         = "application/octet-stream";
    private double  _naturalWidthMm;
    private double  _naturalHeightMm;
    private bool    _suppressSizeSync;

    public ImageWindow(RichTextBox editor)
    {
        InitializeComponent();
        _editor = editor;
    }

    private void OnBrowseClick(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog
        {
            Title  = "그림 파일 선택",
            Filter = "이미지 파일|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tif;*.tiff;*.webp|모든 파일|*.*",
        };
        if (dlg.ShowDialog(this) != true) return;

        try
        {
            var path   = dlg.FileName;
            var data   = File.ReadAllBytes(path);
            var bitmap = LoadBitmap(data);

            _imageData = data;
            _mediaType = DetectMimeType(path);
            FilePathBox.Text    = path;
            PreviewImage.Source = bitmap;

            double dpiX = bitmap.DpiX > 0 ? bitmap.DpiX : 96.0;
            double dpiY = bitmap.DpiY > 0 ? bitmap.DpiY : 96.0;
            _naturalWidthMm  = bitmap.PixelWidth  / (dpiX / 25.4);
            _naturalHeightMm = bitmap.PixelHeight / (dpiY / 25.4);

            // 본문 최대 너비(A4 기준 160mm) 초과 시 비율 유지하며 축소
            const double maxWidthMm = 160.0;
            double wMm = _naturalWidthMm;
            double hMm = _naturalHeightMm;
            if (wMm > maxWidthMm)
            {
                hMm = hMm * maxWidthMm / wMm;
                wMm = maxWidthMm;
            }

            _suppressSizeSync = true;
            WidthBox.Text  = wMm.ToString("F1");
            HeightBox.Text = hMm.ToString("F1");
            _suppressSizeSync = false;
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"파일을 불러올 수 없습니다.\n{ex.Message}", "그림 삽입",
                MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private void OnWidthChanged(object sender, TextChangedEventArgs e)
    {
        if (_suppressSizeSync || LockRatioCheck.IsChecked != true) return;
        if (!double.TryParse(WidthBox.Text, out double w) || w <= 0) return;
        if (_naturalWidthMm <= 0) return;

        _suppressSizeSync = true;
        HeightBox.Text = (w * _naturalHeightMm / _naturalWidthMm).ToString("F1");
        _suppressSizeSync = false;
    }

    private void OnHeightChanged(object sender, TextChangedEventArgs e)
    {
        if (_suppressSizeSync || LockRatioCheck.IsChecked != true) return;
        if (!double.TryParse(HeightBox.Text, out double h) || h <= 0) return;
        if (_naturalHeightMm <= 0) return;

        _suppressSizeSync = true;
        WidthBox.Text = (h * _naturalWidthMm / _naturalHeightMm).ToString("F1");
        _suppressSizeSync = false;
    }

    private void OnInsertClick(object sender, RoutedEventArgs e)
    {
        if (_imageData is null || _imageData.Length == 0)
        {
            MessageBox.Show(this, "그림 파일을 먼저 선택하세요.", "그림 삽입",
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }
        if (!double.TryParse(WidthBox.Text,  out double wMm) || wMm <= 0 ||
            !double.TryParse(HeightBox.Text, out double hMm) || hMm <= 0)
        {
            MessageBox.Show(this, "너비와 높이를 올바르게 입력하세요 (단위: mm).", "그림 삽입",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var hash = SHA256.HashData(_imageData);

        var image = new ImageBlock
        {
            Data        = _imageData,
            MediaType   = _mediaType,
            WidthMm     = wMm,
            HeightMm    = hMm,
            Sha256      = Convert.ToHexString(hash).ToLowerInvariant(),
            Description = DescriptionBox.Text.Trim() is { Length: > 0 } d ? d : null,
        };

        InsertImageBlock(_editor, image);
        DialogResult = true;
        Close();
    }

    private void OnCloseClick(object sender, RoutedEventArgs e) => Close();

    // ── 삽입 헬퍼 ────────────────────────────────────────────────────────

    /// <summary>
    /// 편집기의 캐럿 위치 블록 바로 뒤에 ImageBlock 을 BlockUIContainer 로 삽입한다.
    /// Tag = ImageBlock 을 심어 FlowDocumentParser 의 라운드트립을 보장.
    /// </summary>
    public static void InsertImageBlock(RichTextBox editor, ImageBlock image)
    {
        // WrapMode 에 따라 BlockUIContainer 또는 Paragraph(Floater 포함) 가 반환된다.
        var imageBlock = FlowDocumentBuilder.BuildImage(image);
        var flowDoc    = editor.Document;
        var current    = FindEnclosingBlock(editor.CaretPosition, flowDoc);

        if (current is not null)
            flowDoc.Blocks.InsertAfter(current, imageBlock);
        else
            flowDoc.Blocks.Add(imageBlock);

        try { editor.CaretPosition = imageBlock.ContentEnd; }
        catch { /* 포지션 이동 실패는 무시 */ }

        editor.Focus();
    }

    /// <summary>TextPointer 를 감싸는 최상위 Block 을 FlowDocument 에서 찾는다.</summary>
    private static WpfDocs.Block? FindEnclosingBlock(WpfDocs.TextPointer pos, WpfDocs.FlowDocument doc)
    {
        // 빠른 경로: Paragraph.ElementStart 로 직접 찾기
        if (pos.Paragraph is { } para)
        {
            foreach (var b in doc.Blocks)
                if (b == para) return para;
        }

        // 범위 기반 순차 탐색 (BlockUIContainer, Table 등 포함)
        foreach (var b in doc.Blocks)
        {
            if (b.ContentStart.CompareTo(pos) <= 0 && pos.CompareTo(b.ContentEnd) <= 0)
                return b;
        }
        return null;
    }

    // ── 유틸 ─────────────────────────────────────────────────────────────

    private static WpfMedia.Imaging.BitmapImage LoadBitmap(byte[] data)
    {
        var bmp = new WpfMedia.Imaging.BitmapImage();
        // OnLoad + 명시 Dispose — EndInit 후 내부 캐시에 데이터가 복사되므로 원본 stream 해제 안전.
        var ms = new MemoryStream(data, writable: false);
        bmp.BeginInit();
        bmp.CacheOption  = WpfMedia.Imaging.BitmapCacheOption.OnLoad;
        bmp.StreamSource = ms;
        bmp.EndInit();
        ms.Dispose();
        bmp.Freeze();
        return bmp;
    }

    private static string DetectMimeType(string filePath) =>
        Path.GetExtension(filePath).ToLowerInvariant() switch
        {
            ".jpg" or ".jpeg" => "image/jpeg",
            ".png"            => "image/png",
            ".gif"            => "image/gif",
            ".bmp"            => "image/bmp",
            ".tif" or ".tiff" => "image/tiff",
            ".webp"           => "image/webp",
            _                 => "application/octet-stream",
        };
}

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Input;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Win32;
using PolyDoc.App.Models;
using PolyDoc.App.Services;
using PolyDoc.App.Views;
using PolyDoc.Core;
using SR = PolyDoc.App.Properties.Resources;
using Wpf = System.Windows.Documents;

namespace PolyDoc.App.ViewModels;

public partial class MainViewModel : ObservableObject
{
    /// <summary>편집 모델의 source-of-truth 인 PolyDocument. FlowDocument 와는 Builder/Parser 로 동기화.</summary>
    private PolyDocument _document = PolyDocument.Empty();

    [ObservableProperty]
    private Wpf.FlowDocument _flowDocument = new();

    [ObservableProperty]
    private string _documentTitle = SR.DlgNewDocTitle;

    [ObservableProperty]
    private string _currentFilePath = string.Empty;

    [ObservableProperty]
    private string _statusMessage = SR.StatusReady;

    [ObservableProperty]
    private bool _hasUnsavedChanges;

    // 상태 표시줄 우측 4칸: 메모리·Insert/Overwrite·CapsLock·NumLock.
    // MainWindow code-behind 가 DispatcherTimer 로 1초마다 RefreshSystemKeys/Memory 를 호출한다.
    [ObservableProperty]
    private string _memoryUsage = "0.0 MB";

    [ObservableProperty]
    private string _insertModeText = SR.StatusInsert;

    [ObservableProperty]
    private string _capsLockText = "    ";

    [ObservableProperty]
    private string _numLockText = "Num";

    /// <summary>true 면 Overwrite (수정) 모드. 표시 외 실제 overwrite 동작은 후속 사이클.</summary>
    public bool IsOverwriteMode { get; private set; }

    public void ToggleInsertMode()
    {
        IsOverwriteMode = !IsOverwriteMode;
        InsertModeText = IsOverwriteMode ? SR.StatusOverwrite : SR.StatusInsert;
    }

    public void RefreshSystemKeys()
    {
        var caps = Keyboard.IsKeyToggled(Key.CapsLock);
        var num = Keyboard.IsKeyToggled(Key.NumLock);
        CapsLockText = caps ? "Caps" : "    ";
        NumLockText = num ? "Num " : "    ";
    }

    public void RefreshMemoryUsage()
    {
        var bytes = DocumentMeasurement.EstimateBytes(_document);
        MemoryUsage = DocumentMeasurement.FormatBytes(bytes);
    }

    public string WindowTitle
        => HasUnsavedChanges
            ? $"PolyDoc — {DocumentTitle} *"
            : $"PolyDoc — {DocumentTitle}";

    partial void OnDocumentTitleChanged(string value)
        => OnPropertyChanged(nameof(WindowTitle));

    partial void OnHasUnsavedChangesChanged(bool value)
        => OnPropertyChanged(nameof(WindowTitle));

    public void MarkDirty()
    {
        if (_suppressDirty) return;
        HasUnsavedChanges = true;
    }

    private bool _suppressDirty;

    private void LoadDocument(PolyDocument document, string? path)
    {
        _document = document;
        _suppressDirty = true;
        FlowDocument = FlowDocumentBuilder.Build(document);
        _suppressDirty = false;
        CurrentFilePath = path ?? string.Empty;
        DocumentTitle = string.IsNullOrEmpty(path) ? SR.DlgNewDocTitle : Path.GetFileName(path);
        HasUnsavedChanges = false;
    }

    [RelayCommand]
    private void New()
    {
        if (!ConfirmDiscardChanges()) return;
        LoadDocument(PolyDocument.Empty(), null);
        StatusMessage = SR.StatusNewDoc;
    }

    [RelayCommand]
    private void Open()
    {
        if (!ConfirmDiscardChanges()) return;

        var dlg = new OpenFileDialog
        {
            Filter = KnownFormats.OpenFilter,
            Title = SR.DlgFileOpenTitle,
        };
        if (dlg.ShowDialog() != true) return;

        OpenPath(dlg.FileName);
    }

    /// <summary>드래그&드롭 등 외부에서 파일 경로를 직접 전달받아 열기.</summary>
    public void OpenFile(string path)
    {
        if (!ConfirmDiscardChanges()) return;
        OpenPath(path);
    }

    private void OpenPath(string path)
    {
        if (KnownFormats.RequiresExternalConverter(path) && !KnownFormats.IsSupportedNatively(path))
        {
            var ext = Path.GetExtension(path).TrimStart('.').ToLowerInvariant();
            MessageBox.Show(
                string.Format(SR.DlgUnsupportedFormat, ext),
                SR.DlgUnsupportedFormatTitle,
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var reader = KnownFormats.PickReader(path);
        if (reader is null)
        {
            MessageBox.Show(SR.DlgUnknownFormat, SR.DlgOpenFail,
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        try
        {
            using var fs = File.OpenRead(path);
            var doc = reader.Read(fs);
            LoadDocument(doc, path);
            StatusMessage = BuildOpenStatusMessage(path, doc);
        }
        catch (Exception ex)
        {
            ReportError(SR.DlgOpenError, ex);
        }
    }

    private static string BuildOpenStatusMessage(string path, PolyDocument doc)
    {
        var name = Path.GetFileName(path);
        if (doc.Metadata.Custom.TryGetValue("hwpx.paragraphCount", out var pCount)
            && int.TryParse(pCount, out var pc) && pc == 0)
        {
            doc.Metadata.Custom.TryGetValue("hwpx.sectionFilesFound", out var sCount);
            doc.Metadata.Custom.TryGetValue("hwpx.firstSectionEntryHit", out var hit);
            doc.Metadata.Custom.TryGetValue("hwpx.firstSectionPath", out var firstPath);
            doc.Metadata.Custom.TryGetValue("hwpx.firstSectionRoot", out var rootName);
            doc.Metadata.Custom.TryGetValue("hwpx.firstSectionTags", out var tags);
            doc.Metadata.Custom.TryGetValue("hwpx.xmlEntries", out var xmlEntries);
            doc.Metadata.Custom.TryGetValue("hwpx.parseErrors", out var parseErrors);
            var errPart = string.IsNullOrEmpty(parseErrors) ? "" : $" | XML 오류: {parseErrors}";
            return $"열기 완료 — {name} | 본문 0건, 섹션 {sCount ?? "?"}개 (entry-hit={hit ?? "?"}) | path={firstPath ?? "?"} | root=<{rootName ?? "?"}> | 자식: {tags ?? "(없음)"} | xml: {xmlEntries ?? "(없음)"}{errPart}";
        }
        if (doc.Metadata.Custom.TryGetValue("hwpx.parseErrors", out var perr) && !string.IsNullOrEmpty(perr))
        {
            return $"열기 완료 — {name} (일부 XML 파싱 경고: {perr})";
        }
        return string.Format(SR.StatusOpenDone, name);
    }

    [RelayCommand]
    private void Save()
    {
        if (string.IsNullOrEmpty(CurrentFilePath))
        {
            SaveAs();
            return;
        }
        SaveTo(CurrentFilePath);
    }

    [RelayCommand]
    private void SaveAs()
    {
        var dlg = new SaveFileDialog
        {
            Filter = KnownFormats.SaveFilter,
            Title = SR.DlgFileSaveTitle,
            FileName = string.IsNullOrEmpty(CurrentFilePath)
                ? SR.DlgDefaultFileName
                : Path.GetFileName(CurrentFilePath),
        };
        if (dlg.ShowDialog() != true) return;

        if (KnownFormats.RequiresExternalConverter(dlg.FileName) && !KnownFormats.IsSupportedNatively(dlg.FileName))
        {
            var ext = Path.GetExtension(dlg.FileName).TrimStart('.').ToLowerInvariant();
            MessageBox.Show(
                string.Format(SR.DlgSaveConverterNeeded, ext),
                SR.DlgSaveConverterTitle,
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        SaveTo(dlg.FileName);
    }

    [RelayCommand]
    private void About()
    {
        var window = new AboutWindow
        {
            Owner = Application.Current.MainWindow,
        };
        window.ShowDialog();
    }

    [RelayCommand]
    private void DocInfo()
    {
        var text = _document.ToPlainText();
        var charsNoSpace = text.Replace(" ", "").Replace("\n", "").Replace("\r", "").Length;
        var charsWithSpace = text.Replace("\n", "").Replace("\r", "").Length;
        var words = text.Split(new[] { ' ', '\n', '\r', '\t' },
                               StringSplitOptions.RemoveEmptyEntries).Length;
        var lines = text.Split('\n').Length;

        var (paras, tables, images) = CountBlocks(_document.Sections);
        var bytes = DocumentMeasurement.EstimateBytes(_document);
        var meta  = _document.Metadata;
        var none  = SR.DocInfoNone;

        var path   = string.IsNullOrEmpty(CurrentFilePath) ? SR.DocInfoNotSaved : CurrentFilePath;
        var format = string.IsNullOrEmpty(CurrentFilePath)
            ? SR.DlgNewDocTitle
            : Path.GetExtension(CurrentFilePath).TrimStart('.').ToUpperInvariant();

        var info = new DocumentInfoModel
        {
            FilePath       = path,
            Format         = format,
            DataSize       = DocumentMeasurement.FormatBytes(bytes),
            DocTitle       = string.IsNullOrWhiteSpace(meta.Title)  ? none : meta.Title,
            Author         = string.IsNullOrWhiteSpace(meta.Author) ? none : meta.Author,
            Language       = meta.Language,
            Created        = meta.Created.LocalDateTime.ToString("yyyy-MM-dd HH:mm"),
            Modified       = meta.Modified.LocalDateTime.ToString("yyyy-MM-dd HH:mm"),
            ParagraphCount = paras.ToString("N0"),
            CharCount      = $"{charsNoSpace:N0}  ({charsWithSpace:N0} 공백 포함)",
            WordCount      = words.ToString("N0"),
            LineCount      = lines.ToString("N0"),
            SectionCount   = _document.Sections.Count.ToString("N0"),
            TableCount     = tables.ToString("N0"),
            ImageCount     = images.ToString("N0"),
        };

        var dlg = new DocumentInfoWindow(info) { Owner = Application.Current.MainWindow };
        dlg.ShowDialog();
    }

    private static (int paragraphs, int tables, int images) CountBlocks(IEnumerable<Section> sections)
    {
        int p = 0, t = 0, i = 0;
        foreach (var section in sections)
        {
            var (sp, st, si) = CountBlocks(section.Blocks);
            p += sp; t += st; i += si;
        }
        return (p, t, i);
    }

    private static (int paragraphs, int tables, int images) CountBlocks(IList<Block> blocks)
    {
        int p = 0, t = 0, i = 0;
        foreach (var block in blocks)
        {
            switch (block)
            {
                case Paragraph:
                    p++;
                    break;
                case Table table:
                    t++;
                    foreach (var row in table.Rows)
                        foreach (var cell in row.Cells)
                        {
                            var (cp, ct, ci) = CountBlocks(cell.Blocks);
                            p += cp; t += ct; i += ci;
                        }
                    break;
                case ImageBlock:
                    i++;
                    break;
            }
        }
        return (p, t, i);
    }

    [RelayCommand]
    private void FindReplace()
    {
        FindReplaceRequested?.Invoke(this, EventArgs.Empty);
    }

    public event EventHandler? FindReplaceRequested;

    [RelayCommand]
    private void Settings()
    {
        SettingsRequested?.Invoke(this, EventArgs.Empty);
    }

    public event EventHandler? SettingsRequested;

    [RelayCommand]
    private void Close()
    {
        if (!ConfirmDiscardChanges()) return;
        LoadDocument(PolyDocument.Empty(), null);
        StatusMessage = SR.StatusDocClosed;
    }

    [RelayCommand]
    private void Exit()
    {
        if (!ConfirmDiscardChanges()) return;
        Application.Current.Shutdown();
    }

    private void SaveTo(string path)
    {
        var writer = KnownFormats.PickWriter(path);
        if (writer is null)
        {
            MessageBox.Show(SR.DlgSaveFail, SR.DlgSaveFailTitle,
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        try
        {
            var rebuilt = FlowDocumentParser.Parse(FlowDocument, _document);

            using var fs = File.Create(path);
            writer.Write(rebuilt, fs);

            _document = rebuilt;
            CurrentFilePath = path;
            DocumentTitle = Path.GetFileName(path);
            HasUnsavedChanges = false;
            StatusMessage = string.Format(SR.StatusSaveDone, Path.GetFileName(path));
        }
        catch (Exception ex)
        {
            ReportError(SR.DlgSaveError, ex);
        }
    }

    private bool ConfirmDiscardChanges()
    {
        if (!HasUnsavedChanges) return true;
        var result = MessageBox.Show(
            SR.DlgUnsavedPrompt,
            "PolyDoc",
            MessageBoxButton.YesNoCancel,
            MessageBoxImage.Question);
        if (result == MessageBoxResult.Cancel) return false;
        if (result == MessageBoxResult.Yes)
        {
            Save();
            return !HasUnsavedChanges;
        }
        return true;
    }

    private void ReportError(string headline, Exception ex)
    {
        StatusMessage = $"오류: {ex.Message}";
        MessageBox.Show($"{headline}\n\n{ex.GetType().Name}: {ex.Message}",
            "PolyDoc", MessageBoxButton.OK, MessageBoxImage.Error);
    }
}

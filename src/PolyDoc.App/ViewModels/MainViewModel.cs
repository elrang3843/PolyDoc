using System;
using System.IO;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Win32;
using PolyDoc.App.Services;
using PolyDoc.App.Views;
using PolyDoc.Core;
using Wpf = System.Windows.Documents;

namespace PolyDoc.App.ViewModels;

public partial class MainViewModel : ObservableObject
{
    /// <summary>편집 모델의 source-of-truth 인 PolyDocument. FlowDocument 와는 Builder/Parser 로 동기화.</summary>
    private PolyDocument _document = PolyDocument.Empty();

    [ObservableProperty]
    private Wpf.FlowDocument _flowDocument = new();

    [ObservableProperty]
    private string _documentTitle = "(새 문서)";

    [ObservableProperty]
    private string _currentFilePath = string.Empty;

    [ObservableProperty]
    private string _statusMessage = "준비됨";

    [ObservableProperty]
    private bool _hasUnsavedChanges;

    public string WindowTitle
        => HasUnsavedChanges
            ? $"PolyDoc — {DocumentTitle} *"
            : $"PolyDoc — {DocumentTitle}";

    partial void OnDocumentTitleChanged(string value)
        => OnPropertyChanged(nameof(WindowTitle));

    partial void OnHasUnsavedChangesChanged(bool value)
        => OnPropertyChanged(nameof(WindowTitle));

    /// <summary>RichTextBox 에서 본문 변경이 일어나면 code-behind 가 호출. Open/Save 등 프로그램적 변경 중에는 _suppressDirty 로 무시.</summary>
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
        DocumentTitle = string.IsNullOrEmpty(path) ? "(새 문서)" : Path.GetFileName(path);
        HasUnsavedChanges = false;
    }

    [RelayCommand]
    private void New()
    {
        if (!ConfirmDiscardChanges()) return;
        LoadDocument(PolyDocument.Empty(), null);
        StatusMessage = "새 문서가 생성되었습니다.";
    }

    [RelayCommand]
    private void Open()
    {
        if (!ConfirmDiscardChanges()) return;

        var dlg = new OpenFileDialog
        {
            Filter = KnownFormats.OpenFilter,
            Title = "문서 열기",
        };
        if (dlg.ShowDialog() != true) return;

        var path = dlg.FileName;

        if (KnownFormats.RequiresExternalConverter(path) && !KnownFormats.IsSupportedNatively(path))
        {
            var ext = Path.GetExtension(path).TrimStart('.').ToLowerInvariant();
            MessageBox.Show(
                $"이 형식 (.{ext}) 은 외부 컨버터 모듈이 필요합니다 (Phase D 이후 지원).\n\n" +
                "현재 PolyDoc 이 직접 열 수 있는 형식: IWPF, DOCX, Markdown, TXT.",
                "지원되지 않는 형식",
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var reader = KnownFormats.PickReader(path);
        if (reader is null)
        {
            MessageBox.Show("알 수 없는 파일 형식입니다.", "열기 실패",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        try
        {
            using var fs = File.OpenRead(path);
            var doc = reader.Read(fs);
            LoadDocument(doc, path);
            StatusMessage = $"열기 완료 — {Path.GetFileName(path)}";
        }
        catch (Exception ex)
        {
            ReportError("문서를 여는 중 오류가 발생했습니다.", ex);
        }
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
            Title = "다른 이름으로 저장",
            FileName = string.IsNullOrEmpty(CurrentFilePath)
                ? "문서.iwpf"
                : Path.GetFileName(CurrentFilePath),
        };
        if (dlg.ShowDialog() != true) return;

        if (KnownFormats.RequiresExternalConverter(dlg.FileName) && !KnownFormats.IsSupportedNatively(dlg.FileName))
        {
            var ext = Path.GetExtension(dlg.FileName).TrimStart('.').ToLowerInvariant();
            MessageBox.Show(
                $"선택하신 형식 (.{ext}) 은 외부 컨버터를 통한 변환이 필요합니다 (Phase D 이후 지원).\n\n" +
                "지금은 PolyDoc 이 직접 처리하는 형식(IWPF, DOCX, Markdown, TXT)으로 저장해 주세요.",
                "외부 컨버터 필요",
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
            MessageBox.Show("이 형식은 PolyDoc 이 직접 저장할 수 없습니다.",
                "저장 실패", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        try
        {
            // FlowDocument 를 다시 PolyDocument 로 회수. 원본을 머지 베이스로 넘겨
            // 메타·페이지 설정·한글 조판 속성을 비파괴 보존한다.
            var rebuilt = FlowDocumentParser.Parse(FlowDocument, _document);

            using var fs = File.Create(path);
            writer.Write(rebuilt, fs);

            _document = rebuilt;
            CurrentFilePath = path;
            DocumentTitle = Path.GetFileName(path);
            HasUnsavedChanges = false;
            StatusMessage = $"저장 완료 — {Path.GetFileName(path)}";
        }
        catch (Exception ex)
        {
            ReportError("문서를 저장하는 중 오류가 발생했습니다.", ex);
        }
    }

    private bool ConfirmDiscardChanges()
    {
        if (!HasUnsavedChanges) return true;
        var result = MessageBox.Show(
            "저장하지 않은 변경 사항이 있습니다. 저장하시겠습니까?",
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

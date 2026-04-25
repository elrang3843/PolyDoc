using System;
using System.IO;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Win32;
using PolyDoc.App.Services;
using PolyDoc.App.Views;
using PolyDoc.Codecs.Text;
using PolyDoc.Core;

namespace PolyDoc.App.ViewModels;

public partial class MainViewModel : ObservableObject
{
    private PolyDocument _document = PolyDocument.Empty();

    [ObservableProperty]
    private string _documentTitle = "(새 문서)";

    [ObservableProperty]
    private string _documentBody = string.Empty;

    [ObservableProperty]
    private string _currentFilePath = string.Empty;

    [ObservableProperty]
    private string _statusMessage = "준비됨";

    [ObservableProperty]
    private bool _hasUnsavedChanges;

    public PolyDocument Document
    {
        get => _document;
        private set
        {
            _document = value;
            DocumentBody = PlainTextWriter.ToText(value);
            DocumentTitle = string.IsNullOrEmpty(CurrentFilePath)
                ? "(새 문서)"
                : Path.GetFileName(CurrentFilePath);
            HasUnsavedChanges = false;
            OnPropertyChanged(nameof(WindowTitle));
        }
    }

    public string WindowTitle
        => HasUnsavedChanges
            ? $"PolyDoc — {DocumentTitle} *"
            : $"PolyDoc — {DocumentTitle}";

    partial void OnDocumentBodyChanged(string value)
    {
        // 사용자가 본문을 편집한 시점부터 dirty 표시.
        if (!_suppressDirty)
        {
            HasUnsavedChanges = true;
            OnPropertyChanged(nameof(WindowTitle));
        }
    }

    partial void OnDocumentTitleChanged(string value)
        => OnPropertyChanged(nameof(WindowTitle));

    partial void OnHasUnsavedChangesChanged(bool value)
        => OnPropertyChanged(nameof(WindowTitle));

    private bool _suppressDirty;

    [RelayCommand]
    private void New()
    {
        if (!ConfirmDiscardChanges()) return;
        CurrentFilePath = string.Empty;
        _suppressDirty = true;
        Document = PolyDocument.Empty();
        _suppressDirty = false;
        StatusMessage = "새 문서가 생성되었습니다.";
    }

    [RelayCommand]
    private void Open()
    {
        if (!ConfirmDiscardChanges()) return;

        var dlg = new OpenFileDialog
        {
            Filter = DocumentFormat.OpenFilter,
            Title = "문서 열기",
        };
        if (dlg.ShowDialog() != true) return;

        var path = dlg.FileName;

        if (DocumentFormat.RequiresExternalConverter(path) && !DocumentFormat.IsSupportedNatively(path))
        {
            MessageBox.Show(
                "이 형식은 외부 컨버터 모듈이 필요합니다 (Phase D 이후 지원).\n" +
                "현재는 IWPF, Markdown, TXT 만 직접 열 수 있습니다.",
                "지원되지 않는 형식",
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var reader = DocumentFormat.PickReader(path);
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
            CurrentFilePath = path;
            _suppressDirty = true;
            Document = doc;
            _suppressDirty = false;
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
            Filter = DocumentFormat.SaveFilter,
            Title = "다른 이름으로 저장",
            FileName = string.IsNullOrEmpty(CurrentFilePath) ? "문서.iwpf" : Path.GetFileName(CurrentFilePath),
        };
        if (dlg.ShowDialog() != true) return;

        // 외부 포맷 저장 시 한 번 더 확인.
        if (DocumentFormat.RequiresExternalConverter(dlg.FileName))
        {
            var ok = MessageBox.Show(
                "선택하신 형식은 외부 컨버터를 통한 변환이 필요합니다 (Phase D 이후 지원).\n" +
                "지금은 PolyDoc 내장 형식(IWPF/MD/TXT)으로 저장해 주세요.",
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
        var writer = DocumentFormat.PickWriter(path);
        if (writer is null)
        {
            MessageBox.Show("이 형식은 PolyDoc 이 직접 저장할 수 없습니다.",
                "저장 실패", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        // 본문 편집 결과를 PolyDocument 로 다시 합친다 (Phase B 첫 사이클: plain text 모델).
        // Phase E 의 리치 편집 도입 후에는 이 로직이 ViewModel <-> 도큐먼트 모델 동기화로 대체된다.
        var rebuilt = PlainTextReader.FromText(DocumentBody);
        rebuilt.Metadata.Title = string.IsNullOrEmpty(CurrentFilePath)
            ? Path.GetFileNameWithoutExtension(path)
            : Path.GetFileNameWithoutExtension(CurrentFilePath);

        try
        {
            using var fs = File.Create(path);
            writer.Write(rebuilt, fs);
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

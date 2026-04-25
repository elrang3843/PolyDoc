using System;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Input;
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

    // 상태 표시줄 우측 4칸: 메모리·Insert/Overwrite·CapsLock·NumLock.
    // MainWindow code-behind 가 DispatcherTimer 로 1초마다 RefreshSystemKeys/Memory 를 호출한다.
    [ObservableProperty]
    private string _memoryUsage = "0.0 MB";

    [ObservableProperty]
    private string _insertModeText = "삽입";

    [ObservableProperty]
    private string _capsLockText = "    ";

    [ObservableProperty]
    private string _numLockText = "Num";

    /// <summary>true 면 Overwrite (수정) 모드. 표시 외 실제 overwrite 동작은 후속 사이클.</summary>
    public bool IsOverwriteMode { get; private set; }

    public void ToggleInsertMode()
    {
        IsOverwriteMode = !IsOverwriteMode;
        InsertModeText = IsOverwriteMode ? "수정" : "삽입";
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
        // "이 문서가 차지하는 데이터 크기" — 앱 전체 메모리가 아니라 본문 콘텐츠 기준.
        // 본문이 비어 있거나 문서를 새로 만든 직후에는 자연스럽게 0 가까이 떨어진다.
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
            StatusMessage = BuildOpenStatusMessage(path, doc);
        }
        catch (Exception ex)
        {
            ReportError("문서를 여는 중 오류가 발생했습니다.", ex);
        }
    }

    private static string BuildOpenStatusMessage(string path, PolyDocument doc)
    {
        var name = Path.GetFileName(path);
        // HWPX reader 가 metadata.Custom 에 진단을 박았으면 본문 인식이 0건일 때 사용자에게 경고.
        if (doc.Metadata.Custom.TryGetValue("hwpx.paragraphCount", out var pCount)
            && int.TryParse(pCount, out var pc) && pc == 0)
        {
            doc.Metadata.Custom.TryGetValue("hwpx.sectionFilesFound", out var sCount);
            return $"열기 완료 — {name} (HWPX 본문 인식 0건, 섹션 파일 {sCount ?? "?"}개. 한컴 변종 가능 — 진단 정보를 메인테이너에게 공유 부탁)";
        }
        return $"열기 완료 — {name}";
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

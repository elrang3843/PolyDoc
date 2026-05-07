using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Win32;
using PolyDonky.App.Models;
using PolyDonky.App.Services;
using PolyDonky.App.Views;
using PolyDonky.Core;
using PolyDonky.Iwpf;
using SR = PolyDonky.App.Properties.Resources;
using Wpf = System.Windows.Documents;

namespace PolyDonky.App.ViewModels;

public partial class MainViewModel : ObservableObject
{
    /// <summary>편집 모델의 source-of-truth 인 PolyDonkyument. FlowDocument 와는 Builder/Parser 로 동기화.</summary>
    private PolyDonkyument _document = PolyDonkyument.Empty();

    /// <summary>현재 편집 중인 문서 (OutlineStyles 등 접근용).</summary>
    public PolyDonkyument Document => _document;

    /// <summary>
    /// 저장 직전에 호출해 현재 에디터 상태를 반영한 PolyDonkyument 를 얻는다.
    /// per-page RTB 모드에서 MainWindow 가 등록한다. null 이면 FlowDocument 파싱으로 폴백.
    /// </summary>
    public Func<PolyDonkyument?>? LiveDocumentProvider { get; set; }

    /// <summary>편집 작업의 Undo/Redo 스냅샷 스택. MainWindow 가 액션 시점에 PushUndo 한다.</summary>
    public UndoRedoManager UndoRedo { get; } = new();

    /// <summary>Undo 가능 여부 (메뉴 활성/비활성 바인딩용).</summary>
    public bool CanUndo => UndoRedo.CanUndo;

    /// <summary>Redo 가능 여부 (메뉴 활성/비활성 바인딩용).</summary>
    public bool CanRedo => UndoRedo.CanRedo;

    public MainViewModel()
    {
        UndoRedo.StateChanged += (_, _) =>
        {
            OnPropertyChanged(nameof(CanUndo));
            OnPropertyChanged(nameof(CanRedo));
        };
    }

    /// <summary>
    /// Undo/Redo 가 호출 직후 사용. 모델을 통째로 교체하고 FlowDocument 를 재빌드한다.
    /// HasUnsavedChanges 는 호출자가 따로 관리(보통 true 로 유지).
    /// </summary>
    internal void ReplaceDocumentForUndo(PolyDonkyument document)
    {
        ArgumentNullException.ThrowIfNull(document);
        _document = document;
        _suppressDirty = true;
        FlowDocument = FlowDocumentBuilder.Build(document);
        _suppressDirty = false;
        // RTB·오버레이 재구성은 MainWindow.ApplyFlowDocument 가 처리.
    }

    /// <summary>
    /// 텍스트 입력 burst 종료 시 호출 — 라이브로 파싱한 결과를 _document 에 동기화.
    /// MarkDirty 를 트리거하지 않는다 (이미 dirty).
    /// </summary>
    internal void SyncDocumentFromLive(PolyDonkyument live)
    {
        ArgumentNullException.ThrowIfNull(live);
        _suppressDirty = true;
        _document = live;
        _suppressDirty = false;
    }

    /// <summary>
    /// 열기 보호(Read/Both)에 사용할 비밀번호. null/빈 문자열이면 평문 저장.
    /// 메모리 외 어디에도 저장되지 않는다 (ViewModel 인스턴스 생명 주기 내에서만 유효).
    /// </summary>
    private string? _documentPassword;

    /// <summary>쓰기 보호(Write/Both)가 설정된 경우의 잠금 레코드. null 이면 쓰기 보호 없음.</summary>
    private IwpfWriteLock? _writeLock;

    /// <summary>
    /// 쓰기 보호 비밀번호 캐시. 세션 내 최초 저장 시 검증 후 저장해 재입력을 방지한다.
    /// LoadDocument 시 초기화. (Read 모드에서는 사용 안 함.)
    /// </summary>
    private string? _writePassword;

    /// <summary>
    /// true 면 현재 문서가 쓰기 보호 상태이며 편집이 잠겨 있다.
    /// (_writeLock 이 있고 아직 비밀번호 검증이 안 된 경우.)
    /// View 가 이 값을 RichTextBox.IsReadOnly 에 바인딩 + 상태 표시줄 인디케이터로 사용한다.
    /// </summary>
    [ObservableProperty]
    private bool _isWriteProtected;

    private void RecomputeWriteProtection()
        => IsWriteProtected = _writeLock is not null && _writePassword is null;

    private PasswordMode CurrentPasswordMode =>
        (!string.IsNullOrEmpty(_documentPassword), _writeLock is not null) switch
        {
            (true,  true)  => PasswordMode.Both,
            (true,  false) => PasswordMode.Read,
            (false, true)  => PasswordMode.Write,
            _              => PasswordMode.None,
        };

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

    /// <summary>
    /// 파일 입출력 등 백그라운드 작업 진행 중 상태. true 일 때 상태 표시줄의 진행 막대가 보인다.
    /// 사용자 입력은 차단되지 않으며, 같은 작업이 동시에 시작되지 않게 호출자에서 직접 가드한다.
    /// </summary>
    [ObservableProperty]
    private bool _isBusy;

    /// <summary>진행 중 작업 설명 — 상태 표시줄의 진행 막대 옆에 표시.</summary>
    [ObservableProperty]
    private string _busyMessage = string.Empty;

    /// <summary>
    /// 외부 CLI 가 보고하는 진행률 (0~100). -1 이면 진행률 미정 → ProgressBar 가 indeterminate 모드.
    /// </summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(BusyProgressIsIndeterminate))]
    private int _busyProgress = -1;

    /// <summary>BusyProgress 가 음수면 indeterminate. ProgressBar.IsIndeterminate 에 바인딩.</summary>
    public bool BusyProgressIsIndeterminate => BusyProgress < 0;

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
            ? $"PolyDonky — {DocumentTitle} *"
            : $"PolyDonky — {DocumentTitle}";

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

    private void LoadDocument(PolyDonkyument document, string? path,
                              string? password = null, IwpfWriteLock? writeLock = null)
    {
        _document = document;
        _documentPassword = password;
        _writeLock = writeLock;
        _writePassword = null; // 파일 로드 시 항상 초기화 — 편집/저장 전에 재검증 필요
        RecomputeWriteProtection();
        _suppressDirty = true;
        FlowDocument = FlowDocumentBuilder.Build(document);
        _suppressDirty = false;
        CurrentFilePath = path ?? string.Empty;
        DocumentTitle = string.IsNullOrEmpty(path) ? SR.DlgNewDocTitle : Path.GetFileName(path);
        HasUnsavedChanges = false;
        UndoRedo.Clear(); // 새 문서·로드된 문서는 이전 편집 이력과 무관
    }

    [RelayCommand]
    private void New()
    {
        if (!ConfirmDiscardChanges()) return;
        LoadDocument(PolyDonkyument.Empty(), null, password: null);
        StatusMessage = SR.StatusNewDoc;
    }

    [RelayCommand]
    private async Task OpenAsync()
    {
        if (!ConfirmDiscardChanges()) return;

        var dlg = new OpenFileDialog
        {
            Filter = KnownFormats.OpenFilter,
            Title = SR.DlgFileOpenTitle,
        };
        if (dlg.ShowDialog() != true) return;

        await OpenPathAsync(dlg.FileName);
    }

    /// <summary>드래그&드롭 등 외부에서 파일 경로를 직접 전달받아 열기.</summary>
    public async Task OpenFileAsync(string path)
    {
        if (!ConfirmDiscardChanges()) return;
        await OpenPathAsync(path);
    }

    /// <summary>비-async 호출 사이트(이벤트 핸들러 등) 호환을 위한 동기 래퍼 — fire-and-forget.</summary>
    public void OpenFile(string path) => _ = OpenFileAsync(path);

    private async Task OpenPathAsync(string path)
    {
        if (IsBusy) return;  // 동시 열기 가드

        // CLAUDE.md §3 — 외부 CLI 변환기로 IWPF 우회 처리하는 포맷 (HTML/XML/HWP/DOC 등).
        // 메인 앱에서 직접 처리 불가 → CLI spawn 후 임시 IWPF 를 읽는다.
        if (KnownFormats.RequiresExternalConverter(path))
        {
            await OpenViaExternalConverterAsync(path);
            return;
        }

        var reader = KnownFormats.PickReader(path);
        if (reader is null)
        {
            MessageBox.Show(SR.DlgUnknownFormat, SR.DlgOpenFail,
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        PolyDonkyument? doc;
        string? usedPassword = null;

        IsBusy       = true;
        BusyMessage  = string.Format(SR.StatusBusyOpen, Path.GetFileName(path));
        try
        {
            try
            {
                doc = await Task.Run(() =>
                {
                    using var fs = File.OpenRead(path);
                    return reader.Read(fs);
                });
            }
            catch (EncryptedIwpfException) when (reader is IwpfReader iwpf)
            {
                // 암호화된 IWPF — 비밀번호 입력 + 재시도 루프 (UI 스레드).
                doc = ReadEncryptedWithPrompt(iwpf, path, out usedPassword);
                if (doc is null) return; // 사용자 취소
            }
            catch (Exception ex)
            {
                ReportError(SR.DlgOpenError, ex);
                return;
            }

            // IwpfReader 가 Metadata.Custom 에 넣어둔 write-lock 데이터를 꺼낸다.
            IwpfWriteLock? writeLock = null;
            if (doc.Metadata.Custom.TryGetValue("iwpf.writeLock", out var wlJson))
            {
                writeLock = JsonSerializer.Deserialize<IwpfWriteLock>(wlJson, JsonDefaults.Options);
                doc.Metadata.Custom.Remove("iwpf.writeLock");
            }

            LoadDocument(doc, path, usedPassword, writeLock);
            StatusMessage = BuildOpenStatusMessage(path, doc)
                + (IsWriteProtected ? "  " + SR.StatusWriteProtectedSuffix : "");
        }
        finally
        {
            IsBusy      = false;
            BusyMessage = string.Empty;
        }
    }

    /// <summary>
    /// 외부 CLI 변환기로 입력을 같은 이름의 정식 IWPF 파일(예: book.html → book.iwpf)로 변환한 뒤
    /// 그 IWPF 파일을 메인 앱이 연다 (CLAUDE.md §3 — IWPF 가 정본).
    /// 사용자에게 변환 사실을 명시 안내하고, 같은 이름의 IWPF 가 이미 있으면 덮어쓰기/기존열기/취소 선택.
    /// </summary>
    private async Task OpenViaExternalConverterAsync(string sourcePath)
    {
        var converter = ExternalConverter.GetConverter(sourcePath);
        if (converter is null)
        {
            var ext = Path.GetExtension(sourcePath).TrimStart('.').ToLowerInvariant();
            MessageBox.Show(
                string.Format(SR.DlgUnsupportedFormat, ext),
                SR.DlgUnsupportedFormatTitle,
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        // 덮어쓰기 방지 모드: 안전한 번호 경로를 바로 사용 (덮어쓰기 대화 건너뜀).
        var iwpfBase = Path.ChangeExtension(sourcePath, ".iwpf");
        var iwpfPath = LanguageService.OverwriteProtection
            ? GetSafePath(iwpfBase)
            : iwpfBase;

        // 사용자 안내 — 변환 후 IWPF 를 연다는 사실을 분명히 알린다.
        var promptMsg = string.Format(SR.DlgConvertOnOpenPrompt,
            Path.GetFileName(sourcePath), Path.GetFileName(iwpfPath));
        var promptResult = MessageBox.Show(
            promptMsg,
            SR.DlgConvertOnOpenTitle,
            MessageBoxButton.OKCancel,
            MessageBoxImage.Information,
            MessageBoxResult.OK);
        if (promptResult != MessageBoxResult.OK) return;

        // 덮어쓰기 방지 비활성화 시: 같은 이름의 IWPF 가 이미 있으면 사용자 선택.
        if (!LanguageService.OverwriteProtection && File.Exists(iwpfPath))
        {
            var owMsg = string.Format(SR.DlgConvertOverwritePrompt, Path.GetFileName(iwpfPath));
            var ow = MessageBox.Show(
                owMsg,
                SR.DlgConvertOverwriteTitle,
                MessageBoxButton.YesNoCancel,
                MessageBoxImage.Warning,
                MessageBoxResult.No);
            if (ow == MessageBoxResult.Cancel) return;
            if (ow == MessageBoxResult.No)
            {
                // 기존 IWPF 를 그대로 연다 (변환 건너뜀).
                await OpenPathAsync(iwpfPath);
                return;
            }
            // Yes → 변환 진행해 덮어쓰기.
        }

        IsBusy        = true;
        BusyProgress  = 0;
        BusyMessage   = string.Format(SR.StatusBusyConvert, Path.GetFileName(sourcePath));
        var sourceName = Path.GetFileName(sourcePath);
        var iwpfName   = Path.GetFileName(iwpfPath);
        var reporter = new Progress<(int Percent, string Message)>(t =>
        {
            BusyProgress = t.Percent;
            BusyMessage  = string.Format(SR.StatusBusyConvertProgress, t.Percent, t.Message, sourceName);
        });
        try
        {
            try
            {
                await ExternalConverter.ConvertAsync(converter, sourcePath, iwpfPath, reporter);
            }
            catch (UnsupportedFormatVersionException ex)
            {
                MessageBox.Show(
                    string.Format(SR.DlgUnsupportedVersionPrompt, Path.GetFileName(sourcePath), ex.Message),
                    SR.DlgUnsupportedVersionTitle,
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            catch (Exception ex)
            {
                ReportError(SR.DlgOpenError, ex);
                return;
            }

            PolyDonkyument doc;
            try
            {
                BusyProgress = -1;  // IWPF 로드 단계는 indeterminate.
                BusyMessage  = string.Format(SR.StatusBusyOpen, iwpfName);
                doc = await Task.Run(() =>
                {
                    using var fs = File.OpenRead(iwpfPath);
                    return new IwpfReader().Read(fs);
                });
            }
            catch (Exception ex)
            {
                ReportError(SR.DlgOpenError, ex);
                return;
            }

            // 정본은 IWPF — CurrentFilePath 가 .iwpf 가 되어 다음 저장은 IWPF 로 직행.
            LoadDocument(doc, iwpfPath, password: null);
            StatusMessage = string.Format(SR.StatusConvertedAndOpened, sourceName, iwpfName);
        }
        finally
        {
            IsBusy       = false;
            BusyProgress = -1;
            BusyMessage  = string.Empty;
        }
    }

    /// <summary>
    /// 암호화된 IWPF 를 위한 비밀번호 프롬프트 + 재시도 루프. 사용자가 [취소] 를 누르면 null
    /// 을 반환한다 — 호출자는 결과가 null 이면 열기 작업 자체를 중단해야 한다.
    /// </summary>
    private PolyDonkyument? ReadEncryptedWithPrompt(IwpfReader reader, string path, out string? password)
    {
        string? errorMessage = null;
        while (true)
        {
            var prompt = new PasswordPromptWindow { Owner = Application.Current.MainWindow };
            if (errorMessage is not null) prompt.ShowError(errorMessage);
            if (prompt.ShowDialog() != true)
            {
                password = null;
                return null;
            }

            try
            {
                using var fs = File.OpenRead(path);
                var doc = reader.Read(fs, prompt.EnteredPassword);
                password = prompt.EnteredPassword;
                return doc;
            }
            catch (WrongIwpfPasswordException)
            {
                errorMessage = SR.PwdWrong;
            }
            catch (Exception ex)
            {
                ReportError(SR.DlgOpenError, ex);
                password = null;
                return null;
            }
        }
    }

    private static string BuildOpenStatusMessage(string path, PolyDonkyument doc)
    {
        var name = Path.GetFileName(path);
        return string.Format(SR.StatusOpenDone, name);
    }

    [RelayCommand]
    private async Task SaveAsync()
    {
        if (string.IsNullOrEmpty(CurrentFilePath))
        {
            await SaveAsAsync();
            return;
        }
        await SaveToAsync(CurrentFilePath);
    }

    [RelayCommand]
    private async Task SaveAsAsync()
    {
        var dlg = new SaveFileDialog
        {
            Filter = KnownFormats.SaveFilter,
            Title = SR.DlgFileSaveTitle,
            FileName = string.IsNullOrEmpty(CurrentFilePath)
                ? SR.DlgDefaultFileName
                : Path.GetFileNameWithoutExtension(CurrentFilePath),
        };
        if (dlg.ShowDialog() != true) return;

        // 외부 CLI 변환 포맷(DOCX/HWPX/HTML/XML)은 SaveToAsync → SaveViaExternalConverterAsync
        // 가 IWPF 정본을 만든 뒤 CLI 로 변환한다. 변환기 미등록 시(HWP/DOC) 그쪽에서 안내.
        await SaveToAsync(dlg.FileName);
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
    private void LicenseInfo()
    {
        var window = new LicenseInfoWindow
        {
            Owner = Application.Current.MainWindow,
        };
        window.ShowDialog();
    }

    [RelayCommand]
    private void UserGuide()
    {
        var window = new UserGuideWindow
        {
            Owner       = Application.Current.MainWindow,
            SelectedTab = UserGuideWindow.Tab.UserGuide,
        };
        window.ShowDialog();
    }

    [RelayCommand]
    private void IwpfFormatInfo()
    {
        var window = new UserGuideWindow
        {
            Owner       = Application.Current.MainWindow,
            SelectedTab = UserGuideWindow.Tab.IwpfFormat,
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

        // 워터마크 초기값 — 없으면 기본값 채워서 다이얼로그가 즉시 입력 가능한 상태가 되게.
        var wm = _document.Watermark ?? new WatermarkSettings { Enabled = false };

        // 편집 암호로 잠긴 상태이면 워터마크 데이터를 노출하지 않음.
        bool wmLocked = IsWriteProtected;

        var info = new DocumentInfoModel
        {
            FilePath       = path,
            Format         = format,
            DataSize       = DocumentMeasurement.FormatBytes(bytes),
            DocTitle       = string.IsNullOrWhiteSpace(meta.Title) ? none : meta.Title,
            HasBeenSaved   = !string.IsNullOrEmpty(CurrentFilePath),
            Author         = meta.Author ?? string.Empty,
            Editor         = meta.Editor ?? string.Empty,
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
            PasswordMode      = CurrentPasswordMode,
            WatermarkEnabled  = wmLocked ? false : wm.Enabled,
            WatermarkText     = wmLocked ? "" : wm.Text,
            WatermarkColor    = wmLocked ? "#FF808080" : wm.Color,
            WatermarkFontSize = wmLocked ? 48 : wm.FontSize,
            WatermarkRotation = wmLocked ? -45.0 : wm.Rotation,
            WatermarkOpacity  = wmLocked ? 0.3 : wm.Opacity,
            PrintWithWatermark = wmLocked ? true : wm.PrintWithWatermark,
            IsPrintable       = wmLocked ? true : _document.IsPrintable,
            IsWatermarkLocked = wmLocked,
        };

        // 잠금 해제 콜백 — 기존 VerifyWritePassword() 재사용.
        if (wmLocked)
        {
            info.UnlockWatermarkAction = () =>
            {
                if (!VerifyWritePassword()) return;
                info.WatermarkEnabled  = wm.Enabled;
                info.WatermarkText     = wm.Text;
                info.WatermarkColor    = wm.Color;
                info.WatermarkFontSize = wm.FontSize;
                info.WatermarkRotation = wm.Rotation;
                info.WatermarkOpacity  = wm.Opacity;
                info.PrintWithWatermark = wm.PrintWithWatermark;
                info.IsPrintable       = _document.IsPrintable;
                info.IsWatermarkLocked = false;
            };
        }

        var dlg = new DocumentInfoWindow(info) { Owner = Application.Current.MainWindow };
        if (dlg.ShowDialog() != true) return;

        ApplyDocumentInfoChanges(info);

        // 워터마크 변경 시 페이지 프레임 재구축 필요
        if (Application.Current.MainWindow is Views.MainWindow mainWindow)
            mainWindow.RebuildPageFrames();
    }

    /// <summary>
    /// 문서 정보 다이얼로그가 [확인] 으로 닫혔을 때 사용자가 편집한 항목을
    /// _document / _documentPassword 에 반영하고 변경이 있으면 dirty 로 표시한다.
    /// </summary>
    private void ApplyDocumentInfoChanges(DocumentInfoModel info)
    {
        var dirty = false;
        var meta = _document.Metadata;

        // ── 제목 ──
        var oldTitle = meta.Title;
        var newTitle = string.IsNullOrWhiteSpace(info.DocTitle) ? null : info.DocTitle.Trim();
        if (!string.Equals(oldTitle, newTitle, StringComparison.Ordinal))
        {
            meta.Title = newTitle;
            dirty = true;
        }

        // ── 작성자 ──
        var oldAuthor = meta.Author;
        var newAuthor = string.IsNullOrWhiteSpace(info.Author) ? null : info.Author.Trim();
        if (!string.Equals(oldAuthor, newAuthor, StringComparison.Ordinal))
        {
            meta.Author = newAuthor;
            // 작성자가 처음 지정되면 생성일도 함께 업데이트
            if (string.IsNullOrEmpty(oldAuthor) && newAuthor is not null)
                meta.Created = DateTimeOffset.UtcNow;
            dirty = true;
        }

        // ── 수정자 ──
        var newEditor = string.IsNullOrWhiteSpace(info.Editor) ? null : info.Editor.Trim();
        if (!string.Equals(meta.Editor, newEditor, StringComparison.Ordinal))
        {
            meta.Editor = newEditor;
            meta.Modified = DateTimeOffset.UtcNow;
            dirty = true;
        }

        // ── 언어 ──
        if (!string.Equals(meta.Language, info.Language, StringComparison.Ordinal))
        {
            meta.Language = info.Language;
            dirty = true;
        }

        // ── 비밀번호 보호 모드 ──
        if (info.PasswordChanged)
        {
            var readPwd  = string.IsNullOrEmpty(info.NewReadPassword)  ? null : info.NewReadPassword;
            var writePwd = string.IsNullOrEmpty(info.NewWritePassword) ? null : info.NewWritePassword;
            switch (info.PasswordMode)
            {
                case PasswordMode.None:
                    _documentPassword = null;
                    _writePassword    = null;
                    _writeLock        = null;
                    break;
                case PasswordMode.Read:
                    _documentPassword = readPwd;
                    _writePassword    = null;
                    _writeLock        = null;
                    break;
                case PasswordMode.Write:
                    _documentPassword = null;
                    _writePassword    = writePwd;
                    _writeLock        = writePwd is not null ? IwpfEncryption.CreateWriteLock(writePwd) : null;
                    break;
                case PasswordMode.Both:
                    _documentPassword = readPwd;
                    _writePassword    = writePwd;
                    _writeLock        = writePwd is not null ? IwpfEncryption.CreateWriteLock(writePwd) : null;
                    break;
            }
            RecomputeWriteProtection();
            dirty = true;
        }

        // ── 워터마크 + 인쇄 가능 여부 ──
        // 잠금 상태로 닫힌 경우 모델 필드는 빈 값/기본값이므로 _document 를 건드리지 않는다.
        if (!info.IsWatermarkLocked)
        {
            WatermarkSettings? newWm = info.WatermarkEnabled
                ? new WatermarkSettings
                {
                    Enabled  = true,
                    Text     = info.WatermarkText ?? "",
                    Color    = info.WatermarkColor ?? "#FF808080",
                    FontSize = Math.Max(1, info.WatermarkFontSize),
                    Rotation = info.WatermarkRotation,
                    Opacity  = Math.Clamp(info.WatermarkOpacity, 0.0, 1.0),
                    PrintWithWatermark = info.PrintWithWatermark,
                }
                : null;
            if (!WatermarkEquals(_document.Watermark, newWm))
            {
                _document.Watermark = newWm;
                dirty = true;
            }

            // ── 인쇄 가능 여부 ──
            if (_document.IsPrintable != info.IsPrintable)
            {
                _document.IsPrintable = info.IsPrintable;
                dirty = true;
            }
        }

        if (dirty) MarkDirty();
    }

    private static bool WatermarkEquals(WatermarkSettings? a, WatermarkSettings? b)
    {
        if (a is null && b is null) return true;
        if (a is null || b is null) return false;
        return a.Enabled == b.Enabled
            && a.Text == b.Text
            && a.Color == b.Color
            && a.FontSize == b.FontSize
            && a.Rotation.Equals(b.Rotation)
            && a.Opacity.Equals(b.Opacity)
            && a.PrintWithWatermark == b.PrintWithWatermark;
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
    private void OutlineStyle()
    {
        OutlineStyleRequested?.Invoke(this, EventArgs.Empty);
    }

    public event EventHandler? OutlineStyleRequested;

    /// <summary>
    /// 개요 서식 적용. live FlowDocument 를 먼저 _document 에 동기화해
    /// 저장 이후 편집한 이미지·글상자를 보존한 뒤 재빌드한다.
    /// </summary>
    public void ApplyOutlineStyles(PolyDonky.Core.OutlineStyleSet styleSet, Wpf.FlowDocument liveDoc)
    {
        _document = FlowDocumentParser.Parse(liveDoc, _document);
        _document.OutlineStyles = styleSet;
        RebuildFlowDocument();
    }

    /// <summary>미리보기·인쇄용 — 현재 live FlowDocument 를 Core 로 동기화한 스냅샷을 반환.</summary>
    public PolyDonkyument GetPreviewDocument()
        => FlowDocumentParser.Parse(FlowDocument, _document);

    // ── 도형 (Insert > Shape) ────────────────────────────────────────────
    // 메뉴 클릭 → ViewModel 이 이벤트 발생 → View 가 드래그 생성 모드 진입.
    // 모델 갱신은 View 가 AddShapeToCurrentSection 으로 위임.

    [RelayCommand]
    private void InsertShape(object? kindParam)
    {
        var kind = kindParam switch
        {
            "Line"           => PolyDonky.Core.ShapeKind.Line,
            "Polyline"       => PolyDonky.Core.ShapeKind.Polyline,
            "Polygon"        => PolyDonky.Core.ShapeKind.Polygon,
            "Spline"         => PolyDonky.Core.ShapeKind.Spline,
            "ClosedSpline"   => PolyDonky.Core.ShapeKind.ClosedSpline,
            "RoundedRect"    => PolyDonky.Core.ShapeKind.RoundedRect,
            "Ellipse"        => PolyDonky.Core.ShapeKind.Ellipse,
            "Triangle"       => PolyDonky.Core.ShapeKind.Triangle,
            "RegularPolygon" => PolyDonky.Core.ShapeKind.RegularPolygon,
            "Star"           => PolyDonky.Core.ShapeKind.Star,
            _                => PolyDonky.Core.ShapeKind.Rectangle,
        };
        InsertShapeRequested?.Invoke(this, kind);
    }

    public event EventHandler<PolyDonky.Core.ShapeKind>? InsertShapeRequested;

    /// <summary>드래그 생성 완료 후 View 가 호출 — 첫 섹션의 Blocks 에 추가하고 Dirty 표시.</summary>
    public void AddShapeToCurrentSection(PolyDonky.Core.ShapeObject shape)
    {
        var section = _document.Sections.FirstOrDefault();
        if (section is null) return;
        section.Blocks.Add(shape);
        MarkDirty();
        RefreshMemoryUsage();
    }

    /// <summary>도형 삭제 시 View 가 호출.</summary>
    public void RemoveShape(PolyDonky.Core.ShapeObject shape)
    {
        foreach (var section in _document.Sections)
        {
            if (section.Blocks.Remove(shape))
            {
                MarkDirty();
                RefreshMemoryUsage();
                return;
            }
        }
    }

    // ── 글상자 (Insert > TextBox) ────────────────────────────────────────
    // 메뉴 클릭 → ViewModel 이 이벤트 발생 → View 가 드래그 생성 모드 진입.
    // 글상자도 다른 부유 개체와 동일하게 Section.Blocks 에 추가됨 (IWPF 통합 모델).

    [RelayCommand]
    private void InsertTextBox(object? shapeParam)
    {
        var shape = shapeParam switch
        {
            "Speech"    => PolyDonky.Core.TextBoxShape.Speech,
            "Cloud"     => PolyDonky.Core.TextBoxShape.Cloud,
            "Spiky"     => PolyDonky.Core.TextBoxShape.Spiky,
            "Lightning" => PolyDonky.Core.TextBoxShape.Lightning,
            "Ellipse"   => PolyDonky.Core.TextBoxShape.Ellipse,
            "Pie"       => PolyDonky.Core.TextBoxShape.Pie,
            _           => PolyDonky.Core.TextBoxShape.Rectangle,
        };
        InsertTextBoxRequested?.Invoke(this, shape);
    }

    public event EventHandler<PolyDonky.Core.TextBoxShape>? InsertTextBoxRequested;

    /// <summary>드래그 생성 완료 후 View 가 호출 — 첫 섹션의 Blocks 에 글상자/도형/이미지/표를 추가.</summary>
    public void AddOverlayBlockToCurrentSection(PolyDonky.Core.Block block)
    {
        var section = _document.Sections.FirstOrDefault();
        if (section is null) return;
        section.Blocks.Add(block);
        MarkDirty();
        RefreshMemoryUsage();
    }

    /// <summary>오버레이 삭제 시 View 가 호출.</summary>
    public void RemoveOverlayBlock(PolyDonky.Core.Block block)
    {
        foreach (var section in _document.Sections)
        {
            if (section.Blocks.Remove(block))
            {
                MarkDirty();
                RefreshMemoryUsage();
                return;
            }
        }
    }

    /// <summary>드래그/리사이즈/본문 편집 변경 시 View 가 호출.</summary>
    public void NotifyOverlayChanged() => MarkDirty();

    /// <summary>
    /// _document 로부터 FlowDocument 를 재빌드한다. 호출 전 반드시 _document 가 최신 상태여야 한다.
    /// 편집 중 그림·글상자를 잃지 않으려면 먼저 live FlowDocument 를 Parse 해 _document 를 동기화할 것.
    /// </summary>
    private void RebuildFlowDocument()
    {
        _suppressDirty = true;
        FlowDocument = FlowDocumentBuilder.Build(_document);
        _suppressDirty = false;
        HasUnsavedChanges = true;
    }

    [RelayCommand]
    private void Close()
    {
        if (!ConfirmDiscardChanges()) return;
        LoadDocument(PolyDonkyument.Empty(), null, password: null);
        StatusMessage = SR.StatusDocClosed;
    }

    [RelayCommand]
    private void Exit()
    {
        if (!ConfirmDiscardChanges()) return;
        Application.Current.Shutdown();
    }

    private async Task SaveToAsync(string path)
    {
        if (IsBusy) return;  // 동시 저장 가드

        IsBusy      = true;
        BusyMessage = string.Format(SR.StatusBusySave, Path.GetFileName(path));
        try
        {
            // CLAUDE.md §3 — 외부 CLI 위탁 포맷은 임시 IWPF 만들고 CLI 가 최종 변환.
            if (KnownFormats.RequiresExternalConverter(path))
            {
                await SaveViaExternalConverterAsync(path);
            }
            else
            {
                await Task.Run(() => SaveToCore(path, showProgress: false));
            }
        }
        finally
        {
            IsBusy      = false;
            BusyMessage = string.Empty;
        }
    }

    private async Task SaveViaExternalConverterAsync(string targetPath)
    {
        var converter = ExternalConverter.GetConverter(targetPath);
        if (converter is null)
        {
            var ext = Path.GetExtension(targetPath).TrimStart('.').ToLowerInvariant();
            MessageBox.Show(string.Format(SR.DlgSaveConverterNeeded, ext),
                SR.DlgSaveConverterTitle, MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        // 같은 이름의 정식 .iwpf 와 외부 포맷을 함께 저장 (IWPF 정본 + 외부 export).
        var iwpfBase = Path.ChangeExtension(targetPath, ".iwpf");
        var iwpfPath = LanguageService.OverwriteProtection ? GetSafePath(iwpfBase) : iwpfBase;
        if (LanguageService.OverwriteProtection)
            targetPath = GetSafePath(targetPath);

        var promptMsg = string.Format(SR.DlgConvertOnSavePrompt,
            Path.GetFileName(iwpfPath), Path.GetFileName(targetPath));
        var promptResult = MessageBox.Show(
            promptMsg,
            SR.DlgConvertOnSaveTitle,
            MessageBoxButton.OKCancel,
            MessageBoxImage.Information,
            MessageBoxResult.OK);
        if (promptResult != MessageBoxResult.OK) return;

        var iwpfName   = Path.GetFileName(iwpfPath);
        var targetName = Path.GetFileName(targetPath);
        try
        {
            // 1) 현재 모델을 같은 이름의 IWPF 정본 파일로 저장 (UI 스레드 — FlowDocument 파싱).
            BusyProgress = -1;
            BusyMessage  = string.Format(SR.StatusBusySave, iwpfName);
            await Task.Run(() => SaveToCore(iwpfPath, showProgress: false));
            if (!File.Exists(iwpfPath))
            {
                ReportError(SR.DlgSaveError, new InvalidOperationException("IWPF 저장 실패"));
                return;
            }

            // 2) CLI 가 IWPF 정본을 외부 포맷으로 변환.
            BusyProgress = 0;
            BusyMessage  = string.Format(SR.StatusBusyConvert, iwpfName);
            var reporter = new Progress<(int Percent, string Message)>(t =>
            {
                BusyProgress = t.Percent;
                BusyMessage  = string.Format(SR.StatusBusyConvertProgress, t.Percent, t.Message, targetName);
            });
            try
            {
                await ExternalConverter.ConvertAsync(converter, iwpfPath, targetPath, reporter);
            }
            catch (Exception ex)
            {
                ReportError(SR.DlgSaveError, ex);
                return;
            }

            // 정본은 IWPF — CurrentFilePath 를 .iwpf 로 갱신해 이후 저장은 IWPF 로 직행.
            CurrentFilePath = iwpfPath;
            DocumentTitle   = iwpfName;
            HasUnsavedChanges = false;
            StatusMessage   = string.Format(SR.StatusSavedAndConverted, iwpfName, targetName);
        }
        catch (Exception ex)
        {
            ReportError(SR.DlgSaveError, ex);
        }
        finally
        {
            BusyProgress = -1;
        }
    }

    /// <summary>
    /// 덮어쓰기 방지: 파일이 이미 존재하면 이름 뒤에 -1, -2, ... 를 붙인 첫 번째 빈 경로를 반환.
    /// 파일이 없으면 그대로 반환한다.
    /// </summary>
    private static string GetSafePath(string path)
    {
        if (!File.Exists(path)) return path;

        var dir  = Path.GetDirectoryName(path) ?? string.Empty;
        var name = Path.GetFileNameWithoutExtension(path);
        var ext  = Path.GetExtension(path);
        for (int i = 1; i <= 9999; i++)
        {
            var candidate = Path.Combine(dir, $"{name}-{i}{ext}");
            if (!File.Exists(candidate)) return candidate;
        }
        return path;  // fallback (사실상 발생 불가)
    }

    /// <summary>
    /// 동기 저장 — 비동기 경로(<see cref="SaveToAsync"/>)와 ConfirmDiscardChanges 같은
    /// 동기 호출자가 공유한다. UI 스레드에서 직접 호출돼도 안전하지만 진행 막대는 보이지 않는다.
    /// </summary>
    private void SaveToCore(string path, bool showProgress)
    {
        var writer = KnownFormats.PickWriter(path);
        if (writer is null)
        {
            // 백그라운드 스레드에서도 MessageBox 는 자체 STA 메시지 펌프로 동작 — 안전.
            MessageBox.Show(SR.DlgSaveFail, SR.DlgSaveFailTitle,
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (writer is IwpfWriter && _writeLock is not null)
        {
            if (!VerifyWritePassword()) return;
        }

        var isFirstSave = string.IsNullOrEmpty(CurrentFilePath);
        var now = DateTimeOffset.UtcNow;

        try
        {
            // FlowDocument → 모델 변환은 UI 스레드 종속 — Dispatcher 로 마샬링.
            PolyDonkyument rebuilt = null!;
            Application.Current.Dispatcher.Invoke(() =>
            {
                rebuilt = LiveDocumentProvider?.Invoke()
                       ?? FlowDocumentParser.Parse(FlowDocument, _document);
            });

            if (isFirstSave) { rebuilt.Metadata.Created  = now; rebuilt.Metadata.Modified = now; }
            else             { rebuilt.Metadata.Modified = now; }

            if (string.IsNullOrWhiteSpace(rebuilt.Metadata.Title))
                rebuilt.Metadata.Title = ExtractFirstLineTitle(rebuilt);

            var actualWriter = writer;
            if (writer is IwpfWriter)
            {
                var mode = CurrentPasswordMode;
                if (mode != PasswordMode.None) actualWriter = BuildIwpfWriter(mode);
            }

            using (var fs = File.Create(path)) actualWriter.Write(rebuilt, fs);

            // UI 스레드에서 상태 갱신.
            Application.Current.Dispatcher.Invoke(() =>
            {
                _document = rebuilt;
                CurrentFilePath = path;
                DocumentTitle = Path.GetFileName(path);
                HasUnsavedChanges = false;
                StatusMessage = string.Format(SR.StatusSaveDone, Path.GetFileName(path));
            });
        }
        catch (Exception ex)
        {
            Application.Current.Dispatcher.Invoke(() => ReportError(SR.DlgSaveError, ex));
        }
    }

    private string? ExtractFirstLineTitle(PolyDonkyument doc)
    {
        var firstSection = doc.Sections.FirstOrDefault();
        if (firstSection is null || firstSection.Blocks.Count == 0) return null;

        var firstBlock = firstSection.Blocks.FirstOrDefault();
        string? text = firstBlock switch
        {
            Paragraph p => p.GetPlainText(),
            TextBoxObject tb => tb.GetPlainText(),
            Table t when t.Rows.Count > 0 && t.Rows[0].Cells.Count > 0
                => ExtractTableCellText(t.Rows[0].Cells[0]),
            _ => null
        };

        if (string.IsNullOrWhiteSpace(text)) return null;

        var words = text.Split(new[] { ' ', '\t', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
        var limited = string.Join(" ", words.Take(6));
        return string.IsNullOrWhiteSpace(limited) ? null : limited;
    }

    private string ExtractTableCellText(TableCell cell)
    {
        var sb = new System.Text.StringBuilder();
        foreach (var block in cell.Blocks)
        {
            if (block is Paragraph p)
                sb.Append(p.GetPlainText());
        }
        return sb.ToString();
    }

    /// <summary>
    /// 쓰기 보호 비밀번호를 프롬프트로 입력받아 _writeLock 과 대조한다.
    /// 이 세션에서 이미 검증됐으면 (_writePassword != null) 즉시 true 반환.
    /// 사용자가 [취소] 를 누르면 false 반환.
    /// </summary>
    private bool VerifyWritePassword()
    {
        if (_writePassword is not null) return true; // 세션 내 이미 검증됨

        string? errorMessage = null;
        while (true)
        {
            var prompt = new PasswordPromptWindow { Owner = Application.Current.MainWindow };
            prompt.SetMessage(SR.PwdWritePromptMessage);
            if (errorMessage is not null) prompt.ShowError(errorMessage);
            if (prompt.ShowDialog() != true) return false;

            if (IwpfEncryption.VerifyWriteLock(prompt.EnteredPassword, _writeLock!))
            {
                _writePassword = prompt.EnteredPassword; // 이후 편집·저장 시 재사용
                RecomputeWriteProtection();              // IsWriteProtected → false
                return true;
            }

            errorMessage = SR.PwdWrong;
        }
    }

    /// <summary>
    /// 쓰기 보호 상태에서 편집 시도가 감지됐을 때 호출. 비밀번호 프롬프트를 띄우고
    /// 검증 성공 시 IsWriteProtected 를 false 로 풀어 이후 편집을 허용한다.
    /// 이미 잠금 해제 상태이면 즉시 반환.
    /// </summary>
    public void TryUnlockForEditing()
    {
        if (!IsWriteProtected) return;
        if (VerifyWritePassword())
            StatusMessage = SR.StatusWriteUnlocked;
    }

    private IwpfWriter BuildIwpfWriter(PasswordMode mode)
        => new IwpfWriter
        {
            PasswordMode  = mode,
            Password      = mode != PasswordMode.Write ? _documentPassword : null,
            WritePassword = mode != PasswordMode.Read  ? _writePassword    : null,
        };

    // ── 확대/축소 ────────────────────────────────────────────────────────────
    private double _zoomPercent = 100;

    /// <summary>편집창 배율 (10–500 %). ScaleTransform 을 통해 PaperStackPanel 에 LayoutTransform 으로 적용.</summary>
    public double ZoomPercent
    {
        get => _zoomPercent;
        set
        {
            var clamped = Math.Clamp(Math.Round(value), 10, 500);
            if (SetProperty(ref _zoomPercent, clamped))
                OnPropertyChanged(nameof(ZoomScale));
        }
    }

    /// <summary>ZoomPercent / 100. ScaleTransform.ScaleX/Y 에 바인딩.</summary>
    public double ZoomScale => _zoomPercent / 100.0;

    [RelayCommand]
    private void ZoomIn()  => ZoomPercent += 5;

    [RelayCommand]
    private void ZoomOut() => ZoomPercent -= 5;

    private bool ConfirmDiscardChanges()
    {
        if (!HasUnsavedChanges) return true;
        var result = MessageBox.Show(
            SR.DlgUnsavedPrompt,
            "PolyDonky",
            MessageBoxButton.YesNoCancel,
            MessageBoxImage.Question);
        if (result == MessageBoxResult.Cancel) return false;
        if (result == MessageBoxResult.Yes)
        {
            // 동기 컨텍스트에서 호출되므로 동기 코어를 직접 사용 — 진행 막대 없음.
            if (string.IsNullOrEmpty(CurrentFilePath))
            {
                // 신규 문서면 다이얼로그가 필요 — 비동기 SaveAs 를 fire-and-forget 으로 실행 불가
                // (호출자가 결과를 기다림). 동기 대화상자 후 동기 저장.
                var dlg = new SaveFileDialog
                {
                    Filter   = KnownFormats.SaveFilter,
                    Title    = SR.DlgFileSaveTitle,
                    FileName = SR.DlgDefaultFileName,
                };
                if (dlg.ShowDialog() != true) return false;
                SaveToCore(dlg.FileName, showProgress: false);
            }
            else
            {
                SaveToCore(CurrentFilePath, showProgress: false);
            }
            return !HasUnsavedChanges;
        }
        return true;
    }

    private void ReportError(string headline, Exception ex)
    {
        StatusMessage = $"오류: {ex.Message}";
        MessageBox.Show($"{headline}\n\n{ex.GetType().Name}: {ex.Message}",
            "PolyDonky", MessageBoxButton.OK, MessageBoxImage.Error);
    }
}

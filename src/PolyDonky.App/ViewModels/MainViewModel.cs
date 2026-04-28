using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.Json;
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

    [ObservableProperty]
    private bool _isSpellCheckEnabled = true;

    [RelayCommand]
    private void ToggleSpellCheck() => IsSpellCheckEnabled = !IsSpellCheckEnabled;

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
    }

    [RelayCommand]
    private void New()
    {
        if (!ConfirmDiscardChanges()) return;
        LoadDocument(PolyDonkyument.Empty(), null, password: null);
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

        PolyDonkyument? doc;
        string? usedPassword = null;
        try
        {
            using var fs = File.OpenRead(path);
            doc = reader.Read(fs);
        }
        catch (EncryptedIwpfException) when (reader is IwpfReader iwpf)
        {
            // 암호화된 IWPF — 비밀번호 입력 + 재시도 루프.
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
                : Path.GetFileNameWithoutExtension(CurrentFilePath),
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
    private void LicenseInfo()
    {
        var window = new LicenseInfoWindow
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

        // 워터마크 초기값 — 없으면 기본값 채워서 다이얼로그가 즉시 입력 가능한 상태가 되게.
        var wm = _document.Watermark ?? new WatermarkSettings { Enabled = false };

        var info = new DocumentInfoModel
        {
            FilePath       = path,
            Format         = format,
            DataSize       = DocumentMeasurement.FormatBytes(bytes),
            DocTitle       = string.IsNullOrWhiteSpace(meta.Title) ? none : meta.Title,
            Author         = meta.Author ?? string.Empty,
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
            WatermarkEnabled  = wm.Enabled,
            WatermarkText     = wm.Text,
            WatermarkColor    = wm.Color,
            WatermarkFontSize = wm.FontSize,
            WatermarkRotation = wm.Rotation,
            WatermarkOpacity  = wm.Opacity,
        };

        var dlg = new DocumentInfoWindow(info) { Owner = Application.Current.MainWindow };
        if (dlg.ShowDialog() != true) return;

        ApplyDocumentInfoChanges(info);
    }

    /// <summary>
    /// 문서 정보 다이얼로그가 [확인] 으로 닫혔을 때 사용자가 편집한 항목을
    /// _document / _documentPassword 에 반영하고 변경이 있으면 dirty 로 표시한다.
    /// </summary>
    private void ApplyDocumentInfoChanges(DocumentInfoModel info)
    {
        var dirty = false;
        var meta = _document.Metadata;

        // ── 작성자 ──
        var newAuthor = string.IsNullOrWhiteSpace(info.Author) ? null : info.Author.Trim();
        if (!string.Equals(meta.Author, newAuthor, StringComparison.Ordinal))
        {
            meta.Author = newAuthor;
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

        // ── 워터마크 ──
        WatermarkSettings? newWm = info.WatermarkEnabled
            ? new WatermarkSettings
            {
                Enabled  = true,
                Text     = info.WatermarkText ?? "",
                Color    = info.WatermarkColor ?? "#FF808080",
                FontSize = Math.Max(1, info.WatermarkFontSize),
                Rotation = info.WatermarkRotation,
                Opacity  = Math.Clamp(info.WatermarkOpacity, 0.0, 1.0),
            }
            : null;
        if (!WatermarkEquals(_document.Watermark, newWm))
        {
            _document.Watermark = newWm;
            dirty = true;
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
            && a.Opacity.Equals(b.Opacity);
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
    // 모델 갱신은 View 가 AddFloatingObjectToCurrentSection 으로 위임.

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

    /// <summary>드래그 생성 완료 후 View 가 호출 — 첫 섹션의 FloatingObjects 에 추가하고 Dirty 표시.</summary>
    public void AddFloatingObjectToCurrentSection(PolyDonky.Core.FloatingObject obj)
    {
        var section = _document.Sections.FirstOrDefault();
        if (section is null) return;
        section.FloatingObjects.Add(obj);
        MarkDirty();
        RefreshMemoryUsage();
    }

    /// <summary>오버레이 삭제 시 View 가 호출.</summary>
    public void RemoveFloatingObject(PolyDonky.Core.FloatingObject obj)
    {
        foreach (var section in _document.Sections)
        {
            if (section.FloatingObjects.Remove(obj))
            {
                MarkDirty();
                RefreshMemoryUsage();
                return;
            }
        }
    }

    /// <summary>드래그/리사이즈/본문 편집 변경 시 View 가 호출.</summary>
    public void NotifyFloatingObjectChanged() => MarkDirty();

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

    private void SaveTo(string path)
    {
        var writer = KnownFormats.PickWriter(path);
        if (writer is null)
        {
            MessageBox.Show(SR.DlgSaveFail, SR.DlgSaveFailTitle,
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        // 쓰기 보호가 설정된 경우 저장 전에 비밀번호를 검증한다.
        if (writer is IwpfWriter && _writeLock is not null)
        {
            if (!VerifyWritePassword()) return; // 사용자 취소 또는 불일치 반복
        }

        // CurrentFilePath 가 비어 있으면 디스크에 한 번도 안 쓰여진 신규 문서 — 작성일 갱신.
        var isFirstSave = string.IsNullOrEmpty(CurrentFilePath);
        var now = DateTimeOffset.UtcNow;

        try
        {
            var rebuilt = FlowDocumentParser.Parse(FlowDocument, _document);

            // 자동 일자 갱신: 첫 저장이면 Created+Modified 둘 다, 아니면 Modified 만.
            if (isFirstSave)
            {
                rebuilt.Metadata.Created  = now;
                rebuilt.Metadata.Modified = now;
            }
            else
            {
                rebuilt.Metadata.Modified = now;
            }

            // IWPF writer 에 현재 보호 모드와 비밀번호를 지정한다.
            var actualWriter = writer;
            if (writer is IwpfWriter)
            {
                var mode = CurrentPasswordMode;
                if (mode != PasswordMode.None)
                    actualWriter = BuildIwpfWriter(mode);
            }

            using var fs = File.Create(path);
            actualWriter.Write(rebuilt, fs);

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
            Save();
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

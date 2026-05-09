using System;
using System.Globalization;
using System.IO;
using System.Text.Json;
using System.Threading;
using PolyDonky.App.Properties;

namespace PolyDonky.App.Services;

public static class LanguageService
{
    public enum Language { Korean, English }

    private static readonly string SettingsPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "Handtech", "PolyDonky", "settings.json");

    private static Language _current = Language.Korean;
    public static Language Current => _current;

    public static bool ShowTypesettingMarks { get; private set; }

    /// <summary>덮어쓰기 방지: 활성화 시 변환 파일 이름에 자동으로 -번호를 붙인다.</summary>
    public static bool OverwriteProtection { get; private set; }

    public static event EventHandler? LanguageChanged;

    /// <summary>앱 시작 시 저장된 언어 설정을 불러와 적용한다.</summary>
    public static void LoadAndApply()
    {
        var lang = Language.Korean;
        try
        {
            if (File.Exists(SettingsPath))
            {
                var json = File.ReadAllText(SettingsPath);
                var data = JsonSerializer.Deserialize<SettingsData>(json);
                if (data?.Language == "en-US") lang = Language.English;
                ShowTypesettingMarks = data?.ShowTypesettingMarks ?? false;
                OverwriteProtection  = data?.OverwriteProtection  ?? false;

                // 저장된 테마 복원 (없으면 Light 유지)
                if (data?.Theme is { } savedTheme &&
                    Enum.TryParse<ThemeService.Theme>(savedTheme, out var t))
                {
                    ThemeService.Apply(t);
                }
            }
        }
        catch { /* 손상된 설정 파일은 무시 */ }

        ApplyCore(lang, save: false);
    }

    public static void Apply(Language lang)
    {
        if (_current == lang) return;
        ApplyCore(lang, save: true);
    }

    private static void ApplyCore(Language lang, bool save)
    {
        _current = lang;
        var ci = lang == Language.English
            ? new CultureInfo("en-US")
            : new CultureInfo("ko-KR");

        // Resources.Culture 를 직접 지정하면 ResourceManager 가 모든 호출에서 이 culture 를 사용한다.
        Resources.Culture = ci;
        CultureInfo.DefaultThreadCurrentUICulture = ci;
        CultureInfo.DefaultThreadCurrentCulture = ci;
        Thread.CurrentThread.CurrentUICulture = ci;
        Thread.CurrentThread.CurrentCulture = ci;

        LocalizedStrings.Instance.Refresh();
        LanguageChanged?.Invoke(null, EventArgs.Empty);

        if (save) Save(lang);
    }

    public static void SetShowTypesettingMarks(bool value)
    {
        ShowTypesettingMarks = value;
        Save(_current);
    }

    /// <summary>테마 변경 시 설정 파일에 저장한다 (SettingsWindow 에서 호출).</summary>
    public static void SaveTheme(ThemeService.Theme theme)
    {
        // ThemeService.Apply 는 이미 호출됐으므로 저장만 수행.
        Save(_current);
    }

    /// <summary>덮어쓰기 방지 설정 변경 시 저장한다.</summary>
    public static void SetOverwriteProtection(bool value)
    {
        OverwriteProtection = value;
        Save(_current);
    }

    private static void Save(Language lang)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(SettingsPath)!);
            var json = JsonSerializer.Serialize(new SettingsData
            {
                Language             = lang == Language.English ? "en-US" : "ko-KR",
                ShowTypesettingMarks = ShowTypesettingMarks,
                Theme                = ThemeService.Current.ToString(),
                OverwriteProtection  = OverwriteProtection,
            });
            File.WriteAllText(SettingsPath, json);
        }
        catch { /* 저장 실패 무시 */ }
    }

    private sealed record SettingsData
    {
        public string? Language             { get; init; }
        public bool?   ShowTypesettingMarks { get; init; }
        public string? Theme                { get; init; }
        public bool?   OverwriteProtection  { get; init; }
    }
}

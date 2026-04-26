using System;
using System.Globalization;
using System.IO;
using System.Text.Json;
using System.Threading;
using PolyDoc.App.Properties;

namespace PolyDoc.App.Services;

public static class LanguageService
{
    public enum Language { Korean, English }

    private static readonly string SettingsPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "Handtech", "PolyDoc", "settings.json");

    private static Language _current = Language.Korean;
    public static Language Current => _current;

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

    private static void Save(Language lang)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(SettingsPath)!);
            var json = JsonSerializer.Serialize(new SettingsData
            {
                Language = lang == Language.English ? "en-US" : "ko-KR",
            });
            File.WriteAllText(SettingsPath, json);
        }
        catch { /* 저장 실패 무시 */ }
    }

    private sealed record SettingsData { public string? Language { get; init; } }
}

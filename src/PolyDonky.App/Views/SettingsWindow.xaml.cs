using System.Windows;
using PolyDonky.App.Services;

namespace PolyDonky.App.Views;

public partial class SettingsWindow : Window
{
    public SettingsWindow()
    {
        InitializeComponent();

        // 현재 테마 반영
        switch (ThemeService.Current)
        {
            case ThemeService.Theme.Dark: ThemeDark.IsChecked  = true; break;
            case ThemeService.Theme.Soft: ThemeSoft.IsChecked  = true; break;
            default:                      ThemeLight.IsChecked = true; break;
        }

        // 현재 언어 반영
        if (LanguageService.Current == LanguageService.Language.English)
            LangEnglish.IsChecked = true;
        else
            LangKorean.IsChecked = true;

        // 덮어쓰기 방지 반영
        OverwriteProtectionCheck.IsChecked = LanguageService.OverwriteProtection;
    }

    private void OnThemeChecked(object sender, RoutedEventArgs e)
    {
        if (ThemeDark is null || ThemeSoft is null) return;

        var theme = ThemeDark.IsChecked == true ? ThemeService.Theme.Dark
                  : ThemeSoft.IsChecked == true ? ThemeService.Theme.Soft
                  : ThemeService.Theme.Light;
        ThemeService.Apply(theme);
        LanguageService.SaveTheme(theme);
    }

    private void OnLanguageChecked(object sender, RoutedEventArgs e)
    {
        if (LangEnglish is null) return;

        var lang = LangEnglish.IsChecked == true
            ? LanguageService.Language.English
            : LanguageService.Language.Korean;
        LanguageService.Apply(lang);
    }

    private void OnOverwriteProtectionChanged(object sender, RoutedEventArgs e)
    {
        if (OverwriteProtectionCheck is null) return;
        LanguageService.SetOverwriteProtection(OverwriteProtectionCheck.IsChecked == true);
    }

    private void OnClose(object sender, RoutedEventArgs e) => Close();
}

using System.Windows;
using PolyDoc.App.Services;

namespace PolyDoc.App.Views;

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
    }

    private void OnThemeChecked(object sender, RoutedEventArgs e)
    {
        if (ThemeDark is null || ThemeSoft is null) return;

        if (ThemeDark.IsChecked == true)       ThemeService.Apply(ThemeService.Theme.Dark);
        else if (ThemeSoft.IsChecked == true)  ThemeService.Apply(ThemeService.Theme.Soft);
        else                                   ThemeService.Apply(ThemeService.Theme.Light);
    }

    private void OnLanguageChecked(object sender, RoutedEventArgs e)
    {
        if (LangEnglish is null) return;

        var lang = LangEnglish.IsChecked == true
            ? LanguageService.Language.English
            : LanguageService.Language.Korean;
        LanguageService.Apply(lang);
    }

    private void OnClose(object sender, RoutedEventArgs e) => Close();
}

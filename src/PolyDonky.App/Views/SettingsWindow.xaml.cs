using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Navigation;
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

        // LibreOffice 경로 반영
        UpdateLibreOfficePathDisplay();
    }

    private void UpdateLibreOfficePathDisplay()
    {
        if (LibreOfficePathBox is null) return;
        if (!string.IsNullOrEmpty(LanguageService.LibreOfficePath))
            LibreOfficePathBox.Text = LanguageService.LibreOfficePath;
        else
            LibreOfficePathBox.Text = LocalizedStrings.Get("SettingsLibreOfficeNotDetected");
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

    private void OnDetectLibreOffice(object sender, RoutedEventArgs e)
    {
        var path = LibreOfficeLocator.DetectLibreOfficePath();
        if (path != null)
        {
            try
            {
                LanguageService.SetLibreOfficePath(path);
                UpdateLibreOfficePathDisplay();
                MessageBox.Show(
                    LocalizedStrings.Get("SettingsLibreOfficeDetected"),
                    LocalizedStrings.Get("SettingsTitle"),
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"{LocalizedStrings.Get("SettingsLibreOfficeError")}\n{ex.Message}",
                    LocalizedStrings.Get("SettingsTitle"),
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }
        else
        {
            MessageBox.Show(
                LocalizedStrings.Get("SettingsLibreOfficeNotFoundMsg"),
                LocalizedStrings.Get("SettingsTitle"),
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }
    }

    private void OnLibreOfficeDownloadClick(object sender, RequestNavigateEventArgs e)
    {
        try
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri) { UseShellExecute = true });
            e.Handled = true;
        }
        catch { /* 브라우저 실행 실패 무시 */ }
    }

    private void OnClose(object sender, RoutedEventArgs e) => Close();
}

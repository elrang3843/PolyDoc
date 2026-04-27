using System;
using System.Windows;

namespace PolyDonky.App.Services;

/// <summary>런타임 테마 전환. App.xaml 의 첫 번째 MergedDictionary 가 테마 사전이라는 규약에 의존.</summary>
public static class ThemeService
{
    public enum Theme { Light, Dark, Soft }

    private static Theme _current = Theme.Light;
    public static Theme Current => _current;

    public static event EventHandler? ThemeChanged;

    public static void Apply(Theme theme)
    {
        if (_current == theme) return;
        _current = theme;

        var uri = theme switch
        {
            Theme.Dark => new Uri("pack://application:,,,/Themes/Dark.xaml"),
            Theme.Soft => new Uri("pack://application:,,,/Themes/Soft.xaml"),
            _          => new Uri("pack://application:,,,/Themes/Light.xaml"),
        };

        var merged = Application.Current.Resources.MergedDictionaries;
        // 규약: index 0 = 테마 사전.
        if (merged.Count > 0)
        {
            merged[0] = new ResourceDictionary { Source = uri };
        }
        ThemeChanged?.Invoke(null, EventArgs.Empty);
    }
}

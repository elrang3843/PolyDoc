using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Windows;

namespace PolyDonky.App.Views;

public partial class LicenseInfoWindow : Window
{
    public LicenseInfoWindow()
    {
        InitializeComponent();
        Loaded += OnLoaded;
    }

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        PolyDonkyLicenseText.Text = ReadBundledText("LICENSE") ?? ReadRepoFile("LICENSE");
        ThirdPartyText.Text     = ReadBundledText("THIRD_PARTY_NOTICES.md")
                               ?? ReadRepoFile("THIRD_PARTY_NOTICES.md");

        ShippedDepsGrid.ItemsSource = new List<DepEntry>
        {
            new(".NET 10 Runtime",             "10.x",  "MIT",         ".NET Foundation and Contributors"),
            new("CommunityToolkit.Mvvm",       "8.4.0", "MIT",         ".NET Foundation and Contributors"),
            new("DocumentFormat.OpenXml",      "3.5.1", "MIT",         "Microsoft Corporation and Contributors"),
            new("Markdig",                     "0.42.0","BSD-2-Clause","Alexandre Mutel"),
        };

        TestDepsGrid.ItemsSource = new List<DepEntry>
        {
            new("xunit",                       "2.9.3",  "Apache-2.0", ".NET Foundation and Contributors"),
            new("xunit.runner.visualstudio",   "3.1.4",  "Apache-2.0", ".NET Foundation and Contributors"),
            new("Microsoft.NET.Test.Sdk",      "17.14.1","MIT",         "Microsoft Corporation"),
            new("coverlet.collector",          "6.0.4",  "MIT",         "tonerdo"),
        };
    }

    private static string? ReadBundledText(string resourceName)
    {
        // Embedded resource path: PolyDonky.App.<name>
        var asm = Assembly.GetExecutingAssembly();
        var name = $"PolyDonky.App.{resourceName.Replace('/', '.')}";
        using var stream = asm.GetManifestResourceStream(name);
        if (stream is null) return null;
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }

    private static string ReadRepoFile(string relativeName)
    {
        // Walk up from the exe location to find the repo root file (dev fallback).
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null)
        {
            var candidate = Path.Combine(dir.FullName, relativeName);
            if (File.Exists(candidate))
                return File.ReadAllText(candidate);
            dir = dir.Parent;
        }
        return $"({relativeName} not found)";
    }

    private void OnCloseClick(object sender, RoutedEventArgs e) => Close();

    private sealed record DepEntry(string Name, string Version, string License, string Copyright);
}

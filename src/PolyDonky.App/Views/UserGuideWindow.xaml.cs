using System;
using System.IO;
using System.Reflection;
using System.Windows;

namespace PolyDonky.App.Views;

/// <summary>
/// 도움말 → "사용 방법" / "IWPF 포맷" 다이얼로그.
/// 임베드된 USER_GUIDE.md 와 IWPF.md 를 그대로 표시한다.
/// 호출자가 <see cref="SelectedTab"/> 으로 초기 탭을 지정할 수 있다.
/// </summary>
public partial class UserGuideWindow : Window
{
    public enum Tab { UserGuide = 0, IwpfFormat = 1 }

    /// <summary>창 오픈 시 선택될 탭. 기본 = 사용 방법.</summary>
    public Tab SelectedTab { get; init; } = Tab.UserGuide;

    public UserGuideWindow()
    {
        InitializeComponent();
        Loaded += OnLoaded;
    }

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        UserGuideText.Text  = ReadBundledText("USER_GUIDE.md") ?? ReadRepoFile("USER_GUIDE.md");
        IwpfFormatText.Text = ReadBundledText("IWPF.md")       ?? ReadRepoFile("IWPF.md");
        MainTabs.SelectedIndex = (int)SelectedTab;
    }

    private static string? ReadBundledText(string resourceName)
    {
        var asm = Assembly.GetExecutingAssembly();
        var name = $"PolyDonky.App.{resourceName.Replace('/', '.')}";
        using var stream = asm.GetManifestResourceStream(name);
        if (stream is null) return null;
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }

    private static string ReadRepoFile(string relativeName)
    {
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
}

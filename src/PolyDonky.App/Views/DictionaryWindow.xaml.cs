using System;
using System.Net;
using System.Windows;
using System.Windows.Input;
using SR = PolyDonky.App.Properties.Resources;

namespace PolyDonky.App.Views;

public partial class DictionaryWindow : Window
{
    private record DictSite(string Label, string HomeUrl, string SearchUrl);

    private static readonly DictSite[] Sites =
    [
        new("네이버 국어사전",  "https://ko.dict.naver.com/",  "https://ko.dict.naver.com/#/search?query={0}"),
        new("네이버 영한사전",  "https://endic.naver.com/",    "https://endic.naver.com/search.nhn?sLn=kr&searchOption=all&query={0}"),
        new("다음 사전",        "https://dic.daum.net/",       "https://dic.daum.net/search.do?q={0}"),
        new("표준국어대사전",   "https://stdict.korean.go.kr/","https://stdict.korean.go.kr/search/searchView.do?pageSize=10&searchKeyword={0}"),
    ];

    private bool _webViewReady;

    public DictionaryWindow(string? initialQuery = null)
    {
        InitializeComponent();

        foreach (var site in Sites)
            SiteCombo.Items.Add(site.Label);
        SiteCombo.SelectedIndex = 0;

        SearchBox.Text = initialQuery ?? string.Empty;

        Loaded += async (_, _) => await InitWebViewAsync(initialQuery);
    }

    private async System.Threading.Tasks.Task InitWebViewAsync(string? initialQuery)
    {
        try
        {
            await WebView.EnsureCoreWebView2Async();
            _webViewReady = true;

            if (!string.IsNullOrWhiteSpace(initialQuery))
                NavigateToSearch(initialQuery);
            else
                NavigateHome();
        }
        catch (Exception ex)
        {
            ErrorOverlay.Visibility = Visibility.Visible;
            ErrorText.Text = SR.DictWindowErrNoRuntime + $"\n\n({ex.Message})";
        }
    }

    private DictSite SelectedSite => Sites[Math.Max(0, SiteCombo.SelectedIndex)];

    private void NavigateHome()
    {
        if (_webViewReady)
            WebView.Source = new Uri(SelectedSite.HomeUrl);
    }

    private void NavigateToSearch(string query)
    {
        if (!_webViewReady) return;
        var url = string.Format(SelectedSite.SearchUrl, WebUtility.UrlEncode(query));
        WebView.Source = new Uri(url);
    }

    /// <summary>MainWindow에서 선택 텍스트를 넘겨 검색 실행.</summary>
    public void SearchFor(string query)
    {
        SearchBox.Text = query;
        NavigateToSearch(query);
    }

    private void OnSearchClick(object sender, RoutedEventArgs e)
    {
        var q = SearchBox.Text.Trim();
        if (string.IsNullOrEmpty(q))
            NavigateHome();
        else
            NavigateToSearch(q);
    }

    private void OnSearchBoxKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            OnSearchClick(sender, e);
            e.Handled = true;
        }
    }

    private void OnSiteChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        if (!_webViewReady) return;
        var q = SearchBox.Text.Trim();
        if (string.IsNullOrEmpty(q))
            NavigateHome();
        else
            NavigateToSearch(q);
    }

    /// <summary>MainWindow 종료 시 WebView2 포함 실제 닫기.</summary>
    public void ForceClose()
    {
        _forceClose = true;
        Close();
    }

    private bool _forceClose;

    private void OnWindowClosing(object sender, System.ComponentModel.CancelEventArgs e)
    {
        if (_forceClose) return;
        // 실제 닫지 않고 숨김 처리 — 재열기 시 재사용
        e.Cancel = true;
        Hide();
    }
}

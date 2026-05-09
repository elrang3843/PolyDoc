using System;
using System.Windows;
using System.Windows.Controls;

namespace PolyDonky.App.Views;

/// <summary>
/// 하이퍼링크 삽입/편집 다이얼로그.
/// <list type="bullet">
/// <item>삽입 모드: existingUrl 없이 생성. ResultUrl + ResultDisplayText 로 결과 읽기.</item>
/// <item>편집 모드: existingUrl 있이 생성 (기존 URL 미리채움). "링크 제거" 버튼도 표시.</item>
/// <item>제거: ResultUrl == null — 호출자가 기존 하이퍼링크를 일반 텍스트로 교체할 것.</item>
/// </list>
/// </summary>
public partial class HyperlinkDialog : Window
{
    /// <summary>확인 후 삽입/갱신할 URL. null 이면 링크 제거를 의미.</summary>
    public string? ResultUrl { get; private set; }

    /// <summary>표시 텍스트. null 이면 URL 그대로 사용.</summary>
    public string? ResultDisplayText { get; private set; }

    /// <summary>삽입/편집 모드로 다이얼로그를 연다.</summary>
    /// <param name="existingUrl">기존 링크 URL (없으면 빈 문자열).</param>
    /// <param name="selectedText">현재 선택 텍스트 — 표시 텍스트 필드 초기값.</param>
    /// <param name="canRemove">true 이면 "링크 제거" 버튼을 표시한다.</param>
    public HyperlinkDialog(string existingUrl = "", string selectedText = "", bool canRemove = false)
    {
        InitializeComponent();

        TxtUrl.Text         = existingUrl;
        TxtDisplayText.Text = selectedText;
        BtnRemove.Visibility = canRemove ? Visibility.Visible : Visibility.Collapsed;

        // OK 버튼은 URL 이 비어있지 않을 때만 활성
        BtnOk.IsEnabled = existingUrl.Length > 0;

        Loaded += (_, _) =>
        {
            TxtUrl.Focus();
            TxtUrl.SelectAll();
        };
    }

    // ── 이벤트 핸들러 ──────────────────────────────────────────────────────

    private void OnUrlChanged(object sender, TextChangedEventArgs e)
        => BtnOk.IsEnabled = TxtUrl.Text.Trim().Length > 0;

    private void OnOk(object sender, RoutedEventArgs e)
    {
        var url = TxtUrl.Text.Trim();
        if (url.Length == 0)
        {
            TxtUrl.Focus();
            return;
        }

        // 스킴이 없으면 https:// 자동 보완
        if (!url.Contains("://", StringComparison.Ordinal)
            && !url.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase)
            && !url.StartsWith("#", StringComparison.Ordinal))
        {
            url = "https://" + url;
        }

        ResultUrl = url;
        var display = TxtDisplayText.Text.Trim();
        ResultDisplayText = display.Length > 0 ? display : null;
        DialogResult = true;
        Close();
    }

    private void OnCancel(object sender, RoutedEventArgs e) => Close();

    private void OnRemove(object sender, RoutedEventArgs e)
    {
        ResultUrl = null;  // null = 링크 제거
        DialogResult = true;
        Close();
    }
}

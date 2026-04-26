using System.Windows;
using PolyDoc.Iwpf;
using SR = PolyDoc.App.Properties.Resources;

namespace PolyDoc.App.Views;

/// <summary>
/// 비밀번호 보호 모드 설정/변경 다이얼로그.
/// 열기/쓰기 체크박스로 각각 독립 설정, 같이사용 체크 시 단일 비밀번호 패널 표시.
/// </summary>
public partial class PasswordChangeWindow : Window
{
    public PasswordChangeWindow(PasswordMode currentMode = PasswordMode.None)
    {
        InitializeComponent();
        InitializeCheckBoxes(currentMode);
        Loaded += OnWindowLoaded;
    }

    /// <summary>사용자가 선택한 보호 모드.</summary>
    public PasswordMode ResultMode { get; private set; }

    /// <summary>열기(AES) 암호. Read / Both 모드 시 설정됨.</summary>
    public string? ResultReadPassword { get; private set; }

    /// <summary>쓰기 잠금 암호. Write / Both 모드 시 설정됨.</summary>
    public string? ResultWritePassword { get; private set; }

    private void InitializeCheckBoxes(PasswordMode mode)
    {
        switch (mode)
        {
            case PasswordMode.Read:
                ChkRead.IsChecked  = true;
                ChkWrite.IsChecked = false;
                break;
            case PasswordMode.Write:
                ChkRead.IsChecked  = false;
                ChkWrite.IsChecked = true;
                break;
            case PasswordMode.Both:
                ChkRead.IsChecked  = true;
                ChkWrite.IsChecked = true;
                ChkSame.IsEnabled  = true;
                ChkSame.IsChecked  = true; // 기존 Both 는 기본적으로 같은 암호로 열기
                break;
            default:
                ChkRead.IsChecked  = false;
                ChkWrite.IsChecked = false;
                break;
        }
        UpdatePanelVisibility();
    }

    private void OnWindowLoaded(object sender, RoutedEventArgs e)
    {
        if (PanelSingle.Visibility == Visibility.Visible)
            PwdNew1.Focus();
        else if (PanelDual.Visibility == Visibility.Visible)
            PwdReadNew.Focus();
    }

    private void OnCheckChanged(object sender, RoutedEventArgs e)
    {
        var bothChecked = ChkRead.IsChecked == true && ChkWrite.IsChecked == true;
        ChkSame.IsEnabled = bothChecked;
        if (!bothChecked)
            ChkSame.IsChecked = false;

        UpdatePanelVisibility();
        ErrorText.Visibility = Visibility.Collapsed;
    }

    private void UpdatePanelVisibility()
    {
        var read  = ChkRead.IsChecked  == true;
        var write = ChkWrite.IsChecked == true;
        var same  = ChkSame.IsChecked  == true;

        // 이중 패널: 열기+쓰기 둘 다 체크 AND 같이사용 미체크
        var dual = read && write && !same;
        PanelDual.Visibility   = dual                  ? Visibility.Visible : Visibility.Collapsed;
        PanelSingle.Visibility = (read || write) && !dual ? Visibility.Visible : Visibility.Collapsed;
    }

    private void OnOk(object sender, RoutedEventArgs e)
    {
        var read  = ChkRead.IsChecked  == true;
        var write = ChkWrite.IsChecked == true;
        var same  = ChkSame.IsChecked  == true;

        if (!read && !write)
        {
            ResultMode          = PasswordMode.None;
            ResultReadPassword  = null;
            ResultWritePassword = null;
            DialogResult        = true;
            return;
        }

        if (read && write && !same)
        {
            // 이중 패널: 열기·쓰기 각각 다른 비밀번호
            var rPwd  = PwdReadNew.Password;
            var rConf = PwdReadConfirm.Password;
            var wPwd  = PwdWriteNew.Password;
            var wConf = PwdWriteConfirm.Password;

            if (string.IsNullOrEmpty(rPwd) || string.IsNullOrEmpty(wPwd))
            {
                ShowError(SR.PwdEmpty);
                return;
            }
            if (rPwd != rConf || wPwd != wConf)
            {
                ShowError(SR.PwdMismatch);
                return;
            }

            ResultMode          = PasswordMode.Both;
            ResultReadPassword  = rPwd;
            ResultWritePassword = wPwd;
        }
        else
        {
            // 단일 패널: 열기만, 쓰기만, 또는 둘 다 같은 암호
            var pwd  = PwdNew1.Password;
            var conf = PwdConfirm1.Password;

            if (string.IsNullOrEmpty(pwd))
            {
                ShowError(SR.PwdEmpty);
                return;
            }
            if (pwd != conf)
            {
                ShowError(SR.PwdMismatch);
                return;
            }

            ResultMode          = (read && write) ? PasswordMode.Both
                                : read            ? PasswordMode.Read
                                :                   PasswordMode.Write;
            ResultReadPassword  = read  ? pwd : null;
            ResultWritePassword = write ? pwd : null;
        }

        DialogResult = true;
    }

    private void ShowError(string message)
    {
        ErrorText.Text       = message;
        ErrorText.Visibility = Visibility.Visible;
    }

    private void OnCancel(object sender, RoutedEventArgs e) => DialogResult = false;
}

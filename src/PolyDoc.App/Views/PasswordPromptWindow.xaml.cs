using System.Windows;
using SR = PolyDoc.App.Properties.Resources;

namespace PolyDoc.App.Views;

/// <summary>
/// 암호화된 IWPF 문서를 열 때 한 번 사용하는 비밀번호 입력 다이얼로그.
/// 호출자는 <see cref="ShowDialog"/> 의 결과가 true 면 <see cref="EnteredPassword"/> 를 읽는다.
/// 잘못된 비밀번호로 거부될 경우 <see cref="ShowError"/> 호출 후 다시 ShowDialog 를 띄울 수 있다.
/// </summary>
public partial class PasswordPromptWindow : Window
{
    public PasswordPromptWindow()
    {
        InitializeComponent();
        Loaded += (_, _) => PasswordInput.Focus();
    }

    /// <summary>사용자가 [확인] 으로 닫았을 때의 입력 비밀번호.</summary>
    public string EnteredPassword { get; private set; } = "";

    /// <summary>메시지 텍스트를 재정의한다 (쓰기 보호 프롬프트 등 맞춤 안내문용).</summary>
    public void SetMessage(string message) => MessageText.Text = message;

    public void ShowError(string message)
    {
        ErrorText.Text = message;
        ErrorText.Visibility = Visibility.Visible;
    }

    private void OnOk(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(PasswordInput.Password))
        {
            ShowError(SR.PwdWrong);
            return;
        }
        EnteredPassword = PasswordInput.Password;
        DialogResult = true;
    }

    private void OnCancel(object sender, RoutedEventArgs e) => DialogResult = false;
}

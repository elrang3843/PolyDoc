using System.Windows;

namespace PolyDonky.App.Services;

/// <summary>
/// WPF 다이얼로그 헬퍼 메서드 — Owner 설정 + ShowDialog 일괄 처리.
/// </summary>
public static class DialogHelper
{
    /// <summary>
    /// 다이얼로그를 모달로 표시. Owner가 지정되지 않으면 메인 윈도우로 설정.
    /// </summary>
    /// <typeparam name="T">Window 파생 클래스</typeparam>
    /// <param name="dialog">표시할 다이얼로그</param>
    /// <param name="owner">Owner 윈도우 (null이면 Application.Current.MainWindow)</param>
    /// <returns>다이얼로그 결과 (DialogResult)</returns>
    public static bool? ShowModal<T>(this T dialog, Window? owner = null) where T : Window
    {
        dialog.Owner = owner ?? Application.Current.MainWindow;
        return dialog.ShowDialog();
    }

    /// <summary>
    /// 주어진 컨트롤이 속한 윈도우를 기준으로 다이얼로그 표시.
    /// </summary>
    /// <typeparam name="T">Window 파생 클래스</typeparam>
    /// <param name="dialog">표시할 다이얼로그</param>
    /// <param name="ownerElement">Owner를 결정할 UI 요소 (보통 this)</param>
    /// <returns>다이얼로그 결과 (DialogResult)</returns>
    public static bool? ShowModalFor<T>(this T dialog, FrameworkElement ownerElement) where T : Window
    {
        dialog.Owner = Window.GetWindow(ownerElement) ?? Application.Current.MainWindow;
        return dialog.ShowDialog();
    }
}

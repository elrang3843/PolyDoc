using System.Globalization;
using System.Threading;
using System.Windows;

namespace PolyDoc.App;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        // 기본 UI 문화권을 한국어로 고정 (도구 → 설정에서 변경 시 갱신).
        var ko = new CultureInfo("ko-KR");
        Thread.CurrentThread.CurrentUICulture = ko;
        CultureInfo.DefaultThreadCurrentUICulture = ko;

        base.OnStartup(e);
    }
}

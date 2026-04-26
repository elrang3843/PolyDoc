using System.Windows;
using PolyDoc.App.Services;

namespace PolyDoc.App;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        // 저장된 언어 설정을 불러와 culture 및 LocalizedStrings 에 적용한다.
        // 기본은 한국어; 설정 파일이 없으면 ko-KR 로 초기화된다.
        LanguageService.LoadAndApply();
    }
}

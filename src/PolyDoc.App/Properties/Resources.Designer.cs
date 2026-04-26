#nullable enable
using System.Globalization;
using System.Resources;

namespace PolyDoc.App.Properties
{
    internal static class Resources
    {
        private static ResourceManager? _resourceMan;

        internal static ResourceManager ResourceManager
        {
            get
            {
                _resourceMan ??= new ResourceManager(
                    "PolyDoc.App.Properties.Resources",
                    typeof(Resources).Assembly);
                return _resourceMan;
            }
        }

        internal static CultureInfo? Culture { get; set; }

        private static string Get(string name)
            => ResourceManager.GetString(name, Culture) ?? name;

        // ── 메뉴: 파일 ─────────────────────────────────────────────
        internal static string MenuFile              => Get(nameof(MenuFile));
        internal static string MenuFileNew           => Get(nameof(MenuFileNew));
        internal static string MenuFileOpen          => Get(nameof(MenuFileOpen));
        internal static string MenuFileSave          => Get(nameof(MenuFileSave));
        internal static string MenuFileSaveAs        => Get(nameof(MenuFileSaveAs));
        internal static string MenuFilePreview       => Get(nameof(MenuFilePreview));
        internal static string MenuFilePrint         => Get(nameof(MenuFilePrint));
        internal static string MenuFileClose         => Get(nameof(MenuFileClose));
        internal static string MenuFileExit          => Get(nameof(MenuFileExit));

        // ── 메뉴: 편집 ─────────────────────────────────────────────
        internal static string MenuEdit              => Get(nameof(MenuEdit));
        internal static string MenuEditCopy          => Get(nameof(MenuEditCopy));
        internal static string MenuEditCut           => Get(nameof(MenuEditCut));
        internal static string MenuEditPaste         => Get(nameof(MenuEditPaste));
        internal static string MenuEditDelete        => Get(nameof(MenuEditDelete));
        internal static string MenuEditFindReplace   => Get(nameof(MenuEditFindReplace));
        internal static string MenuEditDocInfo       => Get(nameof(MenuEditDocInfo));

        // ── 메뉴: 입력 ─────────────────────────────────────────────
        internal static string MenuInsert            => Get(nameof(MenuInsert));
        internal static string MenuInsertTextBox     => Get(nameof(MenuInsertTextBox));
        internal static string MenuInsertTable       => Get(nameof(MenuInsertTable));
        internal static string MenuInsertGraph       => Get(nameof(MenuInsertGraph));
        internal static string MenuInsertSpecialChar => Get(nameof(MenuInsertSpecialChar));
        internal static string MenuInsertEquation    => Get(nameof(MenuInsertEquation));
        internal static string MenuInsertEmoji       => Get(nameof(MenuInsertEmoji));
        internal static string MenuInsertShape       => Get(nameof(MenuInsertShape));
        internal static string MenuInsertImage       => Get(nameof(MenuInsertImage));
        internal static string MenuInsertSign        => Get(nameof(MenuInsertSign));

        // ── 메뉴: 서식 ─────────────────────────────────────────────
        internal static string MenuFormat            => Get(nameof(MenuFormat));
        internal static string MenuFormatChar        => Get(nameof(MenuFormatChar));
        internal static string MenuFormatPara        => Get(nameof(MenuFormatPara));
        internal static string MenuFormatPage        => Get(nameof(MenuFormatPage));

        // ── 메뉴: 도구 ─────────────────────────────────────────────
        internal static string MenuTools             => Get(nameof(MenuTools));
        internal static string MenuToolsSettings     => Get(nameof(MenuToolsSettings));
        internal static string MenuToolsDict         => Get(nameof(MenuToolsDict));
        internal static string MenuToolsSpell        => Get(nameof(MenuToolsSpell));
        internal static string MenuToolsSignMaker    => Get(nameof(MenuToolsSignMaker));

        // ── 메뉴: 도움말 ────────────────────────────────────────────
        internal static string MenuHelp              => Get(nameof(MenuHelp));
        internal static string MenuHelpManual        => Get(nameof(MenuHelpManual));
        internal static string MenuHelpLicense       => Get(nameof(MenuHelpLicense));
        internal static string MenuHelpAbout         => Get(nameof(MenuHelpAbout));

        // ── 찾기/바꾸기 다이얼로그 ──────────────────────────────────
        internal static string FindReplaceTitle       => Get(nameof(FindReplaceTitle));
        internal static string FindReplaceFind        => Get(nameof(FindReplaceFind));
        internal static string FindReplaceReplace     => Get(nameof(FindReplaceReplace));
        internal static string FindReplaceCaseSensitive => Get(nameof(FindReplaceCaseSensitive));
        internal static string FindReplaceFindNext    => Get(nameof(FindReplaceFindNext));
        internal static string FindReplaceReplaceOne  => Get(nameof(FindReplaceReplaceOne));
        internal static string FindReplaceReplaceAll  => Get(nameof(FindReplaceReplaceAll));
        internal static string FindReplaceClose       => Get(nameof(FindReplaceClose));
        internal static string FindReplaceEnterQuery  => Get(nameof(FindReplaceEnterQuery));
        internal static string FindReplaceNotFound    => Get(nameof(FindReplaceNotFound));
        internal static string FindReplaceWrapped     => Get(nameof(FindReplaceWrapped));
        internal static string FindReplaceReplaced    => Get(nameof(FindReplaceReplaced));

        // ── 상태 표시줄 ─────────────────────────────────────────────
        internal static string StatusReady           => Get(nameof(StatusReady));
        internal static string StatusNewDoc          => Get(nameof(StatusNewDoc));
        internal static string StatusOpenDone        => Get(nameof(StatusOpenDone));
        internal static string StatusSaveDone        => Get(nameof(StatusSaveDone));
        internal static string StatusInsert          => Get(nameof(StatusInsert));
        internal static string StatusOverwrite       => Get(nameof(StatusOverwrite));

        // ── 다이얼로그 공통 ─────────────────────────────────────────
        internal static string DlgOK                    => Get(nameof(DlgOK));
        internal static string DlgCancel                => Get(nameof(DlgCancel));
        internal static string DlgYes                   => Get(nameof(DlgYes));
        internal static string DlgNo                    => Get(nameof(DlgNo));
        internal static string DlgUnsavedPrompt         => Get(nameof(DlgUnsavedPrompt));
        internal static string DlgUnsupportedFormat     => Get(nameof(DlgUnsupportedFormat));
        internal static string DlgUnsupportedFormatTitle => Get(nameof(DlgUnsupportedFormatTitle));
        internal static string DlgUnknownFormat         => Get(nameof(DlgUnknownFormat));
        internal static string DlgOpenFail              => Get(nameof(DlgOpenFail));
        internal static string DlgSaveConverterNeeded   => Get(nameof(DlgSaveConverterNeeded));
        internal static string DlgSaveConverterTitle    => Get(nameof(DlgSaveConverterTitle));
        internal static string DlgSaveFail              => Get(nameof(DlgSaveFail));
        internal static string DlgSaveFailTitle         => Get(nameof(DlgSaveFailTitle));
        internal static string DlgOpenError             => Get(nameof(DlgOpenError));
        internal static string DlgSaveError             => Get(nameof(DlgSaveError));
        internal static string DlgNewDocTitle           => Get(nameof(DlgNewDocTitle));
        internal static string DlgFileOpenTitle         => Get(nameof(DlgFileOpenTitle));
        internal static string DlgFileSaveTitle         => Get(nameof(DlgFileSaveTitle));
        internal static string DlgDefaultFileName       => Get(nameof(DlgDefaultFileName));

        // ── About ────────────────────────────────────────────────────
        internal static string AboutTitle   => Get(nameof(AboutTitle));
        internal static string AboutProduct => Get(nameof(AboutProduct));
        internal static string AboutCompany => Get(nameof(AboutCompany));
        internal static string AboutAuthor  => Get(nameof(AboutAuthor));
        internal static string AboutLicense => Get(nameof(AboutLicense));
        internal static string AboutDesc    => Get(nameof(AboutDesc));

        // ── 설정 다이얼로그 ─────────────────────────────────────────
        internal static string SettingsTitle      => Get(nameof(SettingsTitle));
        internal static string SettingsTheme      => Get(nameof(SettingsTheme));
        internal static string SettingsThemeLight => Get(nameof(SettingsThemeLight));
        internal static string SettingsThemeDark  => Get(nameof(SettingsThemeDark));
        internal static string SettingsThemeSoft  => Get(nameof(SettingsThemeSoft));
        internal static string SettingsApply      => Get(nameof(SettingsApply));
        internal static string SettingsLanguage   => Get(nameof(SettingsLanguage));
        internal static string SettingsLangKorean => Get(nameof(SettingsLangKorean));
        internal static string SettingsLangEnglish => Get(nameof(SettingsLangEnglish));
    }
}

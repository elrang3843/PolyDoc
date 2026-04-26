using System.ComponentModel;
using SR = PolyDoc.App.Properties.Resources;

namespace PolyDoc.App.Services;

/// <summary>
/// XAML 바인딩용 다국어 문자열 싱글톤.
/// <c>{Binding MenuFile, Source={x:Static svc:LocalizedStrings.Instance}}</c> 형태로 사용.
/// <see cref="LanguageService.Apply"/> 호출 후 <see cref="Refresh"/> 가 실행되면
/// <see cref="INotifyPropertyChanged.PropertyChanged"/> 가 전체 갱신을 알려
/// 모든 바인딩이 새 언어 문자열로 재평가된다.
/// </summary>
public sealed class LocalizedStrings : INotifyPropertyChanged
{
    public static LocalizedStrings Instance { get; } = new();
    private LocalizedStrings() { }

    public event PropertyChangedEventHandler? PropertyChanged;

    /// <summary>언어 변경 후 호출 — 모든 바인딩 대상에 재평가 요청.</summary>
    public void Refresh() => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(string.Empty));

    // ── 메뉴: 파일 ──────────────────────────────────────────────
    public string MenuFile        => SR.MenuFile;
    public string MenuFileNew     => SR.MenuFileNew;
    public string MenuFileOpen    => SR.MenuFileOpen;
    public string MenuFileSave    => SR.MenuFileSave;
    public string MenuFileSaveAs  => SR.MenuFileSaveAs;
    public string MenuFilePreview => SR.MenuFilePreview;
    public string MenuFilePrint   => SR.MenuFilePrint;
    public string MenuFileClose   => SR.MenuFileClose;
    public string MenuFileExit    => SR.MenuFileExit;

    // ── 메뉴: 편집 ──────────────────────────────────────────────
    public string MenuEdit           => SR.MenuEdit;
    public string MenuEditCopy       => SR.MenuEditCopy;
    public string MenuEditCut        => SR.MenuEditCut;
    public string MenuEditPaste      => SR.MenuEditPaste;
    public string MenuEditDelete     => SR.MenuEditDelete;
    public string MenuEditFindReplace => SR.MenuEditFindReplace;
    public string MenuEditDocInfo    => SR.MenuEditDocInfo;

    // ── 메뉴: 입력 ──────────────────────────────────────────────
    public string MenuInsert          => SR.MenuInsert;
    public string MenuInsertTextBox   => SR.MenuInsertTextBox;
    public string MenuInsertTable     => SR.MenuInsertTable;
    public string MenuInsertGraph     => SR.MenuInsertGraph;
    public string MenuInsertSpecialChar => SR.MenuInsertSpecialChar;
    public string MenuInsertEquation  => SR.MenuInsertEquation;
    public string MenuInsertEmoji     => SR.MenuInsertEmoji;
    public string MenuInsertShape     => SR.MenuInsertShape;
    public string MenuInsertImage     => SR.MenuInsertImage;
    public string MenuInsertSign      => SR.MenuInsertSign;

    // ── 메뉴: 서식 ──────────────────────────────────────────────
    public string MenuFormat     => SR.MenuFormat;
    public string MenuFormatChar => SR.MenuFormatChar;
    public string MenuFormatPara => SR.MenuFormatPara;
    public string MenuFormatPage => SR.MenuFormatPage;

    // ── 메뉴: 도구 ──────────────────────────────────────────────
    public string MenuTools          => SR.MenuTools;
    public string MenuToolsSettings  => SR.MenuToolsSettings;
    public string MenuToolsDict      => SR.MenuToolsDict;
    public string MenuToolsSpell     => SR.MenuToolsSpell;
    public string MenuToolsSignMaker => SR.MenuToolsSignMaker;

    // ── 메뉴: 도움말 ────────────────────────────────────────────
    public string MenuHelp        => SR.MenuHelp;
    public string MenuHelpManual  => SR.MenuHelpManual;
    public string MenuHelpLicense => SR.MenuHelpLicense;
    public string MenuHelpAbout   => SR.MenuHelpAbout;

    // ── 찾기/바꾸기 다이얼로그 ──────────────────────────────────
    public string FindReplaceTitle       => SR.FindReplaceTitle;
    public string FindReplaceFind        => SR.FindReplaceFind;
    public string FindReplaceReplace     => SR.FindReplaceReplace;
    public string FindReplaceCaseSensitive => SR.FindReplaceCaseSensitive;
    public string FindReplaceFindNext    => SR.FindReplaceFindNext;
    public string FindReplaceReplaceOne  => SR.FindReplaceReplaceOne;
    public string FindReplaceReplaceAll  => SR.FindReplaceReplaceAll;
    public string FindReplaceClose       => SR.FindReplaceClose;

    // ── 문서 정보 다이얼로그 ────────────────────────────────────
    public string DocInfoTitle      => SR.DocInfoTitle;
    public string DocInfoGroupFile  => SR.DocInfoGroupFile;
    public string DocInfoGroupMeta  => SR.DocInfoGroupMeta;
    public string DocInfoGroupStats => SR.DocInfoGroupStats;
    public string DocInfoPath       => SR.DocInfoPath;
    public string DocInfoFormat     => SR.DocInfoFormat;
    public string DocInfoDataSize   => SR.DocInfoDataSize;
    public string DocInfoDocTitle   => SR.DocInfoDocTitle;
    public string DocInfoAuthor     => SR.DocInfoAuthor;
    public string DocInfoLanguage   => SR.DocInfoLanguage;
    public string DocInfoCreated    => SR.DocInfoCreated;
    public string DocInfoModified   => SR.DocInfoModified;
    public string DocInfoParas      => SR.DocInfoParas;
    public string DocInfoChars      => SR.DocInfoChars;
    public string DocInfoWords      => SR.DocInfoWords;
    public string DocInfoLines      => SR.DocInfoLines;
    public string DocInfoSections   => SR.DocInfoSections;
    public string DocInfoTables     => SR.DocInfoTables;
    public string DocInfoImages     => SR.DocInfoImages;
    public string DocInfoTabInfo      => SR.DocInfoTabInfo;
    public string DocInfoTabSecurity  => SR.DocInfoTabSecurity;
    public string DocInfoTabWatermark => SR.DocInfoTabWatermark;
    public string DocInfoPwdStatus    => SR.DocInfoPwdStatus;
    public string DocInfoPwdChange    => SR.DocInfoPwdChange;
    public string DocInfoPwdHint      => SR.DocInfoPwdHint;
    public string DocInfoPwdNotIwpf   => SR.DocInfoPwdNotIwpf;
    public string DocInfoWmEnabled    => SR.DocInfoWmEnabled;
    public string DocInfoWmText       => SR.DocInfoWmText;
    public string DocInfoWmFontSize   => SR.DocInfoWmFontSize;
    public string DocInfoWmRotation   => SR.DocInfoWmRotation;
    public string DocInfoWmOpacity    => SR.DocInfoWmOpacity;
    public string DocInfoWmColor      => SR.DocInfoWmColor;
    public string DocInfoWmHint       => SR.DocInfoWmHint;
    public string PwdPromptTitle        => SR.PwdPromptTitle;
    public string PwdPromptMessage      => SR.PwdPromptMessage;
    public string PwdPromptInput        => SR.PwdPromptInput;
    public string PwdChangeTitle        => SR.PwdChangeTitle;
    public string PwdChangeNew          => SR.PwdChangeNew;
    public string PwdChangeConfirm      => SR.PwdChangeConfirm;
    public string PwdChangeHintRemove   => SR.PwdChangeHintRemove;
    public string PwdWritePromptMessage => SR.PwdWritePromptMessage;
    public string PwdModeGroupLabel     => SR.PwdModeGroupLabel;
    public string PwdModeNone           => SR.PwdModeNone;
    public string PwdModeRead           => SR.PwdModeRead;
    public string PwdModeWrite          => SR.PwdModeWrite;
    public string PwdModeBoth           => SR.PwdModeBoth;
    public string PwdChkRead            => SR.PwdChkRead;
    public string PwdChkWrite           => SR.PwdChkWrite;
    public string PwdChkSame            => SR.PwdChkSame;
    public string StatusWriteProtected  => SR.StatusWriteProtected;
    public string DlgConfirm            => SR.DlgConfirm;

    // ── 글자 서식 다이얼로그 ────────────────────────────────────
    public string FormatCharTitle        => SR.FormatCharTitle;
    public string FormatCharFontGroup    => SR.FormatCharFontGroup;
    public string FormatCharFontFamily   => SR.FormatCharFontFamily;
    public string FormatCharFontSize     => SR.FormatCharFontSize;
    public string FormatCharStyleGroup   => SR.FormatCharStyleGroup;
    public string FormatCharBold         => SR.FormatCharBold;
    public string FormatCharItalic       => SR.FormatCharItalic;
    public string FormatCharUnderline    => SR.FormatCharUnderline;
    public string FormatCharStrikethrough => SR.FormatCharStrikethrough;
    public string FormatCharOverline     => SR.FormatCharOverline;
    public string FormatCharSuperscript  => SR.FormatCharSuperscript;
    public string FormatCharSubscript    => SR.FormatCharSubscript;
    public string FormatCharColorGroup   => SR.FormatCharColorGroup;
    public string FormatCharFgColor      => SR.FormatCharFgColor;
    public string FormatCharBgColor      => SR.FormatCharBgColor;
    public string FormatCharPreviewGroup => SR.FormatCharPreviewGroup;
    public string FormatCharWidthPercent  => SR.FormatCharWidthPercent;
    public string FormatCharLetterSpacing => SR.FormatCharLetterSpacing;

    // ── 설정 다이얼로그 ─────────────────────────────────────────
    public string SettingsTitle       => SR.SettingsTitle;
    public string SettingsTheme       => SR.SettingsTheme;
    public string SettingsThemeLight  => SR.SettingsThemeLight;
    public string SettingsThemeDark   => SR.SettingsThemeDark;
    public string SettingsThemeSoft   => SR.SettingsThemeSoft;
    public string SettingsLanguage    => SR.SettingsLanguage;
    public string SettingsLangKorean  => SR.SettingsLangKorean;
    public string SettingsLangEnglish => SR.SettingsLangEnglish;
    public string FindReplaceCloseBtn => SR.FindReplaceClose;
}

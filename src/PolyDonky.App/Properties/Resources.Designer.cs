#nullable enable
using System.Globalization;
using System.Resources;

namespace PolyDonky.App.Properties
{
    internal static class Resources
    {
        private static ResourceManager? _resourceMan;

        internal static ResourceManager ResourceManager
        {
            get
            {
                _resourceMan ??= new ResourceManager(
                    "PolyDonky.App.Properties.Resources",
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
        internal static string MenuFilePreviewPrint  => Get(nameof(MenuFilePreviewPrint));
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
        internal static string MenuInsertSpecialChar => Get(nameof(MenuInsertSpecialChar));
        internal static string MenuInsertEquation    => Get(nameof(MenuInsertEquation));
        internal static string MenuInsertEmoji       => Get(nameof(MenuInsertEmoji));
        internal static string MenuInsertShape       => Get(nameof(MenuInsertShape));
        internal static string MenuInsertImage       => Get(nameof(MenuInsertImage));
        // ── 메뉴: 서식 ─────────────────────────────────────────────
        internal static string MenuFormat            => Get(nameof(MenuFormat));
        internal static string MenuFormatChar        => Get(nameof(MenuFormatChar));
        internal static string MenuFormatPara        => Get(nameof(MenuFormatPara));
        internal static string MenuFormatPage        => Get(nameof(MenuFormatPage));

        // ── 메뉴: 보기 ─────────────────────────────────────────────
        internal static string MenuView                   => Get(nameof(MenuView));
        internal static string MenuViewTypesettingMarks   => Get(nameof(MenuViewTypesettingMarks));

        // ── 조판부호 보기 레이블 ────────────────────────────────────
        internal static string TyposettingMarkHeader       => Get(nameof(TyposettingMarkHeader));
        internal static string TyposettingMarkFooter       => Get(nameof(TyposettingMarkFooter));
        internal static string TyposettingMarkImage        => Get(nameof(TyposettingMarkImage));
        internal static string TyposettingMarkImageBehind  => Get(nameof(TyposettingMarkImageBehind));
        internal static string TyposettingMarkShape        => Get(nameof(TyposettingMarkShape));
        internal static string TyposettingMarkShapeBehind  => Get(nameof(TyposettingMarkShapeBehind));
        internal static string TyposettingMarkTable        => Get(nameof(TyposettingMarkTable));
        internal static string TyposettingMarkTableBehind  => Get(nameof(TyposettingMarkTableBehind));
        internal static string TyposettingMarkTextBox      => Get(nameof(TyposettingMarkTextBox));

        // ── 메뉴: 도구 ─────────────────────────────────────────────
        internal static string MenuTools             => Get(nameof(MenuTools));
        internal static string MenuToolsSettings     => Get(nameof(MenuToolsSettings));
        internal static string MenuToolsDict         => Get(nameof(MenuToolsDict));
        // ── 메뉴: 도움말 ────────────────────────────────────────────
        internal static string MenuHelp              => Get(nameof(MenuHelp));
        internal static string MenuHelpManual        => Get(nameof(MenuHelpManual));
        internal static string MenuHelpFormat        => Get(nameof(MenuHelpFormat));
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
        internal static string StatusBusyOpen        => Get(nameof(StatusBusyOpen));
        internal static string StatusBusySave        => Get(nameof(StatusBusySave));
        internal static string StatusBusyConvert     => Get(nameof(StatusBusyConvert));
        internal static string StatusBusyConvertProgress => Get(nameof(StatusBusyConvertProgress));
        internal static string StatusConvertedAndOpened => Get(nameof(StatusConvertedAndOpened));
        internal static string StatusSavedAndConverted  => Get(nameof(StatusSavedAndConverted));
        internal static string DlgConvertOnOpenTitle    => Get(nameof(DlgConvertOnOpenTitle));
        internal static string DlgConvertOnOpenPrompt   => Get(nameof(DlgConvertOnOpenPrompt));
        internal static string DlgConvertOnSaveTitle    => Get(nameof(DlgConvertOnSaveTitle));
        internal static string DlgConvertOnSavePrompt   => Get(nameof(DlgConvertOnSavePrompt));
        internal static string DlgConvertOverwriteTitle  => Get(nameof(DlgConvertOverwriteTitle));
        internal static string DlgConvertOverwritePrompt => Get(nameof(DlgConvertOverwritePrompt));
        internal static string DlgUnsupportedVersionTitle  => Get(nameof(DlgUnsupportedVersionTitle));
        internal static string DlgUnsupportedVersionPrompt => Get(nameof(DlgUnsupportedVersionPrompt));
        internal static string DlgHtmlTruncatedTitle => Get(nameof(DlgHtmlTruncatedTitle));
        internal static string DlgHtmlTruncatedPrompt => Get(nameof(DlgHtmlTruncatedPrompt));
        internal static string StatusInsert          => Get(nameof(StatusInsert));
        internal static string StatusOverwrite       => Get(nameof(StatusOverwrite));
        internal static string StatusDocClosed       => Get(nameof(StatusDocClosed));
        internal static string StatusDrawTextBox     => Get(nameof(StatusDrawTextBox));
        internal static string StatusDrawShape       => Get(nameof(StatusDrawShape));
        internal static string StatusDrawPolyline   => Get(nameof(StatusDrawPolyline));

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

        // ── 문서 정보 다이얼로그 ────────────────────────────────────
        internal static string DocInfoTitle      => Get(nameof(DocInfoTitle));
        internal static string DocInfoGroupFile  => Get(nameof(DocInfoGroupFile));
        internal static string DocInfoGroupMeta  => Get(nameof(DocInfoGroupMeta));
        internal static string DocInfoGroupStats => Get(nameof(DocInfoGroupStats));
        internal static string DocInfoPath       => Get(nameof(DocInfoPath));
        internal static string DocInfoFormat     => Get(nameof(DocInfoFormat));
        internal static string DocInfoDataSize   => Get(nameof(DocInfoDataSize));
        internal static string DocInfoDocTitle   => Get(nameof(DocInfoDocTitle));
        internal static string DocInfoAuthor     => Get(nameof(DocInfoAuthor));
        internal static string DocInfoEditor     => Get(nameof(DocInfoEditor));
        internal static string DocInfoLanguage   => Get(nameof(DocInfoLanguage));
        internal static string DocInfoCreated    => Get(nameof(DocInfoCreated));
        internal static string DocInfoModified   => Get(nameof(DocInfoModified));
        internal static string DocInfoParas      => Get(nameof(DocInfoParas));
        internal static string DocInfoChars      => Get(nameof(DocInfoChars));
        internal static string DocInfoWords      => Get(nameof(DocInfoWords));
        internal static string DocInfoLines      => Get(nameof(DocInfoLines));
        internal static string DocInfoSections   => Get(nameof(DocInfoSections));
        internal static string DocInfoTables     => Get(nameof(DocInfoTables));
        internal static string DocInfoImages     => Get(nameof(DocInfoImages));
        internal static string DocInfoNotSaved   => Get(nameof(DocInfoNotSaved));
        internal static string DocInfoNone       => Get(nameof(DocInfoNone));

        // 탭 + 보안 + 워터마크
        internal static string DocInfoTabInfo      => Get(nameof(DocInfoTabInfo));
        internal static string DocInfoTabSecurity  => Get(nameof(DocInfoTabSecurity));
        internal static string DocInfoTabWatermark => Get(nameof(DocInfoTabWatermark));
        internal static string DocInfoPwdStatus    => Get(nameof(DocInfoPwdStatus));
        internal static string DocInfoPwdSet       => Get(nameof(DocInfoPwdSet));
        internal static string DocInfoPwdNone      => Get(nameof(DocInfoPwdNone));
        internal static string DocInfoPwdChange    => Get(nameof(DocInfoPwdChange));
        internal static string DocInfoPwdHint      => Get(nameof(DocInfoPwdHint));
        internal static string DocInfoPwdNotIwpf   => Get(nameof(DocInfoPwdNotIwpf));
        internal static string DocInfoWmEnabled    => Get(nameof(DocInfoWmEnabled));
        internal static string DocInfoWmText       => Get(nameof(DocInfoWmText));
        internal static string DocInfoWmFontSize   => Get(nameof(DocInfoWmFontSize));
        internal static string DocInfoWmRotation   => Get(nameof(DocInfoWmRotation));
        internal static string DocInfoWmOpacity    => Get(nameof(DocInfoWmOpacity));
        internal static string DocInfoWmColor      => Get(nameof(DocInfoWmColor));
        internal static string DocInfoWmHint       => Get(nameof(DocInfoWmHint));
        internal static string DocInfoWmPrintWith  => Get(nameof(DocInfoWmPrintWith));
        internal static string DocInfoIsPrintable  => Get(nameof(DocInfoIsPrintable));
        internal static string DocInfoNotPrintable => Get(nameof(DocInfoNotPrintable));
        internal static string DocInfoWmLockedHint => Get(nameof(DocInfoWmLockedHint));
        internal static string DocInfoUnlock       => Get(nameof(DocInfoUnlock));

        // 비밀번호 다이얼로그
        internal static string PwdPromptTitle        => Get(nameof(PwdPromptTitle));
        internal static string PwdPromptMessage      => Get(nameof(PwdPromptMessage));
        internal static string PwdPromptInput        => Get(nameof(PwdPromptInput));
        internal static string PwdChangeTitle        => Get(nameof(PwdChangeTitle));
        internal static string PwdChangeNew          => Get(nameof(PwdChangeNew));
        internal static string PwdChangeConfirm      => Get(nameof(PwdChangeConfirm));
        internal static string PwdChangeHintRemove   => Get(nameof(PwdChangeHintRemove));
        internal static string PwdMismatch           => Get(nameof(PwdMismatch));
        internal static string PwdEmpty              => Get(nameof(PwdEmpty));
        internal static string PwdWrong              => Get(nameof(PwdWrong));
        internal static string PwdWritePromptMessage   => Get(nameof(PwdWritePromptMessage));
        internal static string StatusWriteProtected       => Get(nameof(StatusWriteProtected));
        internal static string StatusWriteProtectedSuffix => Get(nameof(StatusWriteProtectedSuffix));
        internal static string StatusWriteUnlocked        => Get(nameof(StatusWriteUnlocked));

        // 비밀번호 보호 모드
        internal static string PwdModeGroupLabel   => Get(nameof(PwdModeGroupLabel));
        internal static string PwdModeNone         => Get(nameof(PwdModeNone));
        internal static string PwdModeRead         => Get(nameof(PwdModeRead));
        internal static string PwdModeWrite        => Get(nameof(PwdModeWrite));
        internal static string PwdModeBoth         => Get(nameof(PwdModeBoth));
        internal static string PwdChkRead          => Get(nameof(PwdChkRead));
        internal static string PwdChkWrite         => Get(nameof(PwdChkWrite));
        internal static string PwdChkSame          => Get(nameof(PwdChkSame));
        internal static string DocInfoPwdModeNone  => Get(nameof(DocInfoPwdModeNone));
        internal static string DocInfoPwdModeRead  => Get(nameof(DocInfoPwdModeRead));
        internal static string DocInfoPwdModeWrite => Get(nameof(DocInfoPwdModeWrite));
        internal static string DocInfoPwdModeBoth  => Get(nameof(DocInfoPwdModeBoth));

        // 공통
        internal static string DlgConfirm          => Get(nameof(DlgConfirm));

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

        // ── 글자 서식 다이얼로그 ────────────────────────────────────
        internal static string FormatCharTitle        => Get(nameof(FormatCharTitle));
        internal static string FormatCharFontGroup    => Get(nameof(FormatCharFontGroup));
        internal static string FormatCharFontFamily   => Get(nameof(FormatCharFontFamily));
        internal static string FormatCharFontSize     => Get(nameof(FormatCharFontSize));
        internal static string FormatCharStyleGroup   => Get(nameof(FormatCharStyleGroup));
        internal static string FormatCharBold         => Get(nameof(FormatCharBold));
        internal static string FormatCharItalic       => Get(nameof(FormatCharItalic));
        internal static string FormatCharUnderline    => Get(nameof(FormatCharUnderline));
        internal static string FormatCharStrikethrough => Get(nameof(FormatCharStrikethrough));
        internal static string FormatCharOverline     => Get(nameof(FormatCharOverline));
        internal static string FormatCharSuperscript  => Get(nameof(FormatCharSuperscript));
        internal static string FormatCharSubscript    => Get(nameof(FormatCharSubscript));
        internal static string FormatCharColorGroup   => Get(nameof(FormatCharColorGroup));
        internal static string FormatCharFgColor      => Get(nameof(FormatCharFgColor));
        internal static string FormatCharBgColor      => Get(nameof(FormatCharBgColor));
        internal static string FormatCharPreviewGroup => Get(nameof(FormatCharPreviewGroup));
        internal static string FormatCharWidthPercent  => Get(nameof(FormatCharWidthPercent));
        internal static string FormatCharLetterSpacing => Get(nameof(FormatCharLetterSpacing));

        internal static string FormatParaTitle           => Get(nameof(FormatParaTitle));
        internal static string FormatParaAlignGroup      => Get(nameof(FormatParaAlignGroup));
        internal static string FormatParaAlignLeft       => Get(nameof(FormatParaAlignLeft));
        internal static string FormatParaAlignCenter     => Get(nameof(FormatParaAlignCenter));
        internal static string FormatParaAlignRight      => Get(nameof(FormatParaAlignRight));
        internal static string FormatParaAlignJustify    => Get(nameof(FormatParaAlignJustify));
        internal static string FormatParaAlignDistributed => Get(nameof(FormatParaAlignDistributed));
        internal static string FormatParaSpacingGroup    => Get(nameof(FormatParaSpacingGroup));
        internal static string FormatParaLineHeight      => Get(nameof(FormatParaLineHeight));
        internal static string FormatParaSpaceBefore     => Get(nameof(FormatParaSpaceBefore));
        internal static string FormatParaSpaceAfter      => Get(nameof(FormatParaSpaceAfter));
        internal static string FormatParaIndentGroup     => Get(nameof(FormatParaIndentGroup));
        internal static string FormatParaIndentFirst     => Get(nameof(FormatParaIndentFirst));
        internal static string FormatParaIndentLeft      => Get(nameof(FormatParaIndentLeft));
        internal static string FormatParaIndentRight     => Get(nameof(FormatParaIndentRight));
        internal static string FormatParaOutline         => Get(nameof(FormatParaOutline));
        internal static string FormatParaOutlineBody     => Get(nameof(FormatParaOutlineBody));
        internal static string FormatParaPreviewGroup    => Get(nameof(FormatParaPreviewGroup));

        internal static string MenuFormatOutline             => Get(nameof(MenuFormatOutline));
        internal static string FormatOutlineTitle            => Get(nameof(FormatOutlineTitle));
        internal static string FormatOutlinePreset           => Get(nameof(FormatOutlinePreset));
        internal static string FormatOutlineLevels           => Get(nameof(FormatOutlineLevels));
        internal static string FormatOutlineLevelBody        => Get(nameof(FormatOutlineLevelBody));
        internal static string FormatOutlineCharGroup        => Get(nameof(FormatOutlineCharGroup));
        internal static string FormatOutlineParaGroup        => Get(nameof(FormatOutlineParaGroup));
        internal static string FormatOutlineNumberGroup      => Get(nameof(FormatOutlineNumberGroup));
        internal static string FormatOutlineBorderGroup      => Get(nameof(FormatOutlineBorderGroup));
        internal static string FormatOutlineBgGroup          => Get(nameof(FormatOutlineBgGroup));
        internal static string FormatOutlineEdit             => Get(nameof(FormatOutlineEdit));
        internal static string FormatOutlineNumberStyle      => Get(nameof(FormatOutlineNumberStyle));
        internal static string FormatOutlinePrefix           => Get(nameof(FormatOutlinePrefix));
        internal static string FormatOutlineSuffix           => Get(nameof(FormatOutlineSuffix));
        internal static string FormatOutlineBorderTop        => Get(nameof(FormatOutlineBorderTop));
        internal static string FormatOutlineBorderBottom     => Get(nameof(FormatOutlineBorderBottom));
        internal static string FormatOutlineBorderColor      => Get(nameof(FormatOutlineBorderColor));
        internal static string FormatOutlineBgColor          => Get(nameof(FormatOutlineBgColor));
        internal static string FormatOutlineBgNone           => Get(nameof(FormatOutlineBgNone));
        internal static string FormatOutlinePreviewGroup     => Get(nameof(FormatOutlinePreviewGroup));
        internal static string FormatOutlineResetLevel       => Get(nameof(FormatOutlineResetLevel));
        internal static string FormatOutlineApplyAll         => Get(nameof(FormatOutlineApplyAll));
        internal static string FormatOutlineNumberNone       => Get(nameof(FormatOutlineNumberNone));
        internal static string FormatOutlineNumberDecimal    => Get(nameof(FormatOutlineNumberDecimal));
        internal static string FormatOutlineNumberAlphaLower => Get(nameof(FormatOutlineNumberAlphaLower));
        internal static string FormatOutlineNumberAlphaUpper => Get(nameof(FormatOutlineNumberAlphaUpper));
        internal static string FormatOutlineNumberRomanLower => Get(nameof(FormatOutlineNumberRomanLower));
        internal static string FormatOutlineNumberRomanUpper => Get(nameof(FormatOutlineNumberRomanUpper));
        internal static string FormatOutlineNumberHangul     => Get(nameof(FormatOutlineNumberHangul));
        internal static string FormatOutlinePresetDefault    => Get(nameof(FormatOutlinePresetDefault));
        internal static string FormatOutlinePresetAcademic   => Get(nameof(FormatOutlinePresetAcademic));
        internal static string FormatOutlinePresetBusiness   => Get(nameof(FormatOutlinePresetBusiness));
        internal static string FormatOutlinePresetModern     => Get(nameof(FormatOutlinePresetModern));

        // ── 페이지 서식 ──
        internal static string PageFormatTitle             => Get(nameof(PageFormatTitle));
        internal static string PageFormatPaperTab          => Get(nameof(PageFormatPaperTab));
        internal static string PageFormatMarginsTab        => Get(nameof(PageFormatMarginsTab));
        internal static string PageFormatLayoutTab         => Get(nameof(PageFormatLayoutTab));
        internal static string PageFormatSizeLabel         => Get(nameof(PageFormatSizeLabel));
        internal static string PageFormatCustomSize        => Get(nameof(PageFormatCustomSize));
        internal static string PageFormatWidthLabel        => Get(nameof(PageFormatWidthLabel));
        internal static string PageFormatHeightLabel       => Get(nameof(PageFormatHeightLabel));
        internal static string PageFormatOrientation       => Get(nameof(PageFormatOrientation));
        internal static string PageFormatPortrait          => Get(nameof(PageFormatPortrait));
        internal static string PageFormatLandscape         => Get(nameof(PageFormatLandscape));
        internal static string PageFormatPaperColor        => Get(nameof(PageFormatPaperColor));
        internal static string PageFormatPaperColorDefault => Get(nameof(PageFormatPaperColorDefault));
        internal static string PageFormatMarginTop         => Get(nameof(PageFormatMarginTop));
        internal static string PageFormatMarginBottom      => Get(nameof(PageFormatMarginBottom));
        internal static string PageFormatMarginLeft        => Get(nameof(PageFormatMarginLeft));
        internal static string PageFormatMarginRight       => Get(nameof(PageFormatMarginRight));
        internal static string PageFormatMarginHeader      => Get(nameof(PageFormatMarginHeader));
        internal static string PageFormatMarginFooter      => Get(nameof(PageFormatMarginFooter));
        internal static string MenuEditImport              => Get(nameof(MenuEditImport));
        internal static string MenuEditExport              => Get(nameof(MenuEditExport));
        internal static string PageFormatColumns           => Get(nameof(PageFormatColumns));
        internal static string PageFormatColumnGap         => Get(nameof(PageFormatColumnGap));
        internal static string PageFormatColumnDividerVisible    => Get(nameof(PageFormatColumnDividerVisible));
        internal static string PageFormatColumnDividerColor      => Get(nameof(PageFormatColumnDividerColor));
        internal static string PageFormatColumnDividerThickness  => Get(nameof(PageFormatColumnDividerThickness));
        internal static string PageFormatColumnDividerStyle      => Get(nameof(PageFormatColumnDividerStyle));
        internal static string ColumnDividerStyleNone            => Get(nameof(ColumnDividerStyleNone));
        internal static string ColumnDividerStyleSolid           => Get(nameof(ColumnDividerStyleSolid));
        internal static string ColumnDividerStyleDashed          => Get(nameof(ColumnDividerStyleDashed));
        internal static string ColumnDividerStyleDotted          => Get(nameof(ColumnDividerStyleDotted));
        internal static string PageFormatPageNumberStart   => Get(nameof(PageFormatPageNumberStart));
        internal static string PageFormatShowMarginGuides  => Get(nameof(PageFormatShowMarginGuides));
        internal static string PageFormatPreviewGroup      => Get(nameof(PageFormatPreviewGroup));
        internal static string PageSizeGroupIsoA           => Get(nameof(PageSizeGroupIsoA));
        internal static string PageSizeGroupIsoB           => Get(nameof(PageSizeGroupIsoB));
        internal static string PageSizeGroupJisB           => Get(nameof(PageSizeGroupJisB));
        internal static string PageSizeGroupUS             => Get(nameof(PageSizeGroupUS));
        internal static string PageSizeGroupNewspaper      => Get(nameof(PageSizeGroupNewspaper));
        internal static string PageSizeGroupKoreanBook     => Get(nameof(PageSizeGroupKoreanBook));
        internal static string PageSizeGroupIntlBook       => Get(nameof(PageSizeGroupIntlBook));
        internal static string ToolbarZoomIn               => Get(nameof(ToolbarZoomIn));
        internal static string ToolbarZoomOut              => Get(nameof(ToolbarZoomOut));
        internal static string ToolbarFitWidth             => Get(nameof(ToolbarFitWidth));
        internal static string ToolbarFitPage              => Get(nameof(ToolbarFitPage));
        internal static string DictWindowTitle             => Get(nameof(DictWindowTitle));
        internal static string DictWindowSearchHint        => Get(nameof(DictWindowSearchHint));
        internal static string DictWindowSearch            => Get(nameof(DictWindowSearch));
        internal static string DictWindowErrNoRuntime      => Get(nameof(DictWindowErrNoRuntime));
    }
}

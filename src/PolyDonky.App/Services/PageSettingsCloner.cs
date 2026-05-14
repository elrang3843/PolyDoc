using PolyDonky.Core;

namespace PolyDonky.App.Services;

/// <summary>
/// PageSettings 의 얕은 복사 유틸리티.
/// 섹션별 페이지 서식 상속, 다이얼로그 상태 백업 등에 사용.
/// </summary>
public static class PageSettingsCloner
{
    /// <summary>PageSettings 얕은 복사. ColumnWidthsMm 등 가변 컬렉션도 깊은 복사.</summary>
    public static PageSettings Clone(PageSettings src) => new()
    {
        SizeKind                 = src.SizeKind,
        WidthMm                  = src.WidthMm,
        HeightMm                 = src.HeightMm,
        Orientation              = src.Orientation,
        TextOrientation          = src.TextOrientation,
        TextProgression          = src.TextProgression,
        PaperColor               = src.PaperColor,
        MarginTopMm              = src.MarginTopMm,
        MarginBottomMm           = src.MarginBottomMm,
        MarginLeftMm             = src.MarginLeftMm,
        MarginRightMm            = src.MarginRightMm,
        MarginHeaderMm           = src.MarginHeaderMm,
        MarginFooterMm           = src.MarginFooterMm,
        ColumnCount              = src.ColumnCount,
        ColumnGapMm              = src.ColumnGapMm,
        ColumnWidthsMm           = src.ColumnWidthsMm is { } cw ? new(cw) : null,
        ColumnDividerVisible     = src.ColumnDividerVisible,
        ColumnDividerColor       = src.ColumnDividerColor,
        ColumnDividerThicknessPt = src.ColumnDividerThicknessPt,
        ColumnDividerStyle       = src.ColumnDividerStyle,
        PageNumberStart          = src.PageNumberStart,
        Header                   = src.Header.Clone(),
        Footer                   = src.Footer.Clone(),
        DifferentFirstPage       = src.DifferentFirstPage,
        DifferentOddEven         = src.DifferentOddEven,
        ShowMarginGuides         = src.ShowMarginGuides,
    };
}

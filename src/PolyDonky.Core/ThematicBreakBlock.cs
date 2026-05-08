namespace PolyDonky.Core;

/// <summary>수평선(HR / thematic break) 블록의 선 스타일.</summary>
public enum ThematicLineStyle
{
    Solid = 0,
    Dashed,
    Dotted,
    Double,
    DashDot,
}

/// <summary>수평선(HR / thematic break) 블록.</summary>
public sealed class ThematicBreakBlock : Block
{
    /// <summary>선 색상(CSS hex 문자열, 예: "#000000"). null 이면 테마 기본 회색.</summary>
    public string? LineColor { get; set; }

    /// <summary>선 위아래 여백(pt). 0 이면 기본값(6 pt) 적용.</summary>
    public double MarginPt { get; set; }

    /// <summary>선 두께(pt). 0 이면 기본 1 pt. CSS <c>border-top:Npx ...</c> 의 두께를 보존한다.</summary>
    public double ThicknessPt { get; set; }

    /// <summary>선 종류(실선·파선·점선·이중선·일점쇄선). 기본 Solid.</summary>
    public ThematicLineStyle LineStyle { get; set; } = ThematicLineStyle.Solid;
}

namespace PolyDonky.Core;

/// <summary>수평선(HR / thematic break) 블록.</summary>
public sealed class ThematicBreakBlock : Block
{
    /// <summary>선 색상(CSS hex 문자열, 예: "#000000"). null 이면 테마 기본 회색.</summary>
    public string? LineColor { get; set; }

    /// <summary>선 위아래 여백(pt). 0 이면 기본값(6 pt) 적용.</summary>
    public double MarginPt { get; set; }
}

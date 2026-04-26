namespace PolyDoc.Core;

/// <summary>
/// 문서에 적용할 워터마크 설정. 현재는 텍스트 워터마크만 지원하며 페이지 가운데에
/// 회전·반투명 표시를 가정한다. 실제 렌더링은 후속 단계 (편집기/인쇄/PDF) 에서 구현.
/// </summary>
public sealed class WatermarkSettings
{
    /// <summary>true 면 워터마크가 활성화된다.</summary>
    public bool Enabled { get; set; }

    /// <summary>워터마크 본문 텍스트.</summary>
    public string Text { get; set; } = "";

    /// <summary>색상 (ARGB hex, 예: "#80808080" — 반투명 회색).</summary>
    public string Color { get; set; } = "#FF808080";

    /// <summary>폰트 크기 (포인트).</summary>
    public int FontSize { get; set; } = 48;

    /// <summary>회전 각도 (도, 시계방향 +). 보통 -45.</summary>
    public double Rotation { get; set; } = -45.0;

    /// <summary>불투명도 0.0 (완전 투명) ~ 1.0 (불투명).</summary>
    public double Opacity { get; set; } = 0.3;
}

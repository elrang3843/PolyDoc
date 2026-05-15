namespace PolyDonky.Core;

/// <summary>
/// 박스 스타일(테두리·배경·padding·margin) 을 가진 컨테이너 블록 — HTML 의 `&lt;div class="toc"&gt;`,
/// `&lt;div class="alert"&gt;`, `&lt;section&gt;`, `&lt;aside&gt;` 등 자식 블록을 묶는 컨테이너의 시각적 framing 을 표현.
///
/// 단일 단락만 가진 컨테이너는 <see cref="Paragraph"/> 의 4면 보더·배경·padding 으로도 표현 가능하지만,
/// 다중 단락/표/이미지를 묶는 박스는 컨테이너 단위로 framing 이 필요해 별도 블록을 둔다.
///
/// 라운드트립 시 원본 HTML 의 `class` 속성은 <see cref="ClassNames"/> 로 보존돼 후속 export 에서
/// 같은 클래스명으로 복원할 수 있다.
/// </summary>
public sealed class ContainerBlock : Block
{
    /// <summary>안쪽 자식 블록들. 단락·표·이미지·도형·중첩 컨테이너 모두 가능.</summary>
    public IList<Block> Children { get; set; } = new List<Block>();

    // ── 4면 테두리 ──────────────────────────────────────────────
    public double  BorderTopPt    { get; set; }
    public string? BorderTopColor { get; set; }
    public double  BorderRightPt    { get; set; }
    public string? BorderRightColor { get; set; }
    public double  BorderBottomPt    { get; set; }
    public string? BorderBottomColor { get; set; }
    public double  BorderLeftPt    { get; set; }
    public string? BorderLeftColor { get; set; }

    /// <summary>배경색(hex). null 이면 투명.</summary>
    public string? BackgroundColor { get; set; }

    // ── 안쪽 여백 (테두리 내부) ──────────────────────────────
    public double PaddingTopMm    { get; set; }
    public double PaddingRightMm  { get; set; }
    public double PaddingBottomMm { get; set; }
    public double PaddingLeftMm   { get; set; }

    // ── 바깥 여백 ──────────────────────────────────────────────
    public double MarginTopMm    { get; set; }
    public double MarginBottomMm { get; set; }

    /// <summary>너비(mm). 0 = 자동(부모 폭에 맞춤). 페이지 폭 미만이면 정렬에 따라 좌/우/가운데.</summary>
    public double WidthMm { get; set; }

    /// <summary>가로 정렬. 너비가 자동이 아닐 때만 의미가 있다.</summary>
    public ContainerHAlign HAlign { get; set; } = ContainerHAlign.Left;

    /// <summary>원본 HTML 의 class 속성(공백 구분). null/빈 문자열이면 클래스 없음.
    /// fidelity capsule + provenance 기반 클래스명 복원 export 에 활용.</summary>
    public string? ClassNames { get; set; }

    /// <summary>역할 힌트 — `.toc` / `.alert` / `.page-break` / `.header-sim` / 일반(`Generic`) 등.
    /// 클래스명 매핑이 안 되는 환경(MD/DOCX) 에서 의미 보존용.</summary>
    public ContainerRole Role { get; set; } = ContainerRole.Generic;
}

public enum ContainerHAlign { Left, Center, Right }

public enum ContainerRole
{
    Generic,
    /// <summary>목차 박스 (`<div class="toc">` 등).</summary>
    Toc,
    /// <summary>알림/경고 박스 (`<div class="alert">` 등).</summary>
    Alert,
    /// <summary>페이지 나누기 시각 표시 박스.</summary>
    PageBreakMarker,
    /// <summary>헤더/푸터 시뮬레이션 박스.</summary>
    HeaderFooterSim,
    /// <summary>인용 박스 (HTML blockquote 보다 강한 시각 framing).</summary>
    QuoteBox,
    /// <summary>사용자 또는 SVG 분해로 묶인 도형/블록 그룹. 해제(Ungroup) 가 가능하다.</summary>
    Group,
}

namespace PolyDonky.Core;

using System.Text.Json.Serialization;

/// <summary>도형 종류.</summary>
public enum ShapeKind
{
    Line,
    Polyline,
    Spline,
    Rectangle,
    RoundedRect,
    Ellipse,
    Triangle,
    RegularPolygon,
    Star,
    Polygon,
    ClosedSpline,
}

/// <summary>선 종류 (실선·파선·점선·일점쇄선).</summary>
public enum StrokeDash
{
    Solid,
    Dashed,
    Dotted,
    DashDot,
}

/// <summary>끝모양 종류 (선·폴리곤선·스플라인 선의 시작/끝에 적용).</summary>
public enum ShapeArrow
{
    None,
    Open,
    Filled,
    Diamond,
    Circle,
}

/// <summary>도형 레이블 텍스트의 가로 정렬.</summary>
public enum ShapeLabelHAlign { Left, Center, Right }

/// <summary>도형 레이블 텍스트의 세로 정렬.</summary>
public enum ShapeLabelVAlign { Top, Middle, Bottom }

/// <summary>도형 앵커 포인트. 좌표는 도형 바운딩 박스 좌상단 기준, 단위 mm.</summary>
public sealed class ShapePoint
{
    public double X { get; set; }
    public double Y { get; set; }

    /// <summary>이 점에서 다음 점으로 향하는 세그먼트의 나가는 베지어 제어점 X (mm, bbox 기준).
    /// null 이면 Catmull-Rom 자동 계산.</summary>
    public double? OutCtrlX { get; set; }

    /// <summary>이 점에서 다음 점으로 향하는 세그먼트의 나가는 베지어 제어점 Y (mm, bbox 기준).</summary>
    public double? OutCtrlY { get; set; }

    /// <summary>이전 점에서 이 점으로 들어오는 세그먼트의 베지어 제어점 X (mm, bbox 기준).
    /// null 이면 Catmull-Rom 자동 계산.</summary>
    public double? InCtrlX { get; set; }

    /// <summary>이전 점에서 이 점으로 들어오는 세그먼트의 베지어 제어점 Y (mm, bbox 기준).</summary>
    public double? InCtrlY { get; set; }
}

/// <summary>
/// 블록 단위 도형. 그림(<see cref="ImageBlock"/>)과 동일한 <see cref="ImageWrapMode"/> 배치 체계를 공유한다.
///
/// - WrapMode = Inline / WrapLeft / WrapRight : 본문 흐름 안에 블록으로 삽입.
/// - WrapMode = InFrontOfText / BehindText     : 캔버스 오버레이로 절대 위치 배치
///   (<see cref="OverlayXMm"/>, <see cref="OverlayYMm"/> 사용).
///
/// 도형 모양은 <see cref="Kind"/> 에 따라 달리 해석된다.
/// <list type="bullet">
///   <item>Line / Polyline / Spline : <see cref="Points"/> 에 꼭짓점·제어점 (mm) 목록 저장.</item>
///   <item>Rectangle / RoundedRect / Ellipse : 꼭짓점 불필요 — 바운딩 박스로 결정.</item>
///   <item>Triangle : <see cref="Points"/> 3개. 없으면 등변삼각형 기본값 사용.</item>
///   <item>RegularPolygon : <see cref="SideCount"/> 으로 꼭짓점 수 지정.</item>
///   <item>Star : <see cref="SideCount"/> (뾰족 수) + <see cref="InnerRadiusRatio"/> 로 결정.</item>
/// </list>
/// </summary>
public sealed class ShapeObject : Block, IOverlayAnchored
{
    // Line = 0 (enum default) 인 도형이 WhenWritingDefault 정책으로 JSON 에서 누락되어
    // 역직렬화 시 Rectangle(C# 기본값) 로 복원되던 버그 → 항상 직렬화하도록 명시.
    [JsonIgnore(Condition = JsonIgnoreCondition.Never)]
    public ShapeKind Kind { get; set; } = ShapeKind.Rectangle;

    /// <summary>본문과의 배치 관계. 이미지와 동일한 5모드 체계 공유.</summary>
    public ImageWrapMode WrapMode { get; set; } = ImageWrapMode.Inline;

    /// <summary>블록 내 가로 정렬 (WrapMode = Inline 일 때만 적용).</summary>
    public ImageHAlign HAlign { get; set; } = ImageHAlign.Left;

    /// <summary>바운딩 박스 너비 (mm).</summary>
    public double WidthMm { get; set; } = 40;

    /// <summary>바운딩 박스 높이 (mm). Line 계열은 선 두께를 고려한 최소값.</summary>
    public double HeightMm { get; set; } = 30;

    /// <summary>오버레이 anchoring — 0-based 페이지 인덱스 (0 = 첫 페이지).</summary>
    public int AnchorPageIndex { get; set; }

    /// <summary>오버레이 모드 X 위치 (mm, **해당 페이지 좌상단 기준**).</summary>
    public double OverlayXMm { get; set; }

    /// <summary>오버레이 모드 Y 위치 (mm, **해당 페이지 좌상단 기준**).</summary>
    public double OverlayYMm { get; set; }

    /// <summary>꼭짓점·제어점 목록 (mm, 바운딩 박스 좌상단 기준). Line/Polyline/Spline/Triangle 에 사용.</summary>
    public IList<ShapePoint> Points { get; set; } = new List<ShapePoint>();

    // ── 모양별 파라미터 ───────────────────────────────────────────────────────

    /// <summary>RegularPolygon / Star 의 꼭짓점(뾰족) 수. 3 이상. 기본 5.</summary>
    public int SideCount { get; set; } = 5;

    /// <summary>Star 의 내부 반지름 비율 (0 초과 ~ 1 미만). 0.5 = 일반 별.</summary>
    public double InnerRadiusRatio { get; set; } = 0.45;

    /// <summary>RoundedRect 의 모서리 반지름 (mm). 0 이하면 각진 모서리.</summary>
    public double CornerRadiusMm { get; set; } = 3;

    // ── 선 속성 ──────────────────────────────────────────────────────────────

    /// <summary>선 색상 hex (예: "#000000"). 빈 문자열·null 이면 검정.</summary>
    public string StrokeColor { get; set; } = "#000000";

    /// <summary>선 두께 (pt). 0 이하면 선 없음.</summary>
    public double StrokeThicknessPt { get; set; } = 1.0;

    /// <summary>선 종류 (실선·파선·점선·일점쇄선).</summary>
    public StrokeDash StrokeDash { get; set; } = StrokeDash.Solid;

    /// <summary>시작점 끝모양 (Line/Polyline/Spline 에만 의미 있음).</summary>
    public ShapeArrow StartArrow { get; set; } = ShapeArrow.None;

    /// <summary>끝점 끝모양 (Line/Polyline/Spline 에만 의미 있음).</summary>
    public ShapeArrow EndArrow { get; set; } = ShapeArrow.None;

    /// <summary>끝모양 크기 (mm). 0 이하면 선 두께에 비례한 자동 크기.</summary>
    public double EndShapeSizeMm { get; set; } = 0;

    // ── 채우기 속성 ───────────────────────────────────────────────────────────

    /// <summary>채우기 색상 hex. null / 빈 문자열 = 채우기 없음(투명).</summary>
    public string? FillColor { get; set; }

    /// <summary>채우기 불투명도 (0.0 = 완전 투명 ~ 1.0 = 완전 불투명). 기본 1.0.</summary>
    public double FillOpacity { get; set; } = 1.0;

    // ── 변환 ──────────────────────────────────────────────────────────────────

    /// <summary>도형 회전각 (도, 시계 방향). 바운딩 박스 중심 기준.</summary>
    public double RotationAngleDeg { get; set; }

    // ── 레이블 ───────────────────────────────────────────────────────────────

    /// <summary>도형 위에 표시할 레이블 텍스트. null / 빈 문자열이면 레이블 없음.</summary>
    public string? LabelText { get; set; }

    /// <summary>레이블 글자 크기 (pt).</summary>
    public double LabelFontSizePt { get; set; } = 10;

    /// <summary>레이블 색상 hex. null 이면 자동 (채우기 색이 어두우면 흰색, 밝으면 검정).</summary>
    public string? LabelColor { get; set; }

    /// <summary>레이블 글꼴 이름. null 이면 문서 기본 글꼴.</summary>
    public string? LabelFontFamily { get; set; }

    public bool LabelBold { get; set; }
    public bool LabelItalic { get; set; }

    /// <summary>레이블 가로 정렬. 기본 Center.</summary>
    public ShapeLabelHAlign LabelHAlign { get; set; } = ShapeLabelHAlign.Center;

    /// <summary>레이블 세로 정렬. 기본 Middle.</summary>
    public ShapeLabelVAlign LabelVAlign { get; set; } = ShapeLabelVAlign.Middle;

    /// <summary>레이블 배경색 hex. null / 빈 문자열이면 배경 없음(투명).</summary>
    public string? LabelBackgroundColor { get; set; }

    /// <summary>레이블 가로 위치 오프셋 (mm). 정렬 기준점에서 양수면 오른쪽, 음수면 왼쪽으로 이동.</summary>
    public double LabelOffsetXMm { get; set; }

    /// <summary>레이블 세로 위치 오프셋 (mm). 정렬 기준점에서 양수면 아래, 음수면 위로 이동.</summary>
    public double LabelOffsetYMm { get; set; }

    // ── 여백 ─────────────────────────────────────────────────────────────────

    /// <summary>위 여백 (mm). Inline/WrapLeft/WrapRight 에서 본문과의 간격.</summary>
    public double MarginTopMm { get; set; }

    /// <summary>아래 여백 (mm).</summary>
    public double MarginBottomMm { get; set; }
}

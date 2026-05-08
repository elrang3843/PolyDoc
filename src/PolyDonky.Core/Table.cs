namespace PolyDonky.Core;

using System.Text.Json.Serialization;

/// <summary>셀/표 테두리 한 면의 스타일.</summary>
public readonly struct CellBorderSide
{
    public double ThicknessPt { get; init; }
    public string? Color { get; init; }

    public CellBorderSide(double thicknessPt, string? color = null)
    {
        ThicknessPt = thicknessPt;
        Color = color;
    }
}

/// <summary>표 수평 정렬 (페이지 기준).</summary>
public enum TableHAlign { Left, Center, Right }

/// <summary>표와 본문 텍스트의 배치 관계.</summary>
public enum TableWrapMode
{
    /// <summary>본문 흐름 — 표가 단독 블록을 차지 (기본값). HAlign 으로 좌/가/우 정렬.</summary>
    Block,
    /// <summary>텍스트 위로 겹치기 (절대 위치, OverlayXMm/OverlayYMm 사용).</summary>
    InFrontOfText,
    /// <summary>텍스트 아래로 겹치기 (절대 위치, OverlayXMm/OverlayYMm 사용).</summary>
    BehindText,
    /// <summary>페이지 기준 위치 고정 — 본문 텍스트 흐름과 독립 (절대 위치).</summary>
    Fixed,
}

/// <summary>셀 텍스트 수평 정렬.</summary>
public enum CellTextAlign { Left, Center, Right, Justify }

/// <summary>
/// 행·열 구조의 표. 셀은 임의의 Block(주로 Paragraph) 들을 포함한다.
/// 셀 병합은 <see cref="TableCell.RowSpan"/> / <see cref="TableCell.ColumnSpan"/> 로 표현하고,
/// 병합으로 사라진 자리에는 cell 을 두지 않는다 (HWPX/DOCX 양쪽과 호환되는 sparse 표현).
/// </summary>
public sealed class Table : Block, IOverlayAnchored
{
    public IList<TableColumn> Columns { get; set; } = new List<TableColumn>();
    public IList<TableRow> Rows { get; set; } = new List<TableRow>();

    /// <summary>본문 배치 방식 (Block 이면 HAlign 적용, 나머지는 OverlayXMm/OverlayYMm 사용).</summary>
    public TableWrapMode WrapMode { get; set; } = TableWrapMode.Block;

    /// <summary>페이지 기준 수평 정렬 (WrapMode=Block 일 때만 적용).</summary>
    public TableHAlign HAlign { get; set; } = TableHAlign.Left;

    // ── 오버레이 위치 (InFrontOfText/BehindText/Fixed) ───────────────
    /// <summary>오버레이 anchoring — 0-based 페이지 인덱스 (0 = 첫 페이지).</summary>
    public int AnchorPageIndex { get; set; }
    /// <summary>오버레이 모드 X 위치 (mm, **해당 페이지 좌상단 기준**).</summary>
    public double OverlayXMm { get; set; }
    /// <summary>오버레이 모드 Y 위치 (mm, **해당 페이지 좌상단 기준**).</summary>
    public double OverlayYMm { get; set; }

    // ── 표 배경색 ────────────────────────────────────────────────────────
    /// <summary>표 전체 배경색 hex. null 이면 투명.</summary>
    public string? BackgroundColor { get; set; }

    /// <summary>표 캡션 텍스트 (HTML <c>&lt;caption&gt;</c>, DOCX table title 에 대응).
    /// null 이면 캡션 없음. 렌더러는 표 위에 가운데 정렬 작은 단락으로 표시한다.</summary>
    public string? Caption { get; set; }

    // ── 페이지 분할 옵션 ─────────────────────────────────────────────────────
    /// <summary>행 방향으로 페이지를 넘을 때 헤더 행(IsHeader=true) 을 각 조각 상단에 반복할지 여부.</summary>
    public bool RepeatHeaderRowsOnBreak { get; set; } = true;
    /// <summary>열 방향 분할 시 좌측에 반복할 헤더 열 수 (0 = 반복 없음). 열 분할은 추후 구현.</summary>
    public int HeaderColumnCount { get; set; }

    // ── 기본 셀 안여백 (mm). 0 이하면 렌더러 기본값(상하 1.0, 좌우 1.5) ──
    public double DefaultCellPaddingTopMm    { get; set; }
    public double DefaultCellPaddingBottomMm { get; set; }
    public double DefaultCellPaddingLeftMm   { get; set; }
    public double DefaultCellPaddingRightMm  { get; set; }

    // ── 표 바깥 여백 (mm). 0 이하면 여백 없음 ───────────────────────────
    public double OuterMarginTopMm    { get; set; }
    public double OuterMarginBottomMm { get; set; }
    public double OuterMarginLeftMm   { get; set; }
    public double OuterMarginRightMm  { get; set; }

    // ── 표 외곽선 ────────────────────────────────────────────────────────
    /// <summary>표 외곽선 두께 (pt). 0 이하면 외곽선 없음.</summary>
    public double BorderThicknessPt { get; set; }
    /// <summary>표 외곽선 색상 hex. null / 빈 문자열이면 기본 연회색 (#C8C8C8).</summary>
    public string? BorderColor { get; set; }
}

public sealed class TableColumn
{
    /// <summary>0 이하면 자동 너비.</summary>
    public double WidthMm { get; set; }
}

public sealed class TableRow
{
    public IList<TableCell> Cells { get; set; } = new List<TableCell>();
    /// <summary>0 이하면 자동 높이.</summary>
    public double HeightMm { get; set; }
    /// <summary>머리글 행 여부. true 이면 렌더러가 배경색·굵기를 강조해 표시한다.</summary>
    public bool IsHeader { get; set; }
}

public sealed class TableCell
{
    public IList<Block> Blocks { get; set; } = new List<Block>();
    public int RowSpan { get; set; } = 1;
    public int ColumnSpan { get; set; } = 1;
    /// <summary>0 이하면 컬럼 정의 너비를 따른다.</summary>
    public double WidthMm { get; set; }

    // ── 텍스트 정렬 ──────────────────────────────────────────────────────
    public CellTextAlign TextAlign { get; set; } = CellTextAlign.Left;

    // ── 여백 (mm). 0 이하면 렌더러 기본값(상하 1.0, 좌우 1.5) 사용 ───────
    public double PaddingTopMm    { get; set; }
    public double PaddingBottomMm { get; set; }
    public double PaddingLeftMm   { get; set; }
    public double PaddingRightMm  { get; set; }

    // ── 테두리 ───────────────────────────────────────────────────────────
    /// <summary>테두리 두께 (pt). 0 이하면 렌더러 기본값 (0.75pt) 사용. 면별 지정이 없을 때 공통값.</summary>
    public double BorderThicknessPt { get; set; }
    /// <summary>테두리 색상 hex. null / 빈 문자열이면 기본 연회색 (#C8C8C8). 면별 지정이 없을 때 공통값.</summary>
    public string? BorderColor { get; set; }

    /// <summary>위쪽 테두리. null 이면 BorderThicknessPt/BorderColor 공통값 사용.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public CellBorderSide? BorderTop { get; set; }
    /// <summary>아래쪽 테두리. null 이면 공통값 사용.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public CellBorderSide? BorderBottom { get; set; }
    /// <summary>왼쪽 테두리. null 이면 공통값 사용.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public CellBorderSide? BorderLeft { get; set; }
    /// <summary>오른쪽 테두리. null 이면 공통값 사용.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public CellBorderSide? BorderRight { get; set; }

    // ── 배경색 ───────────────────────────────────────────────────────────
    /// <summary>셀 배경색 hex. null / 빈 문자열이면 투명 (배경 없음).</summary>
    public string? BackgroundColor { get; set; }
}

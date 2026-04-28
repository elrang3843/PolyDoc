namespace PolyDonky.Core;

/// <summary>표 수평 정렬 (페이지 기준).</summary>
public enum TableHAlign { Left, Center, Right }

/// <summary>셀 텍스트 수평 정렬.</summary>
public enum CellTextAlign { Left, Center, Right, Justify }

/// <summary>
/// 행·열 구조의 표. 셀은 임의의 Block(주로 Paragraph) 들을 포함한다.
/// 셀 병합은 <see cref="TableCell.RowSpan"/> / <see cref="TableCell.ColumnSpan"/> 로 표현하고,
/// 병합으로 사라진 자리에는 cell 을 두지 않는다 (HWPX/DOCX 양쪽과 호환되는 sparse 표현).
/// </summary>
public sealed class Table : Block
{
    public IList<TableColumn> Columns { get; set; } = new List<TableColumn>();
    public IList<TableRow> Rows { get; set; } = new List<TableRow>();

    /// <summary>페이지 기준 수평 정렬.</summary>
    public TableHAlign HAlign { get; set; } = TableHAlign.Left;

    // ── 표 배경색 ────────────────────────────────────────────────────────
    /// <summary>표 전체 배경색 hex. null 이면 투명.</summary>
    public string? BackgroundColor { get; set; }

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
    /// <summary>테두리 두께 (pt). 0 이하면 렌더러 기본값 (0.75pt) 사용.</summary>
    public double BorderThicknessPt { get; set; }
    /// <summary>테두리 색상 hex. null / 빈 문자열이면 기본 연회색 (#C8C8C8).</summary>
    public string? BorderColor { get; set; }

    // ── 배경색 ───────────────────────────────────────────────────────────
    /// <summary>셀 배경색 hex. null / 빈 문자열이면 투명 (배경 없음).</summary>
    public string? BackgroundColor { get; set; }
}

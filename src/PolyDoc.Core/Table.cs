namespace PolyDoc.Core;

/// <summary>
/// 행·열 구조의 표. 셀은 임의의 Block(주로 Paragraph) 들을 포함한다.
/// 셀 병합은 <see cref="TableCell.RowSpan"/> / <see cref="TableCell.ColumnSpan"/> 로 표현하고,
/// 병합으로 사라진 자리에는 cell 을 두지 않는다 (HWPX/DOCX 양쪽과 호환되는 sparse 표현).
/// </summary>
public sealed class Table : Block
{
    public IList<TableColumn> Columns { get; set; } = new List<TableColumn>();
    public IList<TableRow> Rows { get; set; } = new List<TableRow>();
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
}

public sealed class TableCell
{
    public IList<Block> Blocks { get; set; } = new List<Block>();
    public int RowSpan { get; set; } = 1;
    public int ColumnSpan { get; set; } = 1;
    /// <summary>0 이하면 컬럼 정의 너비를 따른다.</summary>
    public double WidthMm { get; set; }
}

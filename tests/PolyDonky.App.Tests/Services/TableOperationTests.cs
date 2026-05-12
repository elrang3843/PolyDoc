using PolyDonky.App.Services;
using PolyDonky.Core;
using Xunit;

namespace PolyDonky.App.Tests.Services;

public class TableOperationTests
{
    // ── Helper ────────────────────────────────────────────────────────────

    private static Table CreateSimpleTable(int rows, int cols)
    {
        var table = new Table();
        for (int c = 0; c < cols; c++)
            table.Columns.Add(new TableColumn());

        for (int r = 0; r < rows; r++)
        {
            var row = new TableRow();
            for (int c = 0; c < cols; c++)
                row.Cells.Add(new TableCell { Blocks = new List<Block> { new Paragraph() } });
            table.Rows.Add(row);
        }
        return table;
    }

    // ── TableOperationHelpers Tests ────────────────────────────────────────

    [Fact]
    public void GetActualColumnCount_ReturnsCorrectCount()
    {
        var table = CreateSimpleTable(2, 3);
        Assert.Equal(3, TableOperationHelpers.GetActualColumnCount(table));
    }

    [Fact]
    public void GetActualRowCount_ReturnsCorrectCount()
    {
        var table = CreateSimpleTable(5, 3);
        Assert.Equal(5, TableOperationHelpers.GetActualRowCount(table));
    }

    [Fact]
    public void GetLogicalColumnCount_WithoutMerge_ReturnsSame()
    {
        var table = CreateSimpleTable(3, 4);
        Assert.Equal(4, TableOperationHelpers.GetLogicalColumnCount(table));
    }

    [Fact]
    public void GetLogicalColumnCount_WithMerge_ReturnsMax()
    {
        var table = CreateSimpleTable(2, 3);
        // 첫 행: 3열 (colspan 1 + 1 + 1)
        // 둘째 행: 2열 (colspan 2 + 1)
        table.Rows[1].Cells[0].ColumnSpan = 2;
        table.Rows[1].Cells.RemoveAt(1);

        Assert.Equal(3, TableOperationHelpers.GetLogicalColumnCount(table));
    }

    [Fact]
    public void GetCellAt_ReturnsCorrectCell()
    {
        var table = CreateSimpleTable(2, 3);
        var cell = TableOperationHelpers.GetCellAt(table, 0, 1);
        Assert.NotNull(cell);
        Assert.Same(table.Rows[0].Cells[1], cell);
    }

    [Fact]
    public void GetCellAt_OutOfRange_ReturnsNull()
    {
        var table = CreateSimpleTable(2, 3);
        Assert.Null(TableOperationHelpers.GetCellAt(table, 5, 0));
        Assert.Null(TableOperationHelpers.GetCellAt(table, 0, 10));
    }

    [Fact]
    public void FindCellByLogicalPosition_WithoutMerge_ReturnsCell()
    {
        var table = CreateSimpleTable(2, 3);
        var cell = TableOperationHelpers.FindCellByLogicalPosition(table, 0, 1);
        Assert.NotNull(cell);
        Assert.Same(table.Rows[0].Cells[1], cell);
    }

    [Fact]
    public void FindCellByLogicalPosition_WithColspan_ReturnsMergedCell()
    {
        var table = CreateSimpleTable(2, 3);
        // 첫 행 첫 번째 셀을 colspan=2로 설정
        table.Rows[0].Cells[0].ColumnSpan = 2;

        // 논리적 위치 (0,0)과 (0,1)은 같은 셀 반환
        var cell0 = TableOperationHelpers.FindCellByLogicalPosition(table, 0, 0);
        var cell1 = TableOperationHelpers.FindCellByLogicalPosition(table, 0, 1);
        Assert.Same(cell0, cell1);
        Assert.Same(table.Rows[0].Cells[0], cell0);
    }

    [Fact]
    public void GetCellsInRow_ReturnsAllCells()
    {
        var table = CreateSimpleTable(3, 4);
        var cells = TableOperationHelpers.GetCellsInRow(table, 1);
        Assert.Equal(4, cells.Count);
    }

    [Fact]
    public void ValidateTableStructure_ValidTable_ReturnsTrue()
    {
        var table = CreateSimpleTable(3, 3);
        Assert.True(TableOperationHelpers.ValidateTableStructure(table));
    }

    // ── TableModelEditor Tests ─────────────────────────────────────────────

    [Fact]
    public void InsertRowAbove_InsertsNewRow()
    {
        var table = CreateSimpleTable(2, 3);
        int originalRows = table.Rows.Count;

        TableModelEditor.InsertRowAbove(table, 0);

        Assert.Equal(originalRows + 1, table.Rows.Count);
        Assert.Equal(3, table.Rows[0].Cells.Count); // 새 행도 3개 셀 가져야 함
    }

    [Fact]
    public void InsertRowBelow_InsertsNewRow()
    {
        var table = CreateSimpleTable(2, 3);
        int originalRows = table.Rows.Count;

        TableModelEditor.InsertRowBelow(table, 1);

        Assert.Equal(originalRows + 1, table.Rows.Count);
    }

    [Fact]
    public void InsertRowAbove_WithRowspan_AdjustsRowspan()
    {
        var table = CreateSimpleTable(3, 2);
        // 첫 번째 셀의 rowspan을 2로 설정
        table.Rows[0].Cells[0].RowSpan = 2;

        TableModelEditor.InsertRowAbove(table, 1);

        // rowspan이 2 → 3으로 증가해야 함
        Assert.Equal(3, table.Rows[0].Cells[0].RowSpan);
    }

    [Fact]
    public void DeleteRows_RemovesRow()
    {
        var table = CreateSimpleTable(3, 3);
        int originalRows = table.Rows.Count;

        TableModelEditor.DeleteRows(table, 1);

        Assert.Equal(originalRows - 1, table.Rows.Count);
    }

    [Fact]
    public void DeleteRows_MultipleRows_RemovesAll()
    {
        var table = CreateSimpleTable(5, 3);
        TableModelEditor.DeleteRows(table, 1, 3);

        Assert.Equal(3, table.Rows.Count);
    }

    [Fact]
    public void InsertColumnLeft_AddsNewColumn()
    {
        var table = CreateSimpleTable(2, 3);
        int originalCols = table.Columns.Count;

        TableModelEditor.InsertColumnLeft(table, 0);

        Assert.Equal(originalCols + 1, table.Columns.Count);
        // 각 행도 새 셀이 추가되었는가?
        foreach (var row in table.Rows)
            Assert.Equal(originalCols + 1, row.Cells.Count);
    }

    [Fact]
    public void InsertColumnRight_AddsNewColumn()
    {
        var table = CreateSimpleTable(2, 3);
        int originalCols = table.Columns.Count;

        TableModelEditor.InsertColumnRight(table, 2);

        Assert.Equal(originalCols + 1, table.Columns.Count);
    }

    [Fact]
    public void DeleteColumns_RemovesColumn()
    {
        var table = CreateSimpleTable(2, 4);
        TableModelEditor.DeleteColumns(table, 1);

        Assert.Equal(3, table.Columns.Count);
        foreach (var row in table.Rows)
            Assert.Equal(3, row.Cells.Count);
    }

    [Fact]
    public void ResizeRow_SetsHeight()
    {
        var table = CreateSimpleTable(2, 3);
        TableModelEditor.ResizeRow(table, 0, 50.0);

        Assert.Equal(50.0, table.Rows[0].HeightMm);
    }

    [Fact]
    public void ResizeColumn_SetsWidth()
    {
        var table = CreateSimpleTable(2, 3);
        TableModelEditor.ResizeColumn(table, 1, 75.0);

        Assert.Equal(75.0, table.Columns[1].WidthMm);
    }

    [Fact]
    public void UpdateCellBackgroundColor_SetsBgColor()
    {
        var table = CreateSimpleTable(2, 2);
        var cell = table.Rows[0].Cells[0];

        TableModelEditor.UpdateCellBackgroundColor(cell, "#FF0000");

        Assert.Equal("#FF0000", cell.BackgroundColor);
    }

    [Fact]
    public void UpdateCellTextAlign_SetsAlignment()
    {
        var table = CreateSimpleTable(2, 2);
        var cell = table.Rows[0].Cells[0];

        TableModelEditor.UpdateCellTextAlign(cell, CellTextAlign.Center);

        Assert.Equal(CellTextAlign.Center, cell.TextAlign);
    }

    [Fact]
    public void UpdateCellPadding_SetsPaddingValues()
    {
        var table = CreateSimpleTable(2, 2);
        var cell = table.Rows[0].Cells[0];

        TableModelEditor.UpdateCellPadding(cell, top: 2.0, bottom: 3.0);

        Assert.Equal(2.0, cell.PaddingTopMm);
        Assert.Equal(3.0, cell.PaddingBottomMm);
    }

    [Fact]
    public void UpdateMultipleCellsBackgroundColor_UpdatesAllCells()
    {
        var table = CreateSimpleTable(2, 2);
        var cells = new[] { table.Rows[0].Cells[0], table.Rows[0].Cells[1] };

        TableModelEditor.UpdateMultipleCellsBackgroundColor(cells, "#00FF00");

        Assert.Equal("#00FF00", table.Rows[0].Cells[0].BackgroundColor);
        Assert.Equal("#00FF00", table.Rows[0].Cells[1].BackgroundColor);
    }

    [Fact]
    public void MergeCells_MergesSelectedCells()
    {
        var table = CreateSimpleTable(2, 3);
        var cellsToMerge = new[] { table.Rows[0].Cells[0], table.Rows[0].Cells[1] };

        TableModelEditor.MergeCells(table, cellsToMerge);

        // 첫 번째 셀의 colspan이 2가 되어야 함
        Assert.Equal(2, table.Rows[0].Cells[0].ColumnSpan);
        // 두 번째 셀이 제거되었는가?
        Assert.Equal(2, table.Rows[0].Cells.Count); // 원래 3개 → 2개
    }

    [Fact]
    public void UnmergeCells_SplitsMergedCell()
    {
        var table = CreateSimpleTable(2, 3);
        var mergedCell = table.Rows[0].Cells[0];
        mergedCell.ColumnSpan = 2;

        TableModelEditor.UnmergeCells(table, mergedCell);

        // 첫 번째 셀의 colspan이 다시 1이 되어야 함
        Assert.Equal(1, mergedCell.ColumnSpan);
    }

    // ── Integration Tests ──────────────────────────────────────────────────

    [Fact]
    public void ComplexScenario_InsertDeleteMerge()
    {
        var table = CreateSimpleTable(3, 3);

        // 1. 행 추가
        TableModelEditor.InsertRowAbove(table, 1);
        Assert.Equal(4, table.Rows.Count);

        // 2. 열 추가
        TableModelEditor.InsertColumnRight(table, 1);
        Assert.Equal(4, table.Columns.Count);

        // 3. 셀 속성 변경
        var cell = table.Rows[0].Cells[0];
        TableModelEditor.UpdateCellBackgroundColor(cell, "#FF00FF");
        TableModelEditor.UpdateCellTextAlign(cell, CellTextAlign.Right);

        Assert.Equal("#FF00FF", cell.BackgroundColor);
        Assert.Equal(CellTextAlign.Right, cell.TextAlign);

        // 4. 테이블 구조 유효성 검사
        Assert.True(TableOperationHelpers.ValidateTableStructure(table));
    }
}

using PolyDonky.Core;

namespace PolyDonky.App.Services;

/// <summary>표 모델의 행/열/셀을 조작하기 위한 유틸리티 메서드 모음.</summary>
public static class TableOperationHelpers
{
    /// <summary>표의 실제 물리적 열 개수 (sparse 병합 고려)를 구한다.</summary>
    public static int GetActualColumnCount(Table table)
    {
        return table.Columns.Count;
    }

    /// <summary>표의 실제 물리적 행 개수를 구한다.</summary>
    public static int GetActualRowCount(Table table)
    {
        return table.Rows.Count;
    }

    /// <summary>
    /// 표의 모든 행이 차지하는 논리적 열 개수(colspan 고려한 최대 폭)를 구한다.
    /// 예: 2행 테이블에서 첫 행은 3열, 둘째 행은 2열 → 3 반환.
    /// </summary>
    public static int GetLogicalColumnCount(Table table)
    {
        int maxCols = table.Columns.Count;
        foreach (var row in table.Rows)
        {
            int logicalCols = 0;
            foreach (var cell in row.Cells)
                logicalCols += cell.ColumnSpan;
            maxCols = Math.Max(maxCols, logicalCols);
        }
        return maxCols;
    }

    /// <summary>
    /// (rowIndex, cellIndex) 쌍으로 표의 특정 셀을 찾는다.
    /// cellIndex는 해당 행 내 물리적 셀 인덱스 (sparse 표현).
    /// </summary>
    public static TableCell? GetCellAt(Table table, int rowIndex, int cellIndex)
    {
        if (rowIndex < 0 || rowIndex >= table.Rows.Count) return null;
        var row = table.Rows[rowIndex];
        if (cellIndex < 0 || cellIndex >= row.Cells.Count) return null;
        return row.Cells[cellIndex];
    }

    /// <summary>
    /// 논리적 행/열 위치 (rowIdx, colIdx)에서 실제 셀을 찾는다.
    /// rowSpan/columnSpan을 고려해서 병합된 셀도 올바르게 찾아낸다.
    /// 예: colspan=2인 셀이 (0,0)을 차지하면 (0,0)과 (0,1) 모두 같은 셀 반환.
    /// </summary>
    public static TableCell? FindCellByLogicalPosition(Table table, int rowIdx, int colIdx)
    {
        if (rowIdx < 0 || rowIdx >= table.Rows.Count) return null;

        // 각 행을 순회하며 논리적 열 위치를 추적
        int logicalCol = 0;
        foreach (var cell in table.Rows[rowIdx].Cells)
        {
            if (logicalCol <= colIdx && colIdx < logicalCol + cell.ColumnSpan)
            {
                return cell;
            }
            logicalCol += cell.ColumnSpan;
        }

        return null;
    }

    /// <summary>
    /// 표의 특정 행(rowIndex)에 속한 모든 셀을 반환한다 (병합 고려 없이 물리적 셀만).
    /// </summary>
    public static IReadOnlyList<TableCell> GetCellsInRow(Table table, int rowIndex)
    {
        if (rowIndex < 0 || rowIndex >= table.Rows.Count)
            return new List<TableCell>();
        return table.Rows[rowIndex].Cells.AsReadOnly();
    }

    /// <summary>
    /// 표의 특정 열(colIndex)에 속한 모든 셀을 반환한다.
    /// colspan으로 인해 한 행에서 여러 물리적 열을 차지하는 셀도 포함된다.
    /// </summary>
    public static IReadOnlyList<(int rowIndex, int cellIndex, TableCell cell)> GetCellsInColumn(
        Table table, int colIndex)
    {
        var result = new List<(int, int, TableCell)>();

        for (int rowIdx = 0; rowIdx < table.Rows.Count; rowIdx++)
        {
            var row = table.Rows[rowIdx];
            int logicalCol = 0;

            for (int cellIdx = 0; cellIdx < row.Cells.Count; cellIdx++)
            {
                var cell = row.Cells[cellIdx];
                if (logicalCol <= colIndex && colIndex < logicalCol + cell.ColumnSpan)
                {
                    result.Add((rowIdx, cellIdx, cell));
                    break;
                }
                logicalCol += cell.ColumnSpan;
            }
        }

        return result.AsReadOnly();
    }

    /// <summary>
    /// 주어진 행 범위 [startRowIdx..endRowIdx]의 모든 행을 반환한다 (포함).
    /// </summary>
    public static IReadOnlyList<TableRow> GetRowRange(Table table, int startRowIdx, int endRowIdx)
    {
        startRowIdx = Math.Max(0, startRowIdx);
        endRowIdx = Math.Min(table.Rows.Count - 1, endRowIdx);

        var result = new List<TableRow>();
        for (int i = startRowIdx; i <= endRowIdx; i++)
            result.Add(table.Rows[i]);
        return result.AsReadOnly();
    }

    /// <summary>
    /// 표에서 헤더 행(IsHeader=true)인 행 목록을 반환한다.
    /// 일반적으로 맨 앞의 연속된 헤더 행들이 포함됨.
    /// </summary>
    public static IReadOnlyList<int> GetHeaderRowIndices(Table table)
    {
        var result = new List<int>();
        for (int i = 0; i < table.Rows.Count; i++)
        {
            if (table.Rows[i].IsHeader)
                result.Add(i);
            else
                break;
        }
        return result.AsReadOnly();
    }

    /// <summary>
    /// 표의 특정 셀이 다른 셀에 의해 병합되었는지 확인한다.
    /// (즉, 현재 위치가 colspan/rowspan으로 덮여 있는 여역인가?)
    /// </summary>
    public static bool IsCellCovered(Table table, int rowIdx, int colIdx)
    {
        // 논리적 위치 (rowIdx, colIdx)를 차지하는 셀 찾기
        var cell = FindCellByLogicalPosition(table, rowIdx, colIdx);
        if (cell == null) return true; // 범위 밖이면 "덮여 있음" 간주

        // 셀을 찾은 위치가 셀의 시작 위치인지 확인
        // (시작 위치가 아니면 다른 셀에 의해 덮여 있음)
        for (int r = 0; r < table.Rows.Count; r++)
        {
            int logicalCol = 0;
            foreach (var c in table.Rows[r].Cells)
            {
                if (ReferenceEquals(c, cell))
                {
                    // 이 셀이 논리적 위치 (rowIdx, colIdx)를 포함하는가?
                    return !(r <= rowIdx && rowIdx < r + c.RowSpan &&
                             logicalCol <= colIdx && colIdx < logicalCol + c.ColumnSpan);
                }
                logicalCol += c.ColumnSpan;
            }
        }
        return true;
    }

    /// <summary>
    /// 셀 범위를 선택했을 때, 해당 범위에 걸친 모든 셀(병합 고려)을 반환한다.
    /// startRowIdx, startColIdx ~ endRowIdx, endColIdx의 사각형 범위.
    /// </summary>
    public static HashSet<TableCell> GetCellsInRange(Table table, int startRowIdx, int startColIdx,
        int endRowIdx, int endColIdx)
    {
        var result = new HashSet<TableCell>();

        startRowIdx = Math.Max(0, startRowIdx);
        endRowIdx = Math.Min(table.Rows.Count - 1, endRowIdx);
        startColIdx = Math.Max(0, startColIdx);
        endColIdx = Math.Min(GetLogicalColumnCount(table) - 1, endColIdx);

        for (int r = startRowIdx; r <= endRowIdx; r++)
        {
            for (int c = startColIdx; c <= endColIdx; c++)
            {
                var cell = FindCellByLogicalPosition(table, r, c);
                if (cell != null)
                    result.Add(cell);
            }
        }

        return result;
    }

    /// <summary>
    /// 주어진 셀들의 행 인덱스 범위 (최소~최대)를 구한다.
    /// cells가 비어 있으면 (0, 0) 반환.
    /// </summary>
    public static (int minRow, int maxRow) GetRowRangeOfCells(Table table, IEnumerable<TableCell> cells)
    {
        int minRow = int.MaxValue;
        int maxRow = int.MinValue;

        var cellSet = new HashSet<TableCell>(cells);

        for (int r = 0; r < table.Rows.Count; r++)
        {
            foreach (var cell in table.Rows[r].Cells)
            {
                if (cellSet.Contains(cell))
                {
                    minRow = Math.Min(minRow, r);
                    maxRow = Math.Max(maxRow, r + cell.RowSpan - 1);
                }
            }
        }

        if (minRow == int.MaxValue)
            return (0, 0);
        return (minRow, maxRow);
    }

    /// <summary>
    /// 주어진 셀들의 열 인덱스 범위 (최소~최대, 논리적)를 구한다.
    /// cells가 비어 있으면 (0, 0) 반환.
    /// </summary>
    public static (int minCol, int maxCol) GetColumnRangeOfCells(Table table, IEnumerable<TableCell> cells)
    {
        int minCol = int.MaxValue;
        int maxCol = int.MinValue;

        var cellSet = new HashSet<TableCell>(cells);

        for (int r = 0; r < table.Rows.Count; r++)
        {
            int logicalCol = 0;
            foreach (var cell in table.Rows[r].Cells)
            {
                if (cellSet.Contains(cell))
                {
                    minCol = Math.Min(minCol, logicalCol);
                    maxCol = Math.Max(maxCol, logicalCol + cell.ColumnSpan - 1);
                }
                logicalCol += cell.ColumnSpan;
            }
        }

        if (minCol == int.MaxValue)
            return (0, 0);
        return (minCol, maxCol);
    }

    /// <summary>
    /// 표의 구조를 검증한다 (디버깅·테스트용).
    /// rowSpan/columnSpan이 범위를 벗어나거나 모순이 있으면 false 반환.
    /// </summary>
    public static bool ValidateTableStructure(Table table)
    {
        int colCount = GetLogicalColumnCount(table);

        for (int r = 0; r < table.Rows.Count; r++)
        {
            int logicalCol = 0;
            foreach (var cell in table.Rows[r].Cells)
            {
                if (cell.RowSpan < 1 || cell.ColumnSpan < 1)
                    return false;
                if (r + cell.RowSpan > table.Rows.Count)
                    return false;
                if (logicalCol + cell.ColumnSpan > colCount)
                    return false;
                logicalCol += cell.ColumnSpan;
            }
        }

        return true;
    }
}

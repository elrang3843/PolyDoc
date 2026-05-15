using PolyDonky.Core;

namespace PolyDonky.App.Services;

/// <summary>표 모델(Core.Table)을 직접 수정하는 명령어들.</summary>
public static class TableModelEditor
{
    /// <summary>
    /// 지정한 행 위에 새 행을 삽입한다.
    /// 새 행은 기존 행과 같은 열 개수를 가지며, 각 셀은 빈 Paragraph를 포함한다.
    /// </summary>
    public static void InsertRowAbove(Table table, int rowIndex)
    {
        if (rowIndex < 0 || rowIndex > table.Rows.Count)
            throw new ArgumentOutOfRangeException(nameof(rowIndex));

        var newRow = CreateEmptyRow(table);
        table.Rows.Insert(rowIndex, newRow);

        // 기존 rowspan > 1 인 셀 중, 삽입 위치 이상인 셀의 rowspan 증가
        for (int r = 0; r < table.Rows.Count; r++)
        {
            if (r == rowIndex) continue; // 새로 삽입한 행은 건너뛰기
            foreach (var cell in table.Rows[r].Cells)
            {
                // 셀이 삽입 위치를 걸쳐 있는 경우만 rowspan 증가
                if (r < rowIndex && r + cell.RowSpan > rowIndex)
                    cell.RowSpan++;
            }
        }
    }

    /// <summary>
    /// 지정한 행 아래에 새 행을 삽입한다.
    /// </summary>
    public static void InsertRowBelow(Table table, int rowIndex)
    {
        InsertRowAbove(table, rowIndex + 1);
    }

    /// <summary>
    /// 지정한 행들을 삭제한다.
    /// rowIndices는 정렬되지 않아도 자동 정렬 후 처리.
    /// </summary>
    public static void DeleteRows(Table table, params int[] rowIndices)
    {
        if (rowIndices.Length == 0) return;

        var sortedIndices = rowIndices.Distinct().OrderByDescending(i => i).ToList();

        foreach (int idx in sortedIndices)
        {
            if (idx < 0 || idx >= table.Rows.Count)
                throw new ArgumentOutOfRangeException($"Row index {idx} out of range");

            int heightToRemove = (int)table.Rows[idx].HeightMm;
            table.Rows.RemoveAt(idx);

            // 남은 행들의 rowspan 조정
            // 삭제된 행의 범위: [idx, idx + rowspan)
            // 이를 넘어가는 셀들의 rowspan 감소 필요
            for (int r = 0; r < table.Rows.Count; r++)
            {
                foreach (var cell in table.Rows[r].Cells)
                {
                    int cellStartRow = GetCellStartRow(table, r, cell);
                    int cellEndRow = cellStartRow + cell.RowSpan;

                    // 삭제된 행들이 셀 범위와 겹치는 경우
                    if (cellStartRow < idx && cellEndRow > idx)
                    {
                        int overlap = Math.Min(cellEndRow, idx + 1) - Math.Max(cellStartRow, idx);
                        cell.RowSpan = Math.Max(1, cell.RowSpan - overlap);
                    }
                }
            }
        }
    }

    /// <summary>
    /// 지정한 열 왼쪽에 새 열을 삽입한다.
    /// 모든 행에 새 셀이 추가되며, colspan 조정이 필요한 경우 처리된다.
    /// </summary>
    public static void InsertColumnLeft(Table table, int colIndex)
    {
        InsertColumnHelper(table, colIndex, insertBefore: true);
    }

    /// <summary>
    /// 지정한 열 오른쪽에 새 열을 삽입한다.
    /// </summary>
    public static void InsertColumnRight(Table table, int colIndex)
    {
        InsertColumnHelper(table, colIndex, insertBefore: false);
    }

    private static void InsertColumnHelper(Table table, int logicalColIndex, bool insertBefore)
    {
        if (logicalColIndex < 0 || logicalColIndex > TableOperationHelpers.GetLogicalColumnCount(table))
            throw new ArgumentOutOfRangeException(nameof(logicalColIndex));

        int insertLogicalCol = insertBefore ? logicalColIndex : logicalColIndex + 1;
        int actualColIndex = insertBefore ? logicalColIndex : logicalColIndex + 1;

        // 새 TableColumn 추가
        if (actualColIndex <= table.Columns.Count)
            table.Columns.Insert(actualColIndex, new TableColumn());
        else
            table.Columns.Add(new TableColumn());

        // 각 행에서 셀 처리
        for (int r = 0; r < table.Rows.Count; r++)
        {
            var row = table.Rows[r];
            int logicalCol = 0;
            int insertCellIndex = row.Cells.Count; // 기본값: 행 끝에 삽입

            // 삽입 위치를 찾기
            for (int cellIdx = 0; cellIdx < row.Cells.Count; cellIdx++)
            {
                var cell = row.Cells[cellIdx];
                if (logicalCol == insertLogicalCol)
                {
                    insertCellIndex = cellIdx;
                    break;
                }
                if (logicalCol < insertLogicalCol && logicalCol + cell.ColumnSpan > insertLogicalCol)
                {
                    // 삽입 위치가 이 셀의 colspan 범위 내 → colspan 증가
                    cell.ColumnSpan++;
                    insertCellIndex = row.Cells.Count; // 이 행에는 새 셀 추가 안 함
                    break;
                }
                logicalCol += cell.ColumnSpan;
            }

            // 새 셀 삽입 (colspan 범위에 걸리지 않은 경우만)
            if (insertCellIndex < row.Cells.Count || logicalCol == insertLogicalCol)
            {
                var newCell = new TableCell
                {
                    Blocks = new List<Block> { new Paragraph() },
                    RowSpan = 1,
                    ColumnSpan = 1
                };
                // 표의 기본 테두리 속성 상속
                if (table.BorderThicknessPt > 0)
                    newCell.BorderThicknessPt = table.BorderThicknessPt;
                if (!string.IsNullOrEmpty(table.BorderColor))
                    newCell.BorderColor = table.BorderColor;
                row.Cells.Insert(insertCellIndex, newCell);
            }
        }
    }

    /// <summary>
    /// 지정한 열들을 삭제한다.
    /// colIndices는 논리적 열 인덱스.
    /// </summary>
    public static void DeleteColumns(Table table, params int[] colIndices)
    {
        if (colIndices.Length == 0) return;

        var sortedIndices = colIndices.Distinct().OrderByDescending(i => i).ToList();
        int maxCol = TableOperationHelpers.GetLogicalColumnCount(table);

        foreach (int logicalCol in sortedIndices)
        {
            if (logicalCol < 0 || logicalCol >= maxCol)
                throw new ArgumentOutOfRangeException($"Column index {logicalCol} out of range");

            // 각 행에서 해당 열의 셀 제거 또는 colspan 조정
            for (int r = 0; r < table.Rows.Count; r++)
            {
                var row = table.Rows[r];
                int currentLogicalCol = 0;

                for (int cellIdx = row.Cells.Count - 1; cellIdx >= 0; cellIdx--)
                {
                    var cell = row.Cells[cellIdx];
                    int cellStart = currentLogicalCol;
                    int cellEnd = currentLogicalCol + cell.ColumnSpan;

                    if (cellStart == logicalCol && cell.ColumnSpan == 1)
                    {
                        // 정확히 이 열만 차지 → 셀 제거
                        row.Cells.RemoveAt(cellIdx);
                    }
                    else if (cellStart < logicalCol && cellEnd > logicalCol)
                    {
                        // colspan이 삭제 열을 걸쳐 있음 → colspan 감소
                        cell.ColumnSpan--;
                    }

                    currentLogicalCol += cell.ColumnSpan;
                }
            }

            // 실제 TableColumn도 제거 (삽입 시와 대칭)
            if (logicalCol < table.Columns.Count)
                table.Columns.RemoveAt(logicalCol);
        }
    }

    /// <summary>
    /// 특정 행의 높이를 설정한다 (mm 단위).
    /// heightMm <= 0 이면 자동 높이.
    /// </summary>
    public static void ResizeRow(Table table, int rowIndex, double heightMm)
    {
        if (rowIndex < 0 || rowIndex >= table.Rows.Count)
            throw new ArgumentOutOfRangeException(nameof(rowIndex));
        table.Rows[rowIndex].HeightMm = heightMm;
    }

    /// <summary>
    /// 특정 열의 너비를 설정한다 (mm 단위).
    /// widthMm <= 0 이면 자동 너비.
    /// </summary>
    public static void ResizeColumn(Table table, int colIndex, double widthMm)
    {
        if (colIndex < 0 || colIndex >= table.Columns.Count)
            throw new ArgumentOutOfRangeException(nameof(colIndex));
        table.Columns[colIndex].WidthMm = widthMm;
    }

    /// <summary>
    /// 셀의 배경색을 설정한다 (hex 색상값, null이면 투명).
    /// </summary>
    public static void UpdateCellBackgroundColor(TableCell cell, string? colorHex)
    {
        cell.BackgroundColor = colorHex;
    }

    /// <summary>
    /// 셀의 텍스트 정렬을 설정한다.
    /// </summary>
    public static void UpdateCellTextAlign(TableCell cell, CellTextAlign align)
    {
        cell.TextAlign = align;
    }

    /// <summary>
    /// 셀의 여백을 설정한다 (mm 단위).
    /// 값이 0 이하면 렌더러 기본값 사용.
    /// </summary>
    public static void UpdateCellPadding(TableCell cell, double? top = null, double? bottom = null,
        double? left = null, double? right = null)
    {
        if (top.HasValue) cell.PaddingTopMm = top.Value;
        if (bottom.HasValue) cell.PaddingBottomMm = bottom.Value;
        if (left.HasValue) cell.PaddingLeftMm = left.Value;
        if (right.HasValue) cell.PaddingRightMm = right.Value;
    }

    /// <summary>
    /// 셀의 테두리를 설정한다.
    /// null이 전달되면 기본값(BorderThicknessPt/BorderColor)을 사용.
    /// </summary>
    public static void UpdateCellBorder(TableCell cell, CellBorderSide? top = null,
        CellBorderSide? bottom = null, CellBorderSide? left = null, CellBorderSide? right = null)
    {
        if (top is not null) cell.BorderTop = top;
        if (bottom is not null) cell.BorderBottom = bottom;
        if (left is not null) cell.BorderLeft = left;
        if (right is not null) cell.BorderRight = right;
    }

    /// <summary>
    /// 셀의 테두리 두께와 색상 공통값을 설정한다.
    /// </summary>
    public static void UpdateCellBorderDefaults(TableCell cell, double thicknessPt, string? colorHex)
    {
        cell.BorderThicknessPt = thicknessPt;
        cell.BorderColor = colorHex;
    }

    /// <summary>
    /// 여러 셀의 배경색을 동시에 설정한다.
    /// </summary>
    public static void UpdateMultipleCellsBackgroundColor(IEnumerable<TableCell> cells, string? colorHex)
    {
        foreach (var cell in cells)
            cell.BackgroundColor = colorHex;
    }

    /// <summary>
    /// 여러 셀의 텍스트 정렬을 동시에 설정한다.
    /// </summary>
    public static void UpdateMultipleCellsTextAlign(IEnumerable<TableCell> cells, CellTextAlign align)
    {
        foreach (var cell in cells)
            cell.TextAlign = align;
    }

    /// <summary>
    /// 주어진 셀들을 병합한다.
    /// 모든 셀이 같은 행/열 범위에 있어야 하고, 사각형 모양이어야 함.
    /// </summary>
    public static void MergeCells(Table table, IEnumerable<TableCell> cellsToMerge)
    {
        var cells = cellsToMerge.ToList();
        if (cells.Count < 2)
            throw new InvalidOperationException("병합할 셀은 최소 2개 이상이어야 합니다.");

        var (minRow, maxRow) = TableOperationHelpers.GetRowRangeOfCells(table, cells);
        var (minCol, maxCol) = TableOperationHelpers.GetColumnRangeOfCells(table, cells);

        // 첫 번째 셀 찾기 (최상단 좌측)
        TableCell? firstCell = null;
        int firstCellRow = -1;

        for (int r = minRow; r <= maxRow; r++)
        {
            int logicalCol = 0;
            foreach (var cell in table.Rows[r].Cells)
            {
                if (logicalCol == minCol)
                {
                    firstCell = cell;
                    firstCellRow = r;
                    break;
                }
                logicalCol += cell.ColumnSpan;
            }
            if (firstCell != null) break;
        }

        if (firstCell == null)
            throw new InvalidOperationException("병합할 첫 셀을 찾을 수 없습니다.");

        // 첫 셀의 rowspan, colspan 설정
        firstCell.RowSpan = maxRow - minRow + 1;
        firstCell.ColumnSpan = maxCol - minCol + 1;

        // 다른 셀들 제거
        for (int r = minRow; r <= maxRow; r++)
        {
            var row = table.Rows[r];
            for (int cellIdx = row.Cells.Count - 1; cellIdx >= 0; cellIdx--)
            {
                var cell = row.Cells[cellIdx];
                if (!ReferenceEquals(cell, firstCell) && cells.Contains(cell))
                {
                    row.Cells.RemoveAt(cellIdx);
                }
            }
        }
    }

    /// <summary>
    /// 병합된 셀을 분리한다 (rowSpan > 1 또는 columnSpan > 1).
    /// </summary>
    public static void UnmergeCells(Table table, TableCell mergedCell)
    {
        if (mergedCell.RowSpan <= 1 && mergedCell.ColumnSpan <= 1)
            throw new InvalidOperationException("분리할 병합이 없습니다.");

        // 병합된 셀의 시작 위치 찾기
        int startRow = -1, startCol = -1;
        for (int r = 0; r < table.Rows.Count; r++)
        {
            int logicalCol = 0;
            foreach (var cell in table.Rows[r].Cells)
            {
                if (ReferenceEquals(cell, mergedCell))
                {
                    startRow = r;
                    startCol = logicalCol;
                    break;
                }
                logicalCol += cell.ColumnSpan;
            }
            if (startRow >= 0) break;
        }

        if (startRow < 0)
            throw new InvalidOperationException("셀의 시작 위치를 찾을 수 없습니다.");

        // 분리: 각 서브셀마다 새로운 셀 생성
        for (int r = startRow; r < startRow + mergedCell.RowSpan; r++)
        {
            for (int c = startCol; c < startCol + mergedCell.ColumnSpan; c++)
            {
                if (r == startRow && c == startCol)
                {
                    // 원본 셀: rowspan/colspan 1로 리셋
                    mergedCell.RowSpan = 1;
                    mergedCell.ColumnSpan = 1;
                }
                else
                {
                    // 새로운 셀 생성 및 삽입
                    var newCell = new TableCell
                    {
                        Blocks = new List<Block> { new Paragraph() },
                        RowSpan = 1,
                        ColumnSpan = 1
                    };
                    // 표의 기본 테두리 속성 상속
                    if (table.BorderThicknessPt > 0)
                        newCell.BorderThicknessPt = table.BorderThicknessPt;
                    if (!string.IsNullOrEmpty(table.BorderColor))
                        newCell.BorderColor = table.BorderColor;
                    table.Rows[r].Cells.Insert(0, newCell); // 간단하게 맨 앞에 추가
                }
            }
        }
    }

    // ── Helper Methods ────────────────────────────────────────────────────────

    /// <summary>
    /// 기존 표와 같은 열 개수를 가진 빈 행을 생성한다.
    /// </summary>
    private static TableRow CreateEmptyRow(Table table)
    {
        var row = new TableRow();
        for (int c = 0; c < table.Columns.Count; c++)
        {
            var cell = new TableCell
            {
                Blocks = new List<Block> { new Paragraph() },
                RowSpan = 1,
                ColumnSpan = 1,
                // 표의 모든 속성을 그대로 상속 (조건 없이)
                BorderThicknessPt = table.BorderThicknessPt,
                BorderColor = table.BorderColor,
                BackgroundColor = table.BackgroundColor,
            };
            row.Cells.Add(cell);
        }
        return row;
    }

    /// <summary>
    /// 표에서 특정 셀의 시작 행 인덱스를 구한다.
    /// </summary>
    private static int GetCellStartRow(Table table, int currentRow, TableCell cell)
    {
        for (int r = 0; r <= currentRow; r++)
        {
            foreach (var c in table.Rows[r].Cells)
            {
                if (ReferenceEquals(c, cell))
                {
                    // rowspan이 currentRow를 포함하는가?
                    if (r <= currentRow && currentRow < r + c.RowSpan)
                        return r;
                }
            }
        }
        return currentRow;
    }
}

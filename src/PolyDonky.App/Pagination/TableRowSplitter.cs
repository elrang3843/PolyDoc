using System.Collections.Generic;
using System.Linq;
using PolyDonky.Core;
using WpfDocs = System.Windows.Documents;

namespace PolyDonky.App.Pagination;

/// <summary>
/// 페이지 경계를 가로지르는 <see cref="Table"/> 을 행 기준으로 조각(fragment) 으로 분할한다.
/// <para>열 방향 분할은 추후 구현 예정 — 현재는 모델에 <see cref="Table.HeaderColumnCount"/> 만 보유.</para>
/// </summary>
internal static class TableRowSplitter
{
    /// <summary>
    /// <paramref name="wpfTable"/> 의 각 행이 속한 페이지를 paginator 에 질의해
    /// 페이지별 행 그룹을 반환한다. 헤더 행(IsHeader=true) 은 그룹 분류에서 제외된다.
    /// </summary>
    internal static List<(int pageNum, List<int> bodyRowIndices)> GetRowGroups(
        WpfDocs.Table wpfTable,
        Table coreTable,
        WpfDocs.DynamicDocumentPaginator paginator)
    {
        var result = new List<(int, List<int>)>();

        // Wpf.Table 은 단일 RowGroup 으로 빌드된다(BuildTable 참조).
        var wpfRows = wpfTable.RowGroups
            .SelectMany(rg => rg.Rows)
            .ToList();

        int curPage = -1;
        List<int>? curGroup = null;

        for (int i = 0; i < coreTable.Rows.Count; i++)
        {
            if (coreTable.Rows[i].IsHeader) continue; // 헤더 행은 별도 처리

            int pg;
            try
            {
                pg = i < wpfRows.Count
                    ? paginator.GetPageNumber(wpfRows[i].ContentStart)
                    : (result.Count > 0 ? result[^1].Item1 : 0);
            }
            catch
            {
                pg = result.Count > 0 ? result[^1].Item1 : 0;
            }

            if (pg != curPage)
            {
                curGroup = new List<int>();
                result.Add((pg, curGroup));
                curPage = pg;
            }
            curGroup!.Add(i);
        }

        return result;
    }

    /// <summary>
    /// 행 그룹 목록으로부터 <see cref="Table"/> 조각을 생성한다.
    /// </summary>
    /// <param name="source">원본 Core.Table.</param>
    /// <param name="rowGroups"><see cref="GetRowGroups"/> 결과.</param>
    /// <returns>
    /// (fragment, pageIdx) 쌍 목록. 단일 페이지면 조각이 1개이고
    /// 그것이 원본과 동일한 행을 담는다.
    /// </returns>
    internal static List<(Table fragment, int pageIdx)> BuildFragments(
        Table source,
        List<(int pageNum, List<int> bodyRowIndices)> rowGroups)
    {
        if (rowGroups.Count == 0)
            return new List<(Table, int)> { (source, 0) };

        if (rowGroups.Count == 1)
            return new List<(Table, int)> { (source, rowGroups[0].pageNum) };

        // 여러 페이지에 걸친 표 — 조각 생성
        var fragments = new List<(Table, int)>();

        for (int gi = 0; gi < rowGroups.Count; gi++)
        {
            bool isFirst = gi == 0;
            bool isLast  = gi == rowGroups.Count - 1;
            var (pageNum, bodyIndices) = rowGroups[gi];

            // 첫 조각: 캡션 유지. 이후 조각: 캡션 제거(반복 방지).
            bool prependHeaders = !isFirst && source.RepeatHeaderRowsOnBreak;
            var frag = CreateFragment(source, bodyIndices, prependHeaders,
                                      omitCaption: !isFirst, isLastFragment: isLast);
            fragments.Add((frag, pageNum));
        }

        return fragments;
    }

    // ── 내부 헬퍼 ────────────────────────────────────────────────────────────

    private static Table CreateFragment(
        Table source,
        IReadOnlyList<int> bodyRowIndices,
        bool prependHeaders,
        bool omitCaption,
        bool isLastFragment)
    {
        var frag = new Table
        {
            Id                         = source.Id,
            Status                     = source.Status,
            WrapMode                   = source.WrapMode,
            HAlign                     = source.HAlign,
            Caption                    = omitCaption ? null : source.Caption,
            BackgroundColor            = source.BackgroundColor,
            BorderThicknessPt          = source.BorderThicknessPt,
            BorderColor                = source.BorderColor,
            DefaultCellPaddingTopMm    = source.DefaultCellPaddingTopMm,
            DefaultCellPaddingBottomMm = source.DefaultCellPaddingBottomMm,
            DefaultCellPaddingLeftMm   = source.DefaultCellPaddingLeftMm,
            DefaultCellPaddingRightMm  = source.DefaultCellPaddingRightMm,
            OuterMarginTopMm           = source.OuterMarginTopMm,
            OuterMarginBottomMm        = source.OuterMarginBottomMm,
            OuterMarginLeftMm          = source.OuterMarginLeftMm,
            OuterMarginRightMm         = source.OuterMarginRightMm,
            RepeatHeaderRowsOnBreak    = source.RepeatHeaderRowsOnBreak,
            HeaderColumnCount          = source.HeaderColumnCount,
        };

        foreach (var col in source.Columns)
            frag.Columns.Add(new TableColumn { WidthMm = col.WidthMm });

        // 첫 번째 본문 행 / 마지막 본문 행의 원본 인덱스
        // (앞머리 IsHeader vs 꼬리 IsHeader 구분에 사용)
        int firstBodySrcIdx = -1, lastBodySrcIdx = -1;
        for (int i = 0; i < source.Rows.Count; i++)
        {
            if (!source.Rows[i].IsHeader)
            {
                if (firstBodySrcIdx < 0) firstBodySrcIdx = i;
                lastBodySrcIdx = i;
            }
        }

        var bodySet     = new HashSet<int>(bodyRowIndices);
        var posLookup   = new Dictionary<int, int>(bodyRowIndices.Count);
        for (int p = 0; p < bodyRowIndices.Count; p++) posLookup[bodyRowIndices[p]] = p;

        if (prependHeaders)
        {
            // 이후 조각: 앞머리 IsHeader(첫 본문 행 이전) 만 반복, 꼬리 IsHeader 는 마지막 조각에만
            for (int i = 0; i < (firstBodySrcIdx >= 0 ? firstBodySrcIdx : source.Rows.Count); i++)
            {
                if (source.Rows[i].IsHeader)
                    frag.Rows.Add(CloneRow(source.Rows[i], maxRowSpan: null));
            }

            for (int pos = 0; pos < bodyRowIndices.Count; pos++)
            {
                var srcRow = source.Rows[bodyRowIndices[pos]];
                frag.Rows.Add(CloneRow(srcRow, bodyRowIndices.Count - pos));
            }

            if (isLastFragment && lastBodySrcIdx >= 0)
            {
                for (int i = lastBodySrcIdx + 1; i < source.Rows.Count; i++)
                {
                    if (source.Rows[i].IsHeader)
                        frag.Rows.Add(CloneRow(source.Rows[i], maxRowSpan: null));
                }
            }
        }
        else
        {
            // 첫 조각: 원본 행 순서 보존
            // 꼬리 IsHeader(마지막 본문 행 이후) 는 isLastFragment 일 때만 포함
            for (int i = 0; i < source.Rows.Count; i++)
            {
                var row = source.Rows[i];
                if (row.IsHeader)
                {
                    bool isTrailing = lastBodySrcIdx >= 0 && i > lastBodySrcIdx;
                    if (isTrailing && !isLastFragment) continue;
                    frag.Rows.Add(CloneRow(row, maxRowSpan: null));
                }
                else if (bodySet.Contains(i))
                {
                    int pos = posLookup[i];
                    frag.Rows.Add(CloneRow(row, bodyRowIndices.Count - pos));
                }
                // 이 조각에 속하지 않는 본문 행은 건너뜀
            }
        }

        return frag;
    }

    private static TableRow CloneRow(TableRow src, int? maxRowSpan)
    {
        var row = new TableRow { IsHeader = src.IsHeader, HeightMm = src.HeightMm };
        foreach (var cell in src.Cells)
        {
            var c = new TableCell
            {
                RowSpan         = maxRowSpan.HasValue ? System.Math.Min(cell.RowSpan, System.Math.Max(1, maxRowSpan.Value)) : cell.RowSpan,
                ColumnSpan      = cell.ColumnSpan,
                WidthMm         = cell.WidthMm,
                TextAlign       = cell.TextAlign,
                PaddingTopMm    = cell.PaddingTopMm,
                PaddingBottomMm = cell.PaddingBottomMm,
                PaddingLeftMm   = cell.PaddingLeftMm,
                PaddingRightMm  = cell.PaddingRightMm,
                BorderThicknessPt = cell.BorderThicknessPt,
                BorderColor       = cell.BorderColor,
                BorderTop         = cell.BorderTop,
                BorderBottom      = cell.BorderBottom,
                BorderLeft        = cell.BorderLeft,
                BorderRight       = cell.BorderRight,
                BackgroundColor   = cell.BackgroundColor,
            };
            foreach (var b in cell.Blocks) c.Blocks.Add(b);
            row.Cells.Add(c);
        }
        return row;
    }
}

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
                if (i < wpfRows.Count)
                {
                    int startPg = paginator.GetPageNumber(wpfRows[i].ContentStart);
                    pg = startPg;
                    // ContentEnd 가 ContentStart 보다 뒤 페이지에 있으면 행이 페이지를 넘친 것.
                    // 넘치는 행을 현재 페이지에 두면 RTB 가 스크롤을 발생시키므로 다음 페이지로 밀어낸다.
                    try
                    {
                        int endPg = paginator.GetPageNumber(wpfRows[i].ContentEnd);
                        if (endPg > startPg) pg = endPg;
                    }
                    catch { /* ContentEnd 조회 실패 시 startPg 유지 */ }
                }
                else
                {
                    pg = result.Count > 0 ? result[^1].Item1 : 0;
                }
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

        // Fallback: ContentEnd 체크가 실패한 경우(WPF 가 행 내용을 클립) 를 보완.
        // 모든 본문 행이 같은 페이지에 있어도 표 전체의 ContentEnd 가 다음 페이지면
        // 마지막 본문 행 하나를 다음 페이지로 밀어낸다.
        if (result.Count == 1 && result[0].Item2.Count >= 2)
        {
            try
            {
                int tblStartPg = paginator.GetPageNumber(wpfTable.ContentStart);
                int tblEndPg   = paginator.GetPageNumber(wpfTable.ContentEnd);
                if (tblEndPg > tblStartPg)
                {
                    var indices = result[0].Item2;
                    int lastBodyIdx = indices[^1];
                    indices.RemoveAt(indices.Count - 1);
                    result.Add((tblEndPg, new System.Collections.Generic.List<int> { lastBodyIdx }));
                }
            }
            catch { }
        }

        AdjustGroupsForRowSpan(result, coreTable);
        return result;
    }

    /// <summary>
    /// RowSpan 으로 병합된 셀의 기준 행과 피복 행이 서로 다른 그룹에 속하지 않도록
    /// 그룹 경계를 재조정한다. 안정될 때까지(더 이상 이동 없을 때까지) 반복한다.
    /// </summary>
    private static void AdjustGroupsForRowSpan(
        List<(int pageNum, List<int> bodyRowIndices)> groups,
        Table source)
    {
        if (groups.Count <= 1) return;

        bool changed = true;
        while (changed)
        {
            changed = false;

            for (int gi = 0; gi < groups.Count - 1; gi++)
            {
                var curIndices  = groups[gi].bodyRowIndices;
                var nextIndices = groups[gi + 1].bodyRowIndices;
                if (curIndices.Count == 0) continue;

                var nextSet = new HashSet<int>(nextIndices);
                int cutAt = -1;

                for (int k = 0; k < curIndices.Count; k++)
                {
                    int srcIdx = curIndices[k];
                    var row = source.Rows[srcIdx];

                    // 이 행의 셀 중 RowSpan > 1 인 것이 다음 그룹의 행까지 닿는지 확인
                    int maxReach = srcIdx;
                    foreach (var cell in row.Cells)
                    {
                        if (cell.RowSpan > 1)
                            maxReach = System.Math.Max(maxReach, srcIdx + cell.RowSpan - 1);
                    }
                    if (maxReach <= srcIdx) continue;

                    for (int j = srcIdx + 1; j <= maxReach && j < source.Rows.Count; j++)
                    {
                        if (!source.Rows[j].IsHeader && nextSet.Contains(j))
                        {
                            // k 번 이후 행들이 다음 그룹의 행과 병합돼 있으므로 k 부터 이동
                            if (cutAt < 0) cutAt = k;
                            break;
                        }
                    }
                }

                if (cutAt >= 0)
                {
                    int moveCount = curIndices.Count - cutAt;
                    var toMove = curIndices.GetRange(cutAt, moveCount);
                    curIndices.RemoveRange(cutAt, moveCount);
                    nextIndices.InsertRange(0, toMove);
                    changed = true;
                }
            }

            // 이동으로 비어 버린 그룹 제거
            for (int gi = groups.Count - 2; gi >= 0; gi--)
            {
                if (groups[gi].bodyRowIndices.Count == 0)
                    groups.RemoveAt(gi);
            }
        }
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

        // 여러 페이지에 걸친 표 — 조각 생성.
        // 각 조각은 완전히 독립된 표로 취급되므로 고유 Id 를 부여한다.
        // 형식: "{sourceId}§t{index}" (§t = table fragment separator)
        var fragments = new List<(Table, int)>();

        for (int gi = 0; gi < rowGroups.Count; gi++)
        {
            bool isFirst = gi == 0;
            bool isLast  = gi == rowGroups.Count - 1;
            var (pageNum, bodyIndices) = rowGroups[gi];

            bool prependHeaders = !isFirst && source.RepeatHeaderRowsOnBreak;
            var frag = CreateFragment(source, bodyIndices, prependHeaders,
                                      omitCaption: !isFirst, isLastFragment: isLast, isFirstFragment: isFirst);

            // 각 조각에 고유 Id — MergeTableFragments 없이도 독립 표로 관리된다.
            frag.Id = string.IsNullOrEmpty(source.Id)
                ? $"§t{gi}"
                : $"{source.Id}§t{gi}";

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
        bool isLastFragment,
        bool isFirstFragment = true)
    {
        var frag = new Table
        {
            Id                         = source.Id,
            Status                     = source.Status,
            WrapMode                   = source.WrapMode,
            HAlign                     = source.HAlign,
            Caption                    = omitCaption ? null : source.Caption,
            BackgroundColor            = source.BackgroundColor,
            WidthMm                    = source.WidthMm,
            HeightMm                   = 0,  // 조각은 행 내용으로 높이를 결정, 원본 전체 높이 복사 금지
            IsFlexLayout               = source.IsFlexLayout,
            BorderCollapse             = source.BorderCollapse,
            BorderThicknessPt          = source.BorderThicknessPt,
            BorderColor                = source.BorderColor,
            BorderTop                  = source.BorderTop,
            BorderBottom               = source.BorderBottom,
            BorderLeft                 = source.BorderLeft,
            BorderRight                = source.BorderRight,
            InnerBorderHorizontal      = source.InnerBorderHorizontal,
            InnerBorderVertical        = source.InnerBorderVertical,
            DefaultCellPaddingTopMm    = source.DefaultCellPaddingTopMm,
            DefaultCellPaddingBottomMm = source.DefaultCellPaddingBottomMm,
            DefaultCellPaddingLeftMm   = source.DefaultCellPaddingLeftMm,
            DefaultCellPaddingRightMm  = source.DefaultCellPaddingRightMm,
            // 후속 조각은 페이지 상단에 바로 붙으므로 위쪽 여백 제거.
            // 마지막이 아닌 조각은 아래쪽 여백도 제거 (다음 조각과 이어지는 표).
            OuterMarginTopMm           = isFirstFragment  ? source.OuterMarginTopMm    : 0,
            OuterMarginBottomMm        = isLastFragment   ? source.OuterMarginBottomMm : 0,
            OuterMarginLeftMm          = source.OuterMarginLeftMm,
            OuterMarginRightMm         = source.OuterMarginRightMm,
            RepeatHeaderRowsOnBreak    = source.RepeatHeaderRowsOnBreak,
            HeaderColumnCount          = source.HeaderColumnCount,
            AnchorPageIndex            = source.AnchorPageIndex,
            OverlayXMm                 = source.OverlayXMm,
            OverlayYMm                 = source.OverlayYMm,
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
        var row = new TableRow
        {
            IsHeader        = src.IsHeader,
            HeightMm        = src.HeightMm,
            BackgroundColor = src.BackgroundColor,
            VerticalAlign   = src.VerticalAlign,
        };
        foreach (var cell in src.Cells)
        {
            var c = new TableCell
            {
                RowSpan         = maxRowSpan.HasValue ? System.Math.Min(cell.RowSpan, System.Math.Max(1, maxRowSpan.Value)) : cell.RowSpan,
                ColumnSpan      = cell.ColumnSpan,
                WidthMm         = cell.WidthMm,
                TextAlign       = cell.TextAlign,
                VerticalAlign   = cell.VerticalAlign,
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

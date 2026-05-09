using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using PolyDonky.App.Services;
using PolyDonky.Core;

namespace PolyDonky.App.Pagination;

/// <summary>
/// <see cref="PaginatedDocument"/> 의 블록 배정 정보를 기반으로 페이지별·단별
/// <see cref="PerPageDocumentSlice"/> 를 생성한다.
/// <para>
/// <b>STA 스레드에서 호출해야 한다</b> — 내부에서 WPF <c>FlowDocument</c> 를 생성한다.
/// </para>
/// <para>
/// 다단 문서의 경우 페이지당 <c>ColumnCount</c> 개의 슬라이스를 생성한다.
/// 슬라이스 순서: (page 0, col 0), (page 0, col 1), …, (page 1, col 0), …
/// 단일 단이면 기존과 동일하게 페이지당 1개.
/// </para>
/// </summary>
public static class PerPageDocumentSplitter
{
    /// <summary>
    /// <paramref name="paginated"/> 에서 페이지·단별 슬라이스 목록을 생성한다.
    /// </summary>
    /// <param name="paginated">분할할 페이지네이션 결과.</param>
    /// <param name="outlineStyles">
    /// 개요 서식 재정의. <c>null</c> 이면 원본 문서의 서식을 사용하고
    /// 그것도 없으면 기본값을 쓴다.
    /// </param>
    public static IReadOnlyList<PerPageDocumentSlice> Split(
        PaginatedDocument paginated,
        OutlineStyleSet?  outlineStyles = null)
    {
        ArgumentNullException.ThrowIfNull(paginated);

        var page   = paginated.PageSettings;
        var styles = outlineStyles
            ?? paginated.Source.OutlineStyles
            ?? OutlineStyleSet.CreateDefault();
        var geo    = new PageGeometry(page);

        int    colCount  = geo.ColumnCount;
        double colGap    = geo.ColGapDip;
        double bodyH     = Math.Max(1.0, geo.PageHeightDip - geo.PadTopDip - geo.PadBottomDip);

        // 원본 문서의 ContainerBlock 계층 복원에 사용할 부모 맵.
        // FlattenBlocks 가 Section 자식들을 낱개로 꺼내므로 빌더에 전달되는
        // coreBlocks 에서 ContainerBlock 이 사라진다 — 부모 맵으로 다시 감싼다.
        var parentMap = BuildParentMap(paginated.Source);

        int totalSlices = paginated.PageCount * colCount;
        var slices      = new PerPageDocumentSlice[totalSlices];

        for (int pageIdx = 0; pageIdx < paginated.PageCount; pageIdx++)
        {
            var pp = paginated.Pages[pageIdx];

            for (int col = 0; col < colCount; col++)
            {
                double colWidth = col < geo.ColWidthsDip.Length ? geo.ColWidthsDip[col] : geo.ColWidthDip;

                var colBlocks  = pp.BodyBlocks.Where(b => b.ColumnIndex == col).ToList();
                var rawBlocks  = colBlocks.Select(b => b.Source).ToList();
                // ContainerBlock 자식들을 다시 부모로 감싸 배경·보더·마진을 복원한다.
                var coreBlocks = parentMap.Count > 0
                    ? ReassembleContainerBlocks(rawBlocks, parentMap)
                    : rawBlocks;
                var fd         = FlowDocumentBuilder.BuildFromBlocks(coreBlocks, page, styles);

                // per-column RTB는 단 폭만 담당; 여백·단 오프셋은 PerPageEditorHost 가 위치로 처리.
                fd.PageWidth   = colWidth;
                fd.PagePadding = new Thickness(0);

                int sliceIdx = pageIdx * colCount + col;
                slices[sliceIdx] = new PerPageDocumentSlice
                {
                    PageIndex    = pageIdx,
                    ColumnIndex  = col,
                    ColumnCount  = colCount,
                    XOffsetDip   = geo.ColumnXOffsetDip(col),
                    PageSettings = page,
                    BodyBlocks   = colBlocks,
                    FlowDocument = fd,
                    BodyWidthDip  = colWidth,
                    BodyHeightDip = bodyH,
                };
            }
        }

        return slices;
    }

    /// <summary>원본 문서를 순회해 ContainerBlock 의 직접 자식 → 부모 ContainerBlock 맵을 만든다.</summary>
    private static Dictionary<Block, ContainerBlock> BuildParentMap(PolyDonkyument doc)
    {
        var map = new Dictionary<Block, ContainerBlock>(ReferenceEqualityComparer.Instance);
        foreach (var section in doc.Sections)
            CollectParents(section.Blocks, map);
        return map;
    }

    private static void CollectParents(IEnumerable<Block> blocks, Dictionary<Block, ContainerBlock> map)
    {
        foreach (var block in blocks)
        {
            if (block is not ContainerBlock container) continue;
            foreach (var child in container.Children)
                map[child] = container;
            // 중첩 컨테이너도 재귀 처리.
            CollectParents(container.Children, map);
        }
    }

    /// <summary>
    /// 평탄화된 블록 목록에서 같은 <see cref="ContainerBlock"/> 부모를 가진 연속 블록을
    /// 새 <see cref="ContainerBlock"/> 으로 다시 감싼다 — 배경·보더·마진 복원.
    /// 부모가 없는 블록은 그대로 통과시킨다.
    /// </summary>
    private static List<Block> ReassembleContainerBlocks(
        List<Block> flat,
        Dictionary<Block, ContainerBlock> parentMap)
    {
        if (flat.Count == 0) return flat;

        var result = new List<Block>(flat.Count);
        int i = 0;
        while (i < flat.Count)
        {
            var block = flat[i];
            if (!parentMap.TryGetValue(block, out var parent))
            {
                result.Add(block);
                i++;
                continue;
            }

            // 같은 부모를 가진 연속 블록 범위를 구한다.
            int j = i + 1;
            while (j < flat.Count &&
                   parentMap.TryGetValue(flat[j], out var p) &&
                   ReferenceEquals(p, parent))
                j++;

            // 부모의 스타일을 그대로 복사하되 자식은 이 페이지/단에 속하는 것만 포함.
            var wrapped = new ContainerBlock
            {
                BorderTopPt       = parent.BorderTopPt,       BorderTopColor    = parent.BorderTopColor,
                BorderRightPt     = parent.BorderRightPt,     BorderRightColor  = parent.BorderRightColor,
                BorderBottomPt    = parent.BorderBottomPt,    BorderBottomColor = parent.BorderBottomColor,
                BorderLeftPt      = parent.BorderLeftPt,      BorderLeftColor   = parent.BorderLeftColor,
                BackgroundColor   = parent.BackgroundColor,
                PaddingTopMm      = parent.PaddingTopMm,      PaddingRightMm    = parent.PaddingRightMm,
                PaddingBottomMm   = parent.PaddingBottomMm,   PaddingLeftMm     = parent.PaddingLeftMm,
                MarginTopMm       = parent.MarginTopMm,       MarginBottomMm    = parent.MarginBottomMm,
                WidthMm           = parent.WidthMm,           HAlign            = parent.HAlign,
                ClassNames        = parent.ClassNames,        Role              = parent.Role,
                Children          = flat.GetRange(i, j - i),
            };
            result.Add(wrapped);
            i = j;
        }
        return result;
    }
}

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

        int totalSlices = paginated.PageCount * colCount;
        var slices      = new PerPageDocumentSlice[totalSlices];

        for (int pageIdx = 0; pageIdx < paginated.PageCount; pageIdx++)
        {
            var pp = paginated.Pages[pageIdx];

            for (int col = 0; col < colCount; col++)
            {
                double colWidth = col < geo.ColWidthsDip.Length ? geo.ColWidthsDip[col] : geo.ColWidthDip;

                var colBlocks  = pp.BodyBlocks.Where(b => b.ColumnIndex == col).ToList();
                var coreBlocks = colBlocks.Select(b => b.Source).ToList();
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
}

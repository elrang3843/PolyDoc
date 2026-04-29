using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using PolyDonky.App.Services;
using PolyDonky.Core;

namespace PolyDonky.App.Pagination;

/// <summary>
/// <see cref="PaginatedDocument"/> 의 블록 배정 정보를 기반으로 페이지별
/// <see cref="PerPageDocumentSlice"/> 를 생성한다.
/// <para>
/// <b>STA 스레드에서 호출해야 한다</b> — 내부에서 WPF <c>FlowDocument</c> 를 생성한다.
/// </para>
/// <para>
/// 용도: per-page 편집기 초기화 시 각 페이지의 본문 블록을 독립된 FlowDocument 로
/// 분리해 페이지별 RichTextBox 에 할당한다.
/// </para>
/// </summary>
public static class PerPageDocumentSplitter
{
    /// <summary>
    /// <paramref name="paginated"/> 에서 페이지별 슬라이스 목록을 생성한다.
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

        double bodyW = Math.Max(1.0, geo.PageWidthDip  - geo.PadLeftDip  - geo.PadRightDip);
        double bodyH = Math.Max(1.0, geo.PageHeightDip - geo.PadTopDip   - geo.PadBottomDip);

        var slices = new PerPageDocumentSlice[paginated.PageCount];
        for (int i = 0; i < paginated.PageCount; i++)
        {
            var pp         = paginated.Pages[i];
            var coreBlocks = pp.BodyBlocks.Select(b => b.Source).ToList();
            var fd         = FlowDocumentBuilder.BuildFromBlocks(coreBlocks, page, styles);

            // per-page RichTextBox 는 본문 영역 폭만 담당; 여백(padding) 은 호출자가 설정한다.
            fd.PageWidth   = bodyW;
            fd.PagePadding = new Thickness(0);

            slices[i] = new PerPageDocumentSlice
            {
                PageIndex    = i,
                PageSettings = page,
                BodyBlocks   = pp.BodyBlocks,
                FlowDocument = fd,
                BodyWidthDip  = bodyW,
                BodyHeightDip = bodyH,
            };
        }
        return slices;
    }
}

using System.Linq;
using PolyDonky.Core;

namespace PolyDonky.App.Services;

/// <summary>
/// 편집창의 다중-페이지 좌표 체계를 한 곳에서 관리한다.
///
/// PaperHost 의 절대 좌표(연속) ↔ (페이지 인덱스, 페이지 로컬 mm) 변환을 제공.
/// 페이지 사이에는 <see cref="InterPageGapDip"/> 만큼의 시각적 갭이 들어간다.
///
/// 좌표 규칙:
/// <list type="bullet">
///   <item>페이지 N 의 절대 Y 시작 = N × (pageHeightDip + gap)</item>
///   <item>페이지 N 의 본문 영역 시작 = 절대 Y 시작 + padTopDip</item>
///   <item>페이지 N 의 본문 영역 끝   = 절대 Y 시작 + pageHeightDip - padBottomDip</item>
/// </list>
/// </summary>
public sealed class PageGeometry
{
    /// <summary>페이지 사이 시각적 갭 (DIP). 회색 배경이 노출되는 영역.</summary>
    public const double InterPageGapDip = 10.0;

    public double PageWidthDip   { get; }
    public double PageHeightDip  { get; }
    public double PadLeftDip     { get; }
    public double PadTopDip      { get; }
    public double PadRightDip    { get; }
    public double PadBottomDip   { get; }

    /// <summary>단 수 (≥ 1).</summary>
    public int    ColumnCount    { get; }
    /// <summary>단 너비 (DIP). 단일 단이면 전체 본문 폭. 다단이면 단 0의 너비.</summary>
    public double ColWidthDip    { get; }
    /// <summary>단 간격 (DIP). 단일 단이면 0.</summary>
    public double ColGapDip      { get; }
    /// <summary>단별 너비 (DIP). 인덱스 = 단 인덱스 (0-based). 길이 = <see cref="ColumnCount"/>.</summary>
    public double[] ColWidthsDip { get; }

    /// <summary>단 인덱스에 해당하는 본문 영역(PadLeft 이후) 기준 X 오프셋 (DIP).</summary>
    public double ColumnXOffsetDip(int colIdx)
    {
        double offset = 0;
        for (int i = 0; i < colIdx && i < ColWidthsDip.Length; i++)
            offset += ColWidthsDip[i] + ColGapDip;
        return offset;
    }

    /// <summary>페이지 N 의 절대 Y 시작 위치. = N × (pageHeight + gap).</summary>
    public double PageStrideDip  => PageHeightDip + InterPageGapDip;

    public PageGeometry(PageSettings page)
    {
        ArgumentNullException.ThrowIfNull(page);
        PageWidthDip  = FlowDocumentBuilder.MmToDip(page.EffectiveWidthMm);
        PageHeightDip = FlowDocumentBuilder.MmToDip(page.EffectiveHeightMm);
        PadLeftDip    = FlowDocumentBuilder.MmToDip(page.MarginLeftMm);
        PadTopDip     = FlowDocumentBuilder.MmToDip(page.MarginTopMm);
        PadRightDip   = FlowDocumentBuilder.MmToDip(page.MarginRightMm);
        PadBottomDip  = FlowDocumentBuilder.MmToDip(page.MarginBottomMm);

        ColumnCount = Math.Max(1, page.ColumnCount);
        ColGapDip   = ColumnCount > 1 ? FlowDocumentBuilder.MmToDip(page.ColumnGapMm) : 0.0;
        double bodyW = Math.Max(1.0, PageWidthDip - PadLeftDip - PadRightDip);
        if (ColumnCount > 1 &&
            page.ColumnWidthsMm is { Count: > 0 } wList &&
            wList.Count == ColumnCount)
        {
            ColWidthsDip = wList.Select(w => Math.Max(10.0, FlowDocumentBuilder.MmToDip(w))).ToArray();
        }
        else
        {
            double equalW = ColumnCount > 1
                ? Math.Max(10.0, (bodyW - ColGapDip * (ColumnCount - 1)) / ColumnCount)
                : bodyW;
            ColWidthsDip = Enumerable.Repeat(equalW, ColumnCount).ToArray();
        }
        ColWidthDip = ColWidthsDip[0];
    }

    /// <summary>(페이지 인덱스, 페이지 로컬 mm) → PaperHost 절대 DIP 좌표.</summary>
    public (double X, double Y) ToAbsoluteDip(int pageIndex, double xMm, double yMm)
    {
        double xDip = FlowDocumentBuilder.MmToDip(xMm);
        double yDip = pageIndex * PageStrideDip + FlowDocumentBuilder.MmToDip(yMm);
        return (xDip, yDip);
    }

    /// <summary>PaperHost 절대 DIP 좌표 → (페이지 인덱스, 페이지 로컬 mm).</summary>
    /// <remarks>
    /// 갭(InterPageGapDip) 영역에 떨어지는 좌표는 가까운 페이지 안으로 클램프된다 —
    /// 사용자가 페이지 사이 갭에 객체를 드롭해도 인접 페이지 가장자리에 안착시킨다.
    /// </remarks>
    public (int PageIndex, double XMm, double YMm) ToPageLocal(double xDip, double yDip)
    {
        if (yDip < 0) yDip = 0;
        double stride = PageStrideDip;
        int pageIndex = (int)Math.Floor(yDip / stride);
        if (pageIndex < 0) pageIndex = 0;

        double yInPage = yDip - pageIndex * stride;
        // 갭 영역(pageHeight 초과)이면 같은 페이지 끝에 클램프 (다음 페이지로 안 넘김 — 사용자 기대 일치)
        if (yInPage > PageHeightDip) yInPage = PageHeightDip;

        return (pageIndex, FlowDocumentBuilder.DipToMm(xDip), FlowDocumentBuilder.DipToMm(yInPage));
    }

    /// <summary>현재 콘텐츠 높이 + 오버레이 최대 페이지 인덱스 기반 페이지 수 계산.</summary>
    public int ComputePageCount(double contentHeightDip, int maxAnchorPageIndex)
    {
        // 본문 콘텐츠 높이 기준 페이지 수
        int byContent = (int)Math.Ceiling(contentHeightDip / PageHeightDip);
        if (byContent < 1) byContent = 1;
        // 오버레이가 더 멀리 있으면 그쪽이 우선
        int byOverlay = Math.Max(1, maxAnchorPageIndex + 1);
        return Math.Max(byContent, byOverlay);
    }

    /// <summary>전체 PaperHost 높이 (DIP). 페이지 N 개 + 갭 (N-1) 개.</summary>
    public double TotalHeightDip(int pageCount)
        => pageCount > 0
            ? pageCount * PageHeightDip + (pageCount - 1) * InterPageGapDip
            : PageHeightDip;
}

using System.Collections.Generic;
using System.Windows;
using PolyDonky.Core;

namespace PolyDonky.App.Pagination;

/// <summary>
/// 하나의 페이지에 속하는 본문 블록 정보.
/// </summary>
public sealed class BlockOnPage
{
    public required Block Source      { get; init; }
    public int            PageIndex   { get; init; }
    /// <summary>다단 문서에서의 단 인덱스 (0-based). 단일 단이면 항상 0.</summary>
    public int            ColumnIndex { get; init; }

    /// <summary>
    /// 단 본문 영역(padding·단 간격 제외) 기준 경계 상자 (DIP).
    /// 오프스크린 RichTextBox 연속 스크롤 공간 기준이므로 FlowDocument 실제 페이지 좌표와
    /// 미묘하게 다를 수 있다. 측정 실패 시 <see cref="Rect.Empty"/>.
    /// </summary>
    public Rect BodyLocalRect { get; init; } = Rect.Empty;
}

/// <summary>
/// 오버레이 블록(글상자·이미지·도형·오버레이 표)의 페이지 배치 정보.
/// </summary>
public sealed class OverlayOnPage
{
    public required Block Source          { get; init; }
    public int            AnchorPageIndex { get; init; }
    public double         XMm             { get; init; }
    public double         YMm             { get; init; }
}

/// <summary>
/// 페이지 단위로 분할된 문서의 한 페이지.
/// </summary>
public sealed class PaginatedPage
{
    public int PageIndex  { get; init; }
    public int PageNumber => PageIndex + 1;

    public IReadOnlyList<BlockOnPage>   BodyBlocks    { get; init; } = Array.Empty<BlockOnPage>();
    public IReadOnlyList<OverlayOnPage> OverlayBlocks { get; init; } = Array.Empty<OverlayOnPage>();
}

/// <summary>
/// WPF DocumentPaginator 결과 — 문서 전체를 페이지 단위로 분할한 구조.
/// <para>
/// 생산: <see cref="FlowDocumentPaginationAdapter.Paginate"/>
/// </para>
/// </summary>
public sealed class PaginatedDocument
{
    public required PolyDonkyument      Source       { get; init; }
    public required PageSettings        PageSettings { get; init; }
    public IReadOnlyList<PaginatedPage> Pages        { get; init; } = Array.Empty<PaginatedPage>();
    public int                          PageCount    => Pages.Count;

    /// <summary>
    /// 페이지·단 슬롯별로 페이지네이션 측정 단계에서 누적된 콘텐츠 높이 (DIP).
    /// 키 = pageIdx × ColumnCount + colIdx. 값 = 슬롯에 배정된 블록 높이 합.
    /// 페이지 경계 결정에 사용된 핵심 길이로, 페이지 본문 슬롯 높이(bodyH) 와 비교해
    /// 클리핑/오버플로 원인 디버깅에 사용한다. Fast-path 진입 시 비어 있을 수 있다.
    /// </summary>
    public IReadOnlyDictionary<int, double> SlotMeasuredFillDip { get; init; }
        = new Dictionary<int, double>();

    /// <summary>
    /// 블록별 측정 진단 정보 (디버그 전용).
    /// slotIdx = pageIdx × ColumnCount + colIdx.
    /// label = 블록 타입과 식별 정보, topY/blockH = 연속 RTB 측정값.
    /// </summary>
    public IReadOnlyList<BlockMeasurementEntry> DebugBlockMeasurements { get; init; }
        = Array.Empty<BlockMeasurementEntry>();

    /// <summary>
    /// 페이지별 PageSettings. 섹션별 다른 설정이 있을 때 사용.
    /// null 이거나 인덱스 범위 밖이면 <see cref="PageSettings"/> 를 폴백으로 사용한다.
    /// </summary>
    public IReadOnlyList<PageSettings>? PerPageSettings { get; init; }

    /// <summary>지정된 페이지의 <see cref="PageSettings"/> 를 반환한다.</summary>
    public PageSettings GetPageSettings(int pageIndex)
    {
        if (PerPageSettings is { } list && (uint)pageIndex < (uint)list.Count)
            return list[pageIndex];
        return PageSettings;
    }
}

/// <summary>블록 측정 진단 항목.</summary>
public sealed class BlockMeasurementEntry
{
    public int    SlotIdx { get; init; }
    public string Label   { get; init; } = "";
    public double TopY    { get; init; }
    public double BottomY { get; init; }
    public double BlockH  { get; init; }
    public double Gap     { get; init; }
}

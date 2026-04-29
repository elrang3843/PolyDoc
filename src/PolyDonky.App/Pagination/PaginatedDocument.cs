using System.Windows;
using PolyDonky.Core;

namespace PolyDonky.App.Pagination;

/// <summary>
/// 하나의 페이지에 속하는 본문 블록 정보.
/// <see cref="BodyLocalRect"/> 는 Phase 3 (per-page 편집기 재설계) 에서 채워진다.
/// Phase 1·2 에서는 <see cref="Rect.Empty"/>.
/// </summary>
public sealed class BlockOnPage
{
    public required Block Source    { get; init; }
    public int            PageIndex { get; init; }

    /// <summary>
    /// 페이지 본문 영역(padding 제외) 기준 경계 상자 (DIP).
    /// Phase 3 이전에는 <see cref="Rect.Empty"/>.
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
}

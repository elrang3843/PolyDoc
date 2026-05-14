using System;

namespace PolyDonky.App.Services;

/// <summary>
/// 필드 코드(Page, NumPages, Date, Time, Author, Title) 를 실제 값으로 치환할 때 참조하는 컨텍스트.
/// <see cref="FlowDocumentBuilder.BuildFromBlocks"/> 호출 시 전달하면 각 슬라이스의 페이지 번호·
/// 총 페이지 수·문서 메타데이터가 필드 코드에 반영된다.
/// </summary>
public sealed class FieldRenderContext
{
    /// <summary>현재 페이지 번호 (1-based).</summary>
    public int PageNumber { get; init; } = 1;

    /// <summary>총 페이지 수 (본문 페이지만; 미주 페이지 제외).</summary>
    public int TotalPages { get; init; } = 1;

    /// <summary>문서 작성자. null 이면 필드 텍스트를 폴백으로 사용.</summary>
    public string? Author { get; init; }

    /// <summary>문서 제목. null 이면 필드 텍스트를 폴백으로 사용.</summary>
    public string? Title { get; init; }

    /// <summary>날짜·시간 필드에 사용할 기준 시각. 기본값은 생성 시점 <see cref="DateTime.Now"/>.</summary>
    public DateTime Now { get; init; } = DateTime.Now;
}

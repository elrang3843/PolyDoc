using System.Collections.Generic;
using System.Windows.Documents;
using PolyDonky.Core;

namespace PolyDonky.App.Pagination;

/// <summary>
/// 한 페이지에 배속된 본문 블록 목록과 해당 페이지용 WPF FlowDocument 를 묶은 슬라이스.
/// <para>
/// <see cref="PerPageDocumentSplitter.Split"/> 이 생산한다.
/// <see cref="FlowDocument"/> 는 STA 스레드 전용 — 다른 스레드에서 접근하지 말 것.
/// </para>
/// </summary>
public sealed class PerPageDocumentSlice
{
    public int                        PageIndex     { get; init; }
    /// <summary>이 슬라이스가 속한 단 인덱스 (0-based). 단일 단 문서에서는 항상 0.</summary>
    public int                        ColumnIndex   { get; init; }
    /// <summary>이 페이지의 총 단 수. 1이면 단일 단.</summary>
    public int                        ColumnCount   { get; init; } = 1;
    /// <summary>
    /// 이 단의 좌측 가장자리가 본문 영역(PadLeft 이후) 기준으로 얼마나 오른쪽에 있는지 (DIP).
    /// 단일 단이면 0.
    /// </summary>
    public double                     XOffsetDip    { get; init; }
    public PageSettings               PageSettings  { get; init; } = new();
    public IReadOnlyList<BlockOnPage> BodyBlocks    { get; init; } = System.Array.Empty<BlockOnPage>();

    /// <summary>
    /// 이 단(또는 페이지)에 속하는 본문 블록만 포함하는 WPF FlowDocument.
    /// STA 스레드에서만 유효. 오버레이(글상자·이미지·도형·오버레이 표) 블록은 포함하지 않는다.
    /// </summary>
    public FlowDocument               FlowDocument  { get; init; } = null!;

    /// <summary>단 너비 (DIP). 다단이면 단 폭, 단일 단이면 본문 폭.</summary>
    public double BodyWidthDip  { get; init; }
    /// <summary>본문 영역 높이 (DIP). 종이 높이 − 상·하 여백.</summary>
    public double BodyHeightDip { get; init; }
}

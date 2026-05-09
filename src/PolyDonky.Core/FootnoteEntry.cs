namespace PolyDonky.Core;

/// <summary>각주(footnote) 또는 미주(endnote) 항목. Run.FootnoteId / EndnoteId 가 이 Id 를 참조한다.</summary>
public sealed class FootnoteEntry
{
    /// <summary>참조 ID — Run.FootnoteId / EndnoteId 와 1:1 매핑.</summary>
    public string Id { get; set; } = string.Empty;

    /// <summary>각주/미주 본문 블록 목록.</summary>
    public IList<Block> Blocks { get; set; } = new List<Block>();
}

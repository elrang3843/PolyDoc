namespace PolyDonky.Core;

/// <summary>자동 목차 블록. 문서의 개요(H1~H6) 단락을 스캔해 만든 항목 목록을 보존한다.</summary>
public sealed class TocBlock : Block
{
    /// <summary>포함할 최소 개요 수준 (1=H1).</summary>
    public int MinLevel { get; set; } = 1;

    /// <summary>포함할 최대 개요 수준 (3=H3).</summary>
    public int MaxLevel { get; set; } = 3;

    /// <summary>마지막으로 빌드된 목차 항목 목록.</summary>
    public IList<TocEntry> Entries { get; set; } = new List<TocEntry>();
}

/// <summary>목차 한 항목 — 개요 수준, 본문 텍스트, 페이지 번호.</summary>
public sealed class TocEntry
{
    public int    Level      { get; set; } = 1;
    public string Text       { get; set; } = string.Empty;
    public int?   PageNumber { get; set; }
}

namespace PolyDonky.Core;

/// <summary>
/// PolyDonky 의 정본 문서 표현. 모든 코덱은 이 모델을 기준으로 입출력한다.
/// </summary>
public sealed class PolyDonkyument
{
    public DocumentMetadata Metadata { get; set; } = new();
    public StyleSheet Styles { get; set; } = new();
    public IList<Section> Sections { get; set; } = new List<Section>();
    public Provenance Provenance { get; set; } = new();

    /// <summary>워터마크 설정 (선택). null 이면 워터마크 없음.</summary>
    public WatermarkSettings? Watermark { get; set; }

    /// <summary>문서 인쇄 가능 여부. false 면 인쇄/미리보기 차단.</summary>
    public bool IsPrintable { get; set; } = true;

    /// <summary>개요 수준별 서식 세트 (선택). null 이면 내장 기본값 사용.</summary>
    public OutlineStyleSet? OutlineStyles { get; set; }

    /// <summary>문서 수준 각주 목록. Run.FootnoteId 가 여기 FootnoteEntry.Id 를 참조한다.</summary>
    public IList<FootnoteEntry> Footnotes { get; set; } = new List<FootnoteEntry>();

    /// <summary>문서 수준 미주 목록. Run.EndnoteId 가 여기 FootnoteEntry.Id 를 참조한다.</summary>
    public IList<FootnoteEntry> Endnotes { get; set; } = new List<FootnoteEntry>();

    /// <summary>비어 있지 않은 단일 섹션 단일 문단을 가진 최소 문서를 생성한다.</summary>
    public static PolyDonkyument Empty()
    {
        var doc = new PolyDonkyument();
        doc.Sections.Add(new Section());
        return doc;
    }

    public IEnumerable<Paragraph> EnumerateParagraphs()
    {
        foreach (var section in Sections)
        {
            foreach (var block in section.Blocks)
            {
                if (block is Paragraph p)
                {
                    yield return p;
                }
            }
        }
    }

    /// <summary>전체 본문을 plain text 로 직렬화한다 (문단 사이는 LF).</summary>
    public string ToPlainText()
    {
        var paragraphs = EnumerateParagraphs().Select(p => p.GetPlainText());
        return string.Join('\n', paragraphs);
    }
}

public sealed class DocumentMetadata
{
    public string? Title { get; set; }
    public string? Author { get; set; }
    public string? Editor { get; set; }
    public string? Application { get; set; } = "PolyDonky";
    public string Language { get; set; } = "ko-KR";
    public DateTimeOffset Created { get; set; } = DateTimeOffset.UtcNow;
    public DateTimeOffset Modified { get; set; } = DateTimeOffset.UtcNow;
    public IDictionary<string, string> Custom { get; set; } = new Dictionary<string, string>();
}

namespace PolyDoc.Core;

/// <summary>
/// PolyDoc 의 정본 문서 표현. 모든 코덱은 이 모델을 기준으로 입출력한다.
/// </summary>
public sealed class PolyDocument
{
    public DocumentMetadata Metadata { get; set; } = new();
    public StyleSheet Styles { get; set; } = new();
    public IList<Section> Sections { get; set; } = new List<Section>();
    public Provenance Provenance { get; set; } = new();

    /// <summary>비어 있지 않은 단일 섹션 단일 문단을 가진 최소 문서를 생성한다.</summary>
    public static PolyDocument Empty()
    {
        var doc = new PolyDocument();
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
    public string? Application { get; set; } = "PolyDoc";
    public string Language { get; set; } = "ko-KR";
    public DateTimeOffset Created { get; set; } = DateTimeOffset.UtcNow;
    public DateTimeOffset Modified { get; set; } = DateTimeOffset.UtcNow;
    public IDictionary<string, string> Custom { get; set; } = new Dictionary<string, string>();
}

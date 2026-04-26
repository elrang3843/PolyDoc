namespace PolyDoc.Core;

/// <summary>
/// 노드 단위 출처 매핑. 외부 포맷에서 import 한 노드가 원본의 어느 위치에서
/// 왔는지 기록하여, 수정되지 않은 부분은 원본 조각을 재사용하는
/// 하이브리드 export 가 가능하도록 한다.
/// </summary>
public sealed class Provenance
{
    public IDictionary<string, SourceAnchor> NodeAnchors { get; set; } = new Dictionary<string, SourceAnchor>();
}

public sealed class SourceAnchor
{
    /// <summary>예: "doc", "hwp", "docx", "hwpx", "html", "md", "txt".</summary>
    public string Format { get; set; } = string.Empty;

    /// <summary>원본 컨테이너 내부 경로 (OPC 파트 경로, OLE stream 이름 등).</summary>
    public string? Path { get; set; }

    public long? Offset { get; set; }
    public long? Length { get; set; }

    /// <summary>HWP 레코드 식별자 등 포맷별 키.</summary>
    public string? RecordId { get; set; }
}

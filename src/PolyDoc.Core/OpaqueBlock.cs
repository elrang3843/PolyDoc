namespace PolyDoc.Core;

/// <summary>
/// IWPF.md 의 "보존 섬(opaque island)" 정책 구현.
/// 우리가 1급 모델로 끌어올리지 못한 외부 포맷의 요소(예: DOCX 의 SDT,
/// 도형, 텍스트박스, 폼 컨트롤 등)를 원본 조각 그대로 보관해 라운드트립
/// 시점에 같은 형식으로 재출력한다. 에디터에선 read-only 자리표시자로 표시.
/// </summary>
public sealed class OpaqueBlock : Block
{
    /// <summary>예: "docx", "hwpx", "doc-binary", "html". 원본 포맷 식별자.</summary>
    public string Format { get; set; } = string.Empty;

    /// <summary>원본 형식 안에서의 하위 종류 (예: "drawing", "sdt", "altChunk").</summary>
    public string? Kind { get; set; }

    /// <summary>XML 기반 포맷이면 원본 OuterXml. 비어 있으면 <see cref="Bytes"/> 사용.</summary>
    public string? Xml { get; set; }

    /// <summary>바이너리 포맷의 원본 조각.</summary>
    public byte[]? Bytes { get; set; }

    /// <summary>편집기에 표시할 사용자 친화적 라벨. 예: "[표]", "[그림]", "[도형]".</summary>
    public string DisplayLabel { get; set; } = "[보존된 개체]";

    public OpaqueBlock()
    {
        Status = NodeStatus.Opaque;
    }
}

namespace PolyDoc.Iwpf;

/// <summary>IWPF 패키지의 매니페스트. 패키지 안의 모든 파트와 그 해시를 기록한다.</summary>
public sealed class IwpfManifest
{
    /// <summary>고정 식별자 — IWPF 매니페스트임을 알린다.</summary>
    public string PackageType { get; set; } = "polydoc.iwpf";

    /// <summary>매니페스트 스키마 버전 (semantic).</summary>
    public string SchemaVersion { get; set; } = "1.0";

    /// <summary>패키지를 만든 PolyDoc 의 빌드 정보.</summary>
    public string ProducedBy { get; set; } = $"PolyDoc/{IwpfFormat.Version}";

    public DateTimeOffset Created { get; set; } = DateTimeOffset.UtcNow;
    public DateTimeOffset Modified { get; set; } = DateTimeOffset.UtcNow;

    /// <summary>경로 → 파트 메타. 키는 ZIP 내부 경로 (forward slash).</summary>
    public IDictionary<string, IwpfManifestEntry> Parts { get; set; } = new Dictionary<string, IwpfManifestEntry>(StringComparer.Ordinal);
}

public sealed class IwpfManifestEntry
{
    public string Path { get; set; } = string.Empty;
    public string MediaType { get; set; } = IwpfMediaTypes.OctetStream;
    public long Size { get; set; }

    /// <summary>SHA-256 hex (소문자, 64자).</summary>
    public string Sha256 { get; set; } = string.Empty;
}

public static class IwpfFormat
{
    public const string Version = "1.0";
}

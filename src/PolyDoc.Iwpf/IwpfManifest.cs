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

    /// <summary>
    /// true 면 본문(content/styles/provenance/resources)이 평문이 아니라
    /// security/payload.bin 안에 AES-256-GCM 으로 봉인돼 있다.
    /// 이 경우 reader 는 비밀번호로 envelope 을 풀어 inner ZIP 을 추출한 뒤 다시 파싱한다.
    /// </summary>
    public bool Encrypted { get; set; }

    /// <summary>
    /// true 면 security/write-lock.json 이 존재하며 저장 시 비밀번호 검증이 필요하다.
    /// <see cref="Encrypted"/> 와 독립적으로 설정할 수 있다.
    /// Both 모드에서는 write-lock 이 inner (복호화된) ZIP 의 매니페스트에 기록된다.
    /// </summary>
    public bool WriteLocked { get; set; }
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

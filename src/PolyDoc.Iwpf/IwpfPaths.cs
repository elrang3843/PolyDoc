namespace PolyDoc.Iwpf;

/// <summary>IWPF 패키지 내 표준 파트 경로. 슬래시는 ZIP 표준에 맞춰 항상 forward slash.</summary>
public static class IwpfPaths
{
    public const string Manifest = "manifest.json";
    public const string DocumentJson = "content/document.json";
    public const string StylesJson = "content/styles.json";
    public const string ProvenanceJson = "provenance/source-map.json";

    public const string ResourcesDir = "resources/";
    public const string ImagesDir = "resources/images/";
    public const string OleDir = "resources/ole/";
    public const string FontsDir = "resources/fonts/";

    public const string FidelityOriginalDir = "fidelity/original/";
    public const string FidelityCapsulesDir = "fidelity/capsules/";

    public const string RenderDir = "render/";
    public const string SignaturesDir = "signatures/";

    public const string SecurityDir       = "security/";
    public const string SecurityEnvelope  = "security/envelope.json";
    public const string SecurityPayload   = "security/payload.bin";
    public const string SecurityWriteLock = "security/write-lock.json";
}

public static class IwpfMediaTypes
{
    public const string Manifest = "application/vnd.polydoc.iwpf.manifest+json";
    public const string Document = "application/vnd.polydoc.iwpf.document+json";
    public const string Styles = "application/vnd.polydoc.iwpf.styles+json";
    public const string Provenance = "application/vnd.polydoc.iwpf.provenance+json";
    public const string Json = "application/json";
    public const string OctetStream = "application/octet-stream";
}

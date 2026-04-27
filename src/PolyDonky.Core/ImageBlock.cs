namespace PolyDonky.Core;

/// <summary>블록 단위 이미지 정렬.</summary>
public enum ImageHAlign { Left, Center, Right }

/// <summary>그림과 본문 텍스트의 배치 관계.</summary>
public enum ImageWrapMode
{
    /// <summary>블록 단위 — 그림이 자체 줄을 차지하고 텍스트가 위/아래로만 배치.</summary>
    Inline,
    /// <summary>텍스트 캐릭터처럼 인라인 — 한 단락 안에서 글자처럼 흐름. (InlineUIContainer)</summary>
    AsText,
    /// <summary>왼쪽 정렬 + 오른쪽으로 텍스트 흐름.</summary>
    WrapLeft,
    /// <summary>오른쪽 정렬 + 왼쪽으로 텍스트 흐름.</summary>
    WrapRight,
    /// <summary>본문 텍스트 위로 그림이 떠있음 (절대 위치, 텍스트 위에 겹침).</summary>
    InFrontOfText,
    /// <summary>본문 텍스트 뒤로 그림이 깔림 (절대 위치, 텍스트 아래에 겹침).</summary>
    BehindText,
}

/// <summary>
/// 임베드된 이미지(블록 단위). 인라인 이미지는 후속 사이클에서 Run 추상화 확장 후 추가된다.
/// 바이너리는 <see cref="Data"/> 에 직접 보관되지만, IWPF 패키징 시
/// <see cref="ResourcePath"/> 가 채워지면 ZIP 의 resources/images/ 로 분리되고 Data 는 비워진다.
/// </summary>
public sealed class ImageBlock : Block
{
    /// <summary>예: "image/png", "image/jpeg", "image/gif", "image/bmp", "image/tiff".</summary>
    public string MediaType { get; set; } = "application/octet-stream";

    /// <summary>이미지 바이너리. IWPF 직렬화 시 패키지 내부 경로로 분리되어 비워질 수 있다.</summary>
    public byte[] Data { get; set; } = Array.Empty<byte>();

    /// <summary>IWPF 패키지에 분리 저장될 때의 내부 경로 (예: resources/images/img-3.png).</summary>
    public string? ResourcePath { get; set; }

    /// <summary>SHA-256 hex (소문자 64자). 무결성 검증 / 중복 제거에 사용.</summary>
    public string? Sha256 { get; set; }

    public double WidthMm { get; set; }
    public double HeightMm { get; set; }

    /// <summary>접근성·검색용 대체 텍스트.</summary>
    public string? Description { get; set; }

    /// <summary>블록 내 가로 정렬 (WrapMode 가 Inline 일 때만 적용).</summary>
    public ImageHAlign HAlign { get; set; } = ImageHAlign.Left;

    /// <summary>그림과 본문 텍스트의 배치 관계.</summary>
    public ImageWrapMode WrapMode { get; set; } = ImageWrapMode.Inline;

    /// <summary>위 여백 (mm).</summary>
    public double MarginTopMm { get; set; }

    /// <summary>아래 여백 (mm).</summary>
    public double MarginBottomMm { get; set; }

    /// <summary>테두리 색상 hex (예: "#FF0000"). null 이면 테두리 없음.</summary>
    public string? BorderColor { get; set; }

    /// <summary>테두리 두께 (pt). 0 이하면 테두리 없음.</summary>
    public double BorderThicknessPt { get; set; }

    /// <summary>오버레이 모드(InFrontOfText/BehindText) X 위치 (mm, 페이지 좌상단 기준).</summary>
    public double OverlayXMm { get; set; }

    /// <summary>오버레이 모드(InFrontOfText/BehindText) Y 위치 (mm, 페이지 좌상단 기준).</summary>
    public double OverlayYMm { get; set; }
}

namespace PolyDonky.Core;

/// <summary>블록 단위 이미지 정렬.</summary>
public enum ImageHAlign { Left, Center, Right }

/// <summary>그림 제목(캡션)의 위치.</summary>
public enum ImageTitlePosition
{
    /// <summary>그림 위쪽(외부) — 별도 행으로 그림 위에 표시.</summary>
    Above,
    /// <summary>그림 아래쪽(외부) — 별도 행으로 그림 아래에 표시.</summary>
    Below,
    /// <summary>그림 위에 겹침 — 상단.</summary>
    OverlayTop,
    /// <summary>그림 위에 겹침 — 가운데.</summary>
    OverlayMiddle,
    /// <summary>그림 위에 겹침 — 하단.</summary>
    OverlayBottom,
}

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
public sealed class ImageBlock : Block, IOverlayAnchored
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

    /// <summary>오버레이 anchoring — 0-based 페이지 인덱스 (0 = 첫 페이지).</summary>
    public int AnchorPageIndex { get; set; }

    /// <summary>오버레이 모드(InFrontOfText/BehindText) X 위치 (mm, **해당 페이지 좌상단 기준**).</summary>
    public double OverlayXMm { get; set; }

    /// <summary>오버레이 모드(InFrontOfText/BehindText) Y 위치 (mm, **해당 페이지 좌상단 기준**).</summary>
    public double OverlayYMm { get; set; }

    // ── 그림 제목(캡션) ──────────────────────────────────────────────────────

    /// <summary>그림 제목 표시 여부. 기본 false.</summary>
    public bool ShowTitle { get; set; }

    /// <summary>그림 제목 텍스트.</summary>
    public string? Title { get; set; }

    /// <summary>제목 글자 서식 (글꼴·크기·굵게·기울임·글자색·배경색·밑줄·자간 등).
    /// 서식 메뉴의 문자서식 다이얼로그(CharFormatWindow)와 동일한 모델을 공유한다.</summary>
    public RunStyle TitleStyle { get; set; } = new();

    /// <summary>제목 배치 위치. 기본 Below.</summary>
    public ImageTitlePosition TitlePosition { get; set; } = ImageTitlePosition.Below;

    /// <summary>제목 가로 정렬. 기본 Center.</summary>
    public ImageHAlign TitleHAlign { get; set; } = ImageHAlign.Center;

    /// <summary>제목 가로 위치 오프셋 (mm). 양수=오른쪽, 음수=왼쪽.</summary>
    public double TitleOffsetXMm { get; set; }

    /// <summary>제목 세로 위치 오프셋 (mm). 양수=아래, 음수=위.</summary>
    public double TitleOffsetYMm { get; set; }
}

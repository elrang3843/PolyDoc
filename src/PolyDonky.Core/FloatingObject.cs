namespace PolyDonky.Core;

using System.Text.Json.Serialization;

// ── 통합 노트 ──────────────────────────────────────────────────────────────
// IWPF 사양에 맞춰 도형/이미지/표/글상자 모두 단일 Block 트리로 통합 (2026-04-29).
// TextBoxObject 가 Block 을 상속하도록 변경, 좌표는 OverlayXMm/OverlayYMm 로 통일,
// Section.FloatingObjects 는 폐기 (Section.Blocks 로 흡수).
//
// 레거시 호환:
// - 옛 IWPF/클립보드 JSON 의 FloatingObject($type:"textbox") + section.floatingObjects 컬렉션은
//   SectionJsonConverter / BlockJsonConverter 가 읽어들여 Section.Blocks 로 자동 마이그레이션.
// - TextBoxObject.XMm/YMm 는 OverlayXMm/OverlayYMm 로 포워딩되는 deprecated 속성으로 남김
//   (옛 코드 빠르게 깨지지 않게 + 옛 JSON 필드 무손실 라운드트립).
// - FloatingObject 추상 클래스 + FloatingObjectJsonConverter 는 폐기. TextBoxObject 의
//   다형 직렬화는 이제 BlockJsonConverter 가 "textbox" discriminator 로 처리.

/// <summary>글상자 모양.</summary>
public enum TextBoxShape
{
    Rectangle = 0,
    Speech    = 1,
    Cloud     = 2,
    Spiky     = 3,
    Lightning = 4,
    Ellipse   = 5,
    Pie       = 6,
}

/// <summary>말풍선 꼬리 방향 (8방향). 박스의 어느 변/모서리에서 꼬리가 뻗는지.</summary>
public enum SpeechPointerDirection
{
    Bottom      = 0,
    BottomLeft  = 1,
    Left        = 2,
    TopLeft     = 3,
    Top         = 4,
    TopRight    = 5,
    Right       = 6,
    BottomRight = 7,
}

/// <summary>글상자 내 텍스트 가로 정렬.</summary>
public enum TextBoxHAlign { Left = 0, Center = 1, Right = 2, Justify = 3 }

/// <summary>글상자 내 텍스트 세로 정렬.</summary>
public enum TextBoxVAlign { Top = 0, Middle = 1, Bottom = 2 }

/// <summary>글상자 다단 구분선 스타일.</summary>
public enum ColumnDividerStyle
{
    None   = 0,
    Solid  = 1,
    Dashed = 2,
    Dotted = 3,
}

/// <summary>
/// 글상자 — 페이지 위에 자유 위치로 배치되는 텍스트 컨테이너.
/// 도형·이미지·표와 동일한 <see cref="ImageWrapMode"/> 배치 체계 + <see cref="OverlayXMm"/>/
/// <see cref="OverlayYMm"/> 좌표 체계를 공유한다 (IWPF 통합 모델).
/// </summary>
public sealed class TextBoxObject : Block, IOverlayAnchored
{
    public TextBoxShape Shape { get; set; } = TextBoxShape.Rectangle;

    /// <summary>본문과의 배치 관계. 글상자 기본은 InFrontOfText (텍스트 위에 떠있음).</summary>
    public ImageWrapMode WrapMode { get; set; } = ImageWrapMode.InFrontOfText;

    /// <summary>오버레이 anchoring — 0-based 페이지 인덱스 (0 = 첫 페이지).</summary>
    public int AnchorPageIndex { get; set; }

    /// <summary>오버레이 X 위치 (mm, **해당 페이지 좌상단 기준**).</summary>
    public double OverlayXMm { get; set; }
    /// <summary>오버레이 Y 위치 (mm, **해당 페이지 좌상단 기준**).</summary>
    public double OverlayYMm { get; set; }

    public double WidthMm { get; set; }
    public double HeightMm { get; set; }

    /// <summary>z-order. 큰 값이 위에 그려진다.</summary>
    public int ZOrder { get; set; }

    // ── 레거시 별칭 (옛 IWPF/클립보드 JSON 에 들어 있던 xMm/yMm 필드 호환 — 읽기 전용) ──
    // 직렬화 출력에는 포함하지 않고, 역직렬화 시 입력값을 OverlayXMm/OverlayYMm 로 흘려넣는다.
    [JsonPropertyName("xMm"), JsonInclude]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    internal double LegacyXMm
    {
        get => 0;
        set { if (value != 0 && OverlayXMm == 0) OverlayXMm = value; }
    }

    [JsonPropertyName("yMm"), JsonInclude]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    internal double LegacyYMm
    {
        get => 0;
        set { if (value != 0 && OverlayYMm == 0) OverlayYMm = value; }
    }

    /// <summary>테두리 색 (hex). null/빈 = 검정.</summary>
    public string? BorderColor { get; set; }

    public double BorderThicknessPt { get; set; } = 1.0;

    /// <summary>배경색 (hex). null/빈 = 흰색.</summary>
    public string? BackgroundColor { get; set; }

    // ── 4방향 안쪽 여백 (mm) ──────────────────────────────────────────────────
    public double PaddingTopMm    { get; set; } = 2.0;
    public double PaddingBottomMm { get; set; } = 2.0;
    public double PaddingLeftMm   { get; set; } = 2.0;
    public double PaddingRightMm  { get; set; } = 2.0;

    // ── 텍스트 정렬 ───────────────────────────────────────────────────────────
    public TextBoxHAlign HAlign { get; set; } = TextBoxHAlign.Center;
    public TextBoxVAlign VAlign { get; set; } = TextBoxVAlign.Middle;

    /// <summary>박스 전체 회전각 (도, 시계방향). -360~360.</summary>
    public double RotationAngleDeg { get; set; }

    // ── 글자 방향 ─────────────────────────────────────────────────────────────
    public TextOrientation TextOrientation { get; set; } = TextOrientation.Horizontal;
    public TextProgression TextProgression { get; set; } = TextProgression.Rightward;

    // ── 모양별 형태 파라미터 ─────────────────────────────────────────────────
    public SpeechPointerDirection SpeechDirection { get; set; } = SpeechPointerDirection.Bottom;
    public int CloudPuffCount { get; set; } = 10;
    public int SpikeCount { get; set; } = 12;
    public int LightningBendCount { get; set; } = 2;
    public double PieStartAngleDeg { get; set; } = 0;
    public double PieSweepAngleDeg { get; set; } = 270;

    // ── 다단 (글상자 내부 다단 레이아웃) ─────────────────────────────────────
    /// <summary>단 수. 1 이면 단일 단(기본값), 2 이상이면 다단.</summary>
    public int ColumnCount { get; set; } = 1;
    /// <summary>단 간격 (mm). 단일 단이면 무시.</summary>
    public double ColumnGapMm { get; set; } = 5.0;
    /// <summary>각 단의 너비 (mm). null/빈 = 균등 배분. 길이가 ColumnCount 와 다르면 균등 배분.</summary>
    public List<double>? ColumnWidthsMm { get; set; }

    // ── 단 구분선 ────────────────────────────────────────────────────────
    /// <summary>단 구분선 표시 여부. ColumnCount &gt; 1 일 때만 의미 있음.</summary>
    public bool               ColumnDividerVisible      { get; set; } = false;
    /// <summary>단 구분선 색 (hex). 기본값 회색.</summary>
    public string             ColumnDividerColor        { get; set; } = "#888888";
    /// <summary>단 구분선 두께 (pt).</summary>
    public double             ColumnDividerThicknessPt  { get; set; } = 0.7;
    /// <summary>단 구분선 스타일.</summary>
    public ColumnDividerStyle ColumnDividerStyle        { get; set; } = ColumnDividerStyle.Dashed;

    /// <summary>본문 블록. 최소 1개의 빈 Paragraph 를 포함하도록 기본값 설정.</summary>
    public IList<Block> Content { get; set; } = new List<Block> { new Paragraph() };

    /// <summary>본문을 plain text 로 직렬화 (단락 사이는 LF). UI 평문 편집 보조용.</summary>
    public string GetPlainText()
        => string.Join('\n', Content.OfType<Paragraph>().Select(p => p.GetPlainText()));

    /// <summary>plain text 를 단락별로 끊어 Content 에 채운다 (편집기 동기화용).</summary>
    public void SetPlainText(string? text)
    {
        Content.Clear();
        var lines = (text ?? "").Split('\n');
        foreach (var line in lines)
        {
            var p = new Paragraph();
            if (line.Length > 0) p.AddText(line);
            Content.Add(p);
        }
        if (Content.Count == 0) Content.Add(new Paragraph());
    }
}

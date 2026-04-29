using System.Text.Json.Serialization;

namespace PolyDonky.Core;

public sealed class Section
{
    public string? Id { get; set; }
    public PageSettings Page { get; set; } = new();
    public IList<Block> Blocks { get; set; } = new List<Block>();

    /// <summary>
    /// 레거시 호환 — 옛 IWPF/문서 JSON 의 <c>floatingObjects</c> 컬렉션을 받아
    /// <see cref="Blocks"/> 로 즉시 흡수한다 (글상자가 Block 으로 통합된 2026-04-29 이후).
    /// 출력에는 항상 <c>null</c> 이므로 <see cref="JsonIgnoreCondition.WhenWritingNull"/> 가 생략한다.
    /// </summary>
    [JsonPropertyName("floatingObjects"), JsonInclude]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    internal IList<Block>? LegacyFloatingObjects
    {
        get => null;
        set
        {
            if (value is null) return;
            foreach (var item in value) Blocks.Add(item);
        }
    }
}

// ── 용지 크기 종류 ────────────────────────────────────────────────────────
public enum PaperSizeKind
{
    Custom = 0,

    // ISO A 시리즈
    A0, A1, A2, A3, A4, A5, A6, A7,

    // ISO B 시리즈
    B4_ISO, B5_ISO, B6_ISO,

    // JIS/KS B 시리즈 (한국·일본 출판 실무)
    B4_JIS, B5_JIS, B6_JIS,

    // 미국·국제 표준
    Letter,     // 8.5"×11"
    Legal,      // 8.5"×14"
    Ledger,     // 17"×11" (가로)
    Tabloid,    // 11"×17"
    Statement,  // 5.5"×8.5"
    Executive,  // 7.25"×10.5"
    Folio,      // 8.5"×13"

    // 신문 판형
    Broadsheet,         // 전통 대판 585×381 mm
    Berliner,           // 베를리너/노르딕 470×315 mm
    Compact,            // 콤팩트(영국 타블로이드) 380×282 mm
    KoreanBroadsheet,   // 한국 신문 대판 546×393 mm

    // 한국 도서 판형
    SinGukPan,  // 신국판 152×225 mm  (가장 흔한 한국 도서)
    GukPan,     // 국판 148×210 mm  (= A5)
    Pan46Bae,   // 4×6배판 188×257 mm
    CrownPan,   // 크라운판 176×248 mm
    Pan46,      // 46판 127×188 mm
    MassMarket, // 소형 문고판 106×175 mm

    // 국제 도서 판형
    BookRoyal,       // Royal 156×234 mm
    BookDemy,        // Demy 138×216 mm
    BookTrade,       // Trade (6"×9") 152×229 mm
    BookQuarto,      // Quarto 189×246 mm
    BookLargeFormat, // Large Format 216×280 mm
}

public enum PageOrientation { Portrait, Landscape }

// ── 글자 진행 방향 ────────────────────────────────────────────────────────
/// <summary>본문 글자 방향 (가로/세로 쓰기).</summary>
public enum TextOrientation
{
    Horizontal = 0,  // 가로쓰기
    Vertical   = 1,  // 세로쓰기
}

/// <summary>본문/행 진행 방향. 가로쓰기에서는 글자 방향, 세로쓰기에서는 행(=세로 단) 진행 방향.</summary>
public enum TextProgression
{
    Rightward = 0,  // 오른쪽으로 (가로: 왼→오 / 세로: 행이 왼→오)
    Leftward  = 1,  // 왼쪽으로   (가로: 오→왼 / 세로: 행이 오→왼, 전통 CJK)
}

// ── 머리말·꼬리말 정의 (이번 사이클은 모델만; UI는 다음 사이클) ──────────
public sealed class HeaderFooterContent
{
    public string? Left   { get; set; }
    public string? Center { get; set; }
    public string? Right  { get; set; }
}

// ── 페이지 설정 ───────────────────────────────────────────────────────────
/// <summary>
/// 용지·여백·다단·머리말꼬리말 등 섹션 단위 페이지 설정. 단위는 mm.
/// </summary>
public sealed class PageSettings
{
    // 용지 크기 (Custom 이면 WidthMm/HeightMm 를 직접 사용)
    public PaperSizeKind SizeKind { get; set; } = PaperSizeKind.A4;
    public double WidthMm  { get; set; } = 210;   // Custom 또는 PaperDimensions 자동 동기
    public double HeightMm { get; set; } = 297;

    public PageOrientation Orientation { get; set; } = PageOrientation.Portrait;

    // ── 글자 방향 ─────────────────────────────────────────────────────────────
    /// <summary>가로/세로 쓰기. 세로쓰기는 모델만 저장 — 렌더링은 다음 사이클에서 지원.</summary>
    public TextOrientation TextOrientation { get; set; } = TextOrientation.Horizontal;
    /// <summary>본문/행 진행 방향. 가로쓰기에서는 글자 방향, 세로쓰기에서는 행 진행 방향.</summary>
    public TextProgression TextProgression { get; set; } = TextProgression.Rightward;

    // 용지 색상 (null / "" = 흰색 기본)
    public string? PaperColor { get; set; }

    // 여백 (mm)
    public double MarginTopMm    { get; set; } = 20;
    public double MarginBottomMm { get; set; } = 20;
    public double MarginLeftMm   { get; set; } = 25;
    public double MarginRightMm  { get; set; } = 25;
    public double MarginHeaderMm { get; set; } = 10;
    public double MarginFooterMm { get; set; } = 10;

    // 다단
    public int    ColumnCount  { get; set; } = 1;
    public double ColumnGapMm  { get; set; } = 8;

    // 페이지 번호
    public int PageNumberStart { get; set; } = 1;

    // 머리말·꼬리말 (모델; UI는 다음 사이클)
    public HeaderFooterContent Header      { get; set; } = new();
    public HeaderFooterContent Footer      { get; set; } = new();
    public bool DifferentFirstPage { get; set; }
    public bool DifferentOddEven   { get; set; }

    // 여백 안내선 표시 (편집 화면)
    public bool ShowMarginGuides { get; set; } = true;

    // ── 계산 헬퍼 ──────────────────────────────────────────────
    /// <summary>방향 보정 후 실제 용지 너비 (mm).</summary>
    public double EffectiveWidthMm  => Orientation == PageOrientation.Portrait ? WidthMm : HeightMm;
    /// <summary>방향 보정 후 실제 용지 높이 (mm).</summary>
    public double EffectiveHeightMm => Orientation == PageOrientation.Portrait ? HeightMm : WidthMm;

    /// <summary>
    /// <see cref="SizeKind"/> 에 해당하는 표준 치수 (너비×높이 mm, 세로 방향 기준).
    /// Custom 이면 null.
    /// </summary>
    public static (double W, double H)? GetStandardDimensions(PaperSizeKind kind) => kind switch
    {
        // ISO A
        PaperSizeKind.A0 => (841,  1189),
        PaperSizeKind.A1 => (594,   841),
        PaperSizeKind.A2 => (420,   594),
        PaperSizeKind.A3 => (297,   420),
        PaperSizeKind.A4 => (210,   297),
        PaperSizeKind.A5 => (148,   210),
        PaperSizeKind.A6 => (105,   148),
        PaperSizeKind.A7 => (74,    105),
        // ISO B
        PaperSizeKind.B4_ISO => (250,   353),
        PaperSizeKind.B5_ISO => (176,   250),
        PaperSizeKind.B6_ISO => (125,   176),
        // JIS/KS B
        PaperSizeKind.B4_JIS => (257,   364),
        PaperSizeKind.B5_JIS => (182,   257),
        PaperSizeKind.B6_JIS => (128,   182),
        // 미국·국제
        PaperSizeKind.Letter    => (215.9, 279.4),
        PaperSizeKind.Legal     => (215.9, 355.6),
        PaperSizeKind.Ledger    => (431.8, 279.4),
        PaperSizeKind.Tabloid   => (279.4, 431.8),
        PaperSizeKind.Statement => (139.7, 215.9),
        PaperSizeKind.Executive => (184.2, 266.7),
        PaperSizeKind.Folio     => (215.9, 330.2),
        // 신문
        PaperSizeKind.Broadsheet        => (381,   585),
        PaperSizeKind.Berliner          => (315,   470),
        PaperSizeKind.Compact           => (282,   380),
        PaperSizeKind.KoreanBroadsheet  => (393,   546),
        // 한국 도서
        PaperSizeKind.SinGukPan  => (152,   225),
        PaperSizeKind.GukPan     => (148,   210),
        PaperSizeKind.Pan46Bae   => (188,   257),
        PaperSizeKind.CrownPan   => (176,   248),
        PaperSizeKind.Pan46      => (127,   188),
        PaperSizeKind.MassMarket => (106,   175),
        // 국제 도서
        PaperSizeKind.BookRoyal        => (156,   234),
        PaperSizeKind.BookDemy         => (138,   216),
        PaperSizeKind.BookTrade        => (152,   229),
        PaperSizeKind.BookQuarto       => (189,   246),
        PaperSizeKind.BookLargeFormat  => (216,   280),
        // Custom
        _ => null,
    };

    /// <summary>
    /// <paramref name="kind"/> 의 치수를 적용한다. Custom 이면 아무것도 변경하지 않는다.
    /// </summary>
    public void ApplySizeKind(PaperSizeKind kind)
    {
        SizeKind = kind;
        var dim = GetStandardDimensions(kind);
        if (dim is { } d)
        {
            WidthMm  = d.W;
            HeightMm = d.H;
        }
    }
}

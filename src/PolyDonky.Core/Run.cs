namespace PolyDonky.Core;

/// <summary>문단 내 동일 서식의 문자 시퀀스.</summary>
public sealed class Run
{
    public string Text { get; set; } = string.Empty;
    public RunStyle Style { get; set; } = new();

    /// <summary>LaTeX 수식 소스. null 이면 일반 텍스트 Run.</summary>
    public string? LatexSource { get; set; }

    /// <summary>별행(display) 수식 여부. LatexSource 가 null 이면 무시.</summary>
    public bool IsDisplayEquation { get; set; }

    /// <summary>이모지 키 ("{Section}_{name}", 예: "Status_done"). null 이면 일반 텍스트 Run.
    /// Resources/Emojis/{Section}/{name}.png 와 일대일 대응. 라운드트립 시 보존.</summary>
    public string? EmojiKey { get; set; }

    /// <summary>이모지 기준선 정렬. null 이면 Center.</summary>
    public EmojiAlignment? EmojiAlignment { get; set; }

    /// <summary>하이퍼링크 URL. null/빈 문자열이면 일반 텍스트.
    /// Markdown 의 [text](url), 자동 링크 등에서 사용.</summary>
    public string? Url { get; set; }

    /// <summary>각주 참조 ID. null 이면 일반 Run.
    /// PolyDonkyument.Footnotes 의 FootnoteEntry.Id 와 매핑.
    /// DOCX: w:footnoteReference, HWPX: hp:ctrl ctrlID="FOOT_NOTE"</summary>
    public string? FootnoteId { get; set; }

    /// <summary>미주 참조 ID. null 이면 일반 Run.
    /// PolyDonkyument.Endnotes 의 FootnoteEntry.Id 와 매핑.
    /// DOCX: w:endnoteReference, HWPX: hp:ctrl ctrlID="END_NOTE"</summary>
    public string? EndnoteId { get; set; }

    /// <summary>인라인 필드 종류. null 이면 필드가 아닌 일반 텍스트 Run.</summary>
    public FieldType? Field { get; set; }

    /// <summary>루비 주석 텍스트 (한자 위 후리가나 등). null 이면 루비 없음.
    /// HTML <c>&lt;ruby&gt;&lt;rb&gt;base&lt;/rb&gt;&lt;rt&gt;annotation&lt;/rt&gt;&lt;/ruby&gt;</c>
    /// 에서 rt 콘텐츠를 여기에 저장, 베이스 텍스트는 Text 에 저장.</summary>
    public string? RubyText { get; set; }

    /// <summary>모든 필드를 복사한 깊은 복제본 — Style 도 새 인스턴스로.</summary>
    public Run Clone() => new()
    {
        Text              = Text,
        Style             = Style.Clone(),
        LatexSource       = LatexSource,
        IsDisplayEquation = IsDisplayEquation,
        EmojiKey          = EmojiKey,
        EmojiAlignment    = EmojiAlignment,
        Url               = Url,
        FootnoteId        = FootnoteId,
        EndnoteId         = EndnoteId,
        Field             = Field,
        RubyText          = RubyText,
    };
}

/// <summary>인라인 필드 종류 — 렌더링 시 현재 값으로 치환되는 자동 삽입 값.</summary>
public enum FieldType
{
    /// <summary>현재 페이지 번호.</summary>
    Page,
    /// <summary>총 페이지 수.</summary>
    NumPages,
    /// <summary>현재 날짜.</summary>
    Date,
    /// <summary>현재 시간.</summary>
    Time,
    /// <summary>문서 작성자.</summary>
    Author,
    /// <summary>문서 제목.</summary>
    Title,
}

/// <summary>이모지 인라인 이미지의 기준선 정렬.</summary>
public enum EmojiAlignment { TextTop, Center, TextBottom, Baseline }

public sealed class RunStyle
{
    public string? FontFamily { get; set; }
    public double FontSizePt { get; set; } = 11;
    public bool Bold { get; set; }
    public bool Italic { get; set; }
    public bool Underline { get; set; }
    public bool Strikethrough { get; set; }
    public bool Overline { get; set; }
    public bool Superscript { get; set; }
    public bool Subscript { get; set; }
    public Color? Foreground { get; set; }
    public Color? Background { get; set; }

    /// <summary>한글 조판: 글자폭 (장평). 100 = 표준.</summary>
    public double WidthPercent { get; set; } = 100;

    /// <summary>한글 조판: 자간 (px 단위). 0 = 표준.</summary>
    public double LetterSpacingPx { get; set; }

    /// <summary>CSS text-transform. None = 변환 없음.</summary>
    public TextTransform TextTransform { get; set; } = TextTransform.None;

    /// <summary>CSS word-spacing (px 단위). 0 = 표준.</summary>
    public double WordSpacingPx { get; set; }

    /// <summary>CSS font-variant: small-caps. true 이면 소형 대문자 표시.</summary>
    public bool FontVariantSmallCaps { get; set; }

    /// <summary>모든 필드를 복사한 깊은 복제본.</summary>
    public RunStyle Clone() => new()
    {
        FontFamily           = FontFamily,
        FontSizePt           = FontSizePt,
        Bold                 = Bold,
        Italic               = Italic,
        Underline            = Underline,
        Strikethrough        = Strikethrough,
        Overline             = Overline,
        Superscript          = Superscript,
        Subscript            = Subscript,
        Foreground           = Foreground,
        Background           = Background,
        WidthPercent         = WidthPercent,
        LetterSpacingPx      = LetterSpacingPx,
        TextTransform        = TextTransform,
        WordSpacingPx        = WordSpacingPx,
        FontVariantSmallCaps = FontVariantSmallCaps,
    };
}

/// <summary>CSS text-transform 변환 종류.</summary>
public enum TextTransform
{
    None,
    Uppercase,
    Lowercase,
    Capitalize,
}

public readonly record struct Color(byte R, byte G, byte B, byte A = 255)
{
    public static Color Black { get; } = new(0, 0, 0);
    public static Color White { get; } = new(255, 255, 255);

    public string ToHex() => A == 255
        ? $"#{R:X2}{G:X2}{B:X2}"
        : $"#{R:X2}{G:X2}{B:X2}{A:X2}";

    public override string ToString() => ToHex();

    public static Color FromHex(string hex)
    {
        ArgumentNullException.ThrowIfNull(hex);
        var raw = hex.AsSpan().TrimStart('#');

        // 3/4자리 단축 (#RGB / #RGBA) → #RRGGBB / #RRGGBBAA 로 확장.
        string? expanded = null;
        if (raw.Length is 3 or 4)
        {
            var sb = new System.Text.StringBuilder(raw.Length * 2);
            foreach (var ch in raw) { sb.Append(ch); sb.Append(ch); }
            expanded = sb.ToString();
        }

        var span = expanded is not null ? expanded.AsSpan() : raw;

        if (span.Length is not (6 or 8))
            throw new FormatException($"Invalid color hex: '{hex}'. Expected #RGB, #RGBA, #RRGGBB, or #RRGGBBAA.");

        byte r = byte.Parse(span[..2], System.Globalization.NumberStyles.HexNumber);
        byte g = byte.Parse(span[2..4], System.Globalization.NumberStyles.HexNumber);
        byte b = byte.Parse(span[4..6], System.Globalization.NumberStyles.HexNumber);
        byte a = span.Length == 8
            ? byte.Parse(span[6..8], System.Globalization.NumberStyles.HexNumber)
            : (byte)255;
        return new Color(r, g, b, a);
    }
}

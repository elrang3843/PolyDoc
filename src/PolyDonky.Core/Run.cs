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
    };
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

    /// <summary>모든 필드를 복사한 깊은 복제본.</summary>
    public RunStyle Clone() => new()
    {
        FontFamily      = FontFamily,
        FontSizePt      = FontSizePt,
        Bold            = Bold,
        Italic          = Italic,
        Underline       = Underline,
        Strikethrough   = Strikethrough,
        Overline        = Overline,
        Superscript     = Superscript,
        Subscript       = Subscript,
        Foreground      = Foreground,
        Background      = Background,
        WidthPercent    = WidthPercent,
        LetterSpacingPx = LetterSpacingPx,
    };
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
        var span = hex.AsSpan().TrimStart('#');
        if (span.Length is not (6 or 8))
        {
            throw new FormatException($"Invalid color hex: '{hex}'. Expected RRGGBB or RRGGBBAA.");
        }

        byte r = byte.Parse(span[..2], System.Globalization.NumberStyles.HexNumber);
        byte g = byte.Parse(span[2..4], System.Globalization.NumberStyles.HexNumber);
        byte b = byte.Parse(span[4..6], System.Globalization.NumberStyles.HexNumber);
        byte a = span.Length == 8
            ? byte.Parse(span[6..8], System.Globalization.NumberStyles.HexNumber)
            : (byte)255;
        return new Color(r, g, b, a);
    }
}

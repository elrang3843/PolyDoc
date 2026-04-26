namespace PolyDoc.Core;

/// <summary>문단 내 동일 서식의 문자 시퀀스.</summary>
public sealed class Run
{
    public string Text { get; set; } = string.Empty;
    public RunStyle Style { get; set; } = new();
}

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

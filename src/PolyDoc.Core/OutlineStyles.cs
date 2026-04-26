namespace PolyDoc.Core;

public enum NumberingStyle
{
    None,
    Decimal,        // 1. 2. 3.
    AlphaLower,     // a. b. c.
    AlphaUpper,     // A. B. C.
    RomanLower,     // i. ii. iii.
    RomanUpper,     // I. II. III.
    HangulSyllable, // 가. 나. 다.
    HangulOrdinal,  // 첫째. 둘째.
}

public sealed class OutlineNumbering
{
    public NumberingStyle Style { get; set; } = NumberingStyle.None;
    public string Prefix { get; set; } = "";
    public string Suffix { get; set; } = ".";
    public bool RestartFromHigher { get; set; } = true;
}

public sealed class OutlineBorder
{
    public bool ShowTop { get; set; }
    public bool ShowBottom { get; set; }
    public string? Color { get; set; }  // hex (#RRGGBB), null = 기본 OnSurface
}

public sealed class OutlineLevelStyle
{
    public RunStyle Char { get; set; } = new();
    public ParagraphStyle Para { get; set; } = new();
    public OutlineNumbering Numbering { get; set; } = new();
    public OutlineBorder Border { get; set; } = new();
    public string? BackgroundColor { get; set; }  // hex, null = 없음
}

/// <summary>
/// 개요 수준별 서식 세트. 프리셋으로 미리 만들어 두고 사용자가 수정해서 쓴다.
/// </summary>
public sealed class OutlineStyleSet
{
    public string Name { get; set; } = "기본";

    /// <summary>키: (int)OutlineLevel (0=Body, 1=H1 … 6=H6).</summary>
    public IDictionary<int, OutlineLevelStyle> Levels { get; set; }
        = new Dictionary<int, OutlineLevelStyle>();

    public OutlineLevelStyle GetLevel(OutlineLevel level)
        => Levels.TryGetValue((int)level, out var s) ? s : DefaultForLevel(level);

    public void SetLevel(OutlineLevel level, OutlineLevelStyle style)
        => Levels[(int)level] = style;

    // ── 내장 프리셋 ──────────────────────────────────────────────

    public static OutlineStyleSet CreateDefault()
    {
        var set = new OutlineStyleSet { Name = "기본" };
        set.SetLevel(OutlineLevel.Body, new OutlineLevelStyle
        {
            Char = new RunStyle { FontSizePt = 11 },
            Para = new ParagraphStyle { LineHeightFactor = 1.6 },
        });
        set.SetLevel(OutlineLevel.H1, new OutlineLevelStyle
        {
            Char = new RunStyle { FontSizePt = 24, Bold = true },
            Para = new ParagraphStyle { SpaceBeforePt = 12, SpaceAfterPt = 6, LineHeightFactor = 1.3, Outline = OutlineLevel.H1 },
        });
        set.SetLevel(OutlineLevel.H2, new OutlineLevelStyle
        {
            Char = new RunStyle { FontSizePt = 20, Bold = true },
            Para = new ParagraphStyle { SpaceBeforePt = 10, SpaceAfterPt = 4, LineHeightFactor = 1.3, Outline = OutlineLevel.H2 },
        });
        set.SetLevel(OutlineLevel.H3, new OutlineLevelStyle
        {
            Char = new RunStyle { FontSizePt = 17, Bold = true },
            Para = new ParagraphStyle { SpaceBeforePt = 8, SpaceAfterPt = 4, LineHeightFactor = 1.3, Outline = OutlineLevel.H3 },
        });
        set.SetLevel(OutlineLevel.H4, new OutlineLevelStyle
        {
            Char = new RunStyle { FontSizePt = 15, Bold = true },
            Para = new ParagraphStyle { SpaceBeforePt = 6, SpaceAfterPt = 2, Outline = OutlineLevel.H4 },
        });
        set.SetLevel(OutlineLevel.H5, new OutlineLevelStyle
        {
            Char = new RunStyle { FontSizePt = 13, Bold = true },
            Para = new ParagraphStyle { SpaceBeforePt = 4, SpaceAfterPt = 2, Outline = OutlineLevel.H5 },
        });
        set.SetLevel(OutlineLevel.H6, new OutlineLevelStyle
        {
            Char = new RunStyle { FontSizePt = 12, Bold = true, Italic = true },
            Para = new ParagraphStyle { SpaceBeforePt = 4, SpaceAfterPt = 2, Outline = OutlineLevel.H6 },
        });
        return set;
    }

    public static OutlineStyleSet CreateAcademic()
    {
        var set = new OutlineStyleSet { Name = "학술" };
        set.SetLevel(OutlineLevel.Body, new OutlineLevelStyle
        {
            Char = new RunStyle { FontSizePt = 10.5 },
            Para = new ParagraphStyle { LineHeightFactor = 1.8, IndentFirstLineMm = 5 },
        });
        set.SetLevel(OutlineLevel.H1, new OutlineLevelStyle
        {
            Char = new RunStyle { FontSizePt = 16, Bold = true },
            Para = new ParagraphStyle { Alignment = Alignment.Center, SpaceBeforePt = 24, SpaceAfterPt = 12, Outline = OutlineLevel.H1 },
            Numbering = new OutlineNumbering { Style = NumberingStyle.RomanUpper, Suffix = ".", Prefix = "" },
            Border = new OutlineBorder { ShowBottom = true },
        });
        set.SetLevel(OutlineLevel.H2, new OutlineLevelStyle
        {
            Char = new RunStyle { FontSizePt = 14, Bold = true },
            Para = new ParagraphStyle { SpaceBeforePt = 18, SpaceAfterPt = 6, Outline = OutlineLevel.H2 },
            Numbering = new OutlineNumbering { Style = NumberingStyle.Decimal, Suffix = ".", Prefix = "" },
        });
        set.SetLevel(OutlineLevel.H3, new OutlineLevelStyle
        {
            Char = new RunStyle { FontSizePt = 12, Bold = true, Italic = true },
            Para = new ParagraphStyle { SpaceBeforePt = 12, SpaceAfterPt = 4, Outline = OutlineLevel.H3 },
            Numbering = new OutlineNumbering { Style = NumberingStyle.Decimal, Suffix = ")", Prefix = "(" },
        });
        set.SetLevel(OutlineLevel.H4, new OutlineLevelStyle
        {
            Char = new RunStyle { FontSizePt = 11, Bold = true },
            Para = new ParagraphStyle { Outline = OutlineLevel.H4 },
        });
        set.SetLevel(OutlineLevel.H5, new OutlineLevelStyle
        {
            Char = new RunStyle { FontSizePt = 11, Italic = true },
            Para = new ParagraphStyle { Outline = OutlineLevel.H5 },
        });
        set.SetLevel(OutlineLevel.H6, new OutlineLevelStyle
        {
            Char = new RunStyle { FontSizePt = 10.5, Italic = true },
            Para = new ParagraphStyle { Outline = OutlineLevel.H6 },
        });
        return set;
    }

    public static OutlineStyleSet CreateBusiness()
    {
        var set = new OutlineStyleSet { Name = "비즈니스" };
        set.SetLevel(OutlineLevel.Body, new OutlineLevelStyle
        {
            Char = new RunStyle { FontSizePt = 10 },
            Para = new ParagraphStyle { LineHeightFactor = 1.5 },
        });
        set.SetLevel(OutlineLevel.H1, new OutlineLevelStyle
        {
            Char = new RunStyle { FontSizePt = 18, Bold = true, Foreground = new Color(0x1F, 0x49, 0x7D, 0xFF) },
            Para = new ParagraphStyle { SpaceBeforePt = 18, SpaceAfterPt = 6, Outline = OutlineLevel.H1 },
            Border = new OutlineBorder { ShowBottom = true, Color = "#1F497D" },
        });
        set.SetLevel(OutlineLevel.H2, new OutlineLevelStyle
        {
            Char = new RunStyle { FontSizePt = 14, Bold = true, Foreground = new Color(0x2E, 0x74, 0xB5, 0xFF) },
            Para = new ParagraphStyle { SpaceBeforePt = 12, SpaceAfterPt = 4, Outline = OutlineLevel.H2 },
        });
        set.SetLevel(OutlineLevel.H3, new OutlineLevelStyle
        {
            Char = new RunStyle { FontSizePt = 12, Bold = true },
            Para = new ParagraphStyle { SpaceBeforePt = 8, SpaceAfterPt = 4, Outline = OutlineLevel.H3 },
        });
        set.SetLevel(OutlineLevel.H4, new OutlineLevelStyle
        {
            Char = new RunStyle { FontSizePt = 11, Bold = true, Italic = true },
            Para = new ParagraphStyle { Outline = OutlineLevel.H4 },
        });
        set.SetLevel(OutlineLevel.H5, new OutlineLevelStyle
        {
            Char = new RunStyle { FontSizePt = 10.5, Bold = true },
            Para = new ParagraphStyle { Outline = OutlineLevel.H5 },
        });
        set.SetLevel(OutlineLevel.H6, new OutlineLevelStyle
        {
            Char = new RunStyle { FontSizePt = 10, Italic = true },
            Para = new ParagraphStyle { Outline = OutlineLevel.H6 },
        });
        return set;
    }

    public static OutlineStyleSet CreateModern()
    {
        var set = new OutlineStyleSet { Name = "모던" };
        set.SetLevel(OutlineLevel.Body, new OutlineLevelStyle
        {
            Char = new RunStyle { FontSizePt = 11 },
            Para = new ParagraphStyle { LineHeightFactor = 1.7 },
        });
        set.SetLevel(OutlineLevel.H1, new OutlineLevelStyle
        {
            Char = new RunStyle { FontSizePt = 28, Bold = true },
            Para = new ParagraphStyle { SpaceBeforePt = 0, SpaceAfterPt = 16, Outline = OutlineLevel.H1 },
            BackgroundColor = "#F3F3F3",
        });
        set.SetLevel(OutlineLevel.H2, new OutlineLevelStyle
        {
            Char = new RunStyle { FontSizePt = 22 },
            Para = new ParagraphStyle { SpaceBeforePt = 16, SpaceAfterPt = 8, Outline = OutlineLevel.H2 },
            Border = new OutlineBorder { ShowBottom = true },
        });
        set.SetLevel(OutlineLevel.H3, new OutlineLevelStyle
        {
            Char = new RunStyle { FontSizePt = 17, Bold = true },
            Para = new ParagraphStyle { SpaceBeforePt = 12, SpaceAfterPt = 4, Outline = OutlineLevel.H3 },
        });
        set.SetLevel(OutlineLevel.H4, new OutlineLevelStyle
        {
            Char = new RunStyle { FontSizePt = 14 },
            Para = new ParagraphStyle { SpaceBeforePt = 8, Outline = OutlineLevel.H4 },
        });
        set.SetLevel(OutlineLevel.H5, new OutlineLevelStyle
        {
            Char = new RunStyle { FontSizePt = 12, Bold = true },
            Para = new ParagraphStyle { Outline = OutlineLevel.H5 },
        });
        set.SetLevel(OutlineLevel.H6, new OutlineLevelStyle
        {
            Char = new RunStyle { FontSizePt = 11, Italic = true },
            Para = new ParagraphStyle { Outline = OutlineLevel.H6 },
        });
        return set;
    }

    public static IReadOnlyList<OutlineStyleSet> BuiltInPresets { get; } =
        new[] { CreateDefault(), CreateAcademic(), CreateBusiness(), CreateModern() };

    // ── 기본값 (OutlineStyles 없을 때 fallback) ──────────────────

    public static OutlineLevelStyle DefaultForLevel(OutlineLevel level) => level switch
    {
        OutlineLevel.H1 => new OutlineLevelStyle { Char = new RunStyle { FontSizePt = 24, Bold = true } },
        OutlineLevel.H2 => new OutlineLevelStyle { Char = new RunStyle { FontSizePt = 20, Bold = true } },
        OutlineLevel.H3 => new OutlineLevelStyle { Char = new RunStyle { FontSizePt = 17, Bold = true } },
        OutlineLevel.H4 => new OutlineLevelStyle { Char = new RunStyle { FontSizePt = 15, Bold = true } },
        OutlineLevel.H5 => new OutlineLevelStyle { Char = new RunStyle { FontSizePt = 13, Bold = true } },
        OutlineLevel.H6 => new OutlineLevelStyle { Char = new RunStyle { FontSizePt = 12, Bold = true, Italic = true } },
        _               => new OutlineLevelStyle { Char = new RunStyle { FontSizePt = 11 } },
    };
}

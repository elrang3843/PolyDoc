using PolyDoc.Core;

namespace PolyDoc.Codecs.Hwpx;

/// <summary>
/// HWPX 의 header.xml 을 파싱한 결과. section.xml 의
/// paraPrIDRef / charPrIDRef / styleIDRef 가 가리키는 정의를 ID → PolyDoc 모델로 보관한다.
/// 한컴 오피스가 만든 hwpx 는 우리 codec 의 0~5 약속 ID 와 다른 임의 ID 를 쓰므로,
/// 이 컨텍스트를 통해 본문 paragraph/run 이 실제 서식을 회수한다.
/// </summary>
internal sealed class HwpxHeader
{
    public Dictionary<int, string> Fonts { get; } = new();
    public Dictionary<int, RunStyle> CharProperties { get; } = new();
    public Dictionary<int, ParagraphStyle> ParaProperties { get; } = new();
    public Dictionary<int, HwpxStyleDef> Styles { get; } = new();

    public RunStyle GetRunStyle(int? charPrId)
    {
        if (charPrId is { } id && CharProperties.TryGetValue(id, out var src))
        {
            return CloneRunStyle(src);
        }
        return new RunStyle();
    }

    public ParagraphStyle? GetParagraphStyle(int? paraPrId)
    {
        if (paraPrId is { } id && ParaProperties.TryGetValue(id, out var src))
        {
            return CloneParagraphStyle(src);
        }
        return null;
    }

    public HwpxStyleDef? GetStyle(int? styleId)
        => styleId is { } id && Styles.TryGetValue(id, out var s) ? s : null;

    public static RunStyle CloneRunStyle(RunStyle s) => new()
    {
        FontFamily = s.FontFamily,
        FontSizePt = s.FontSizePt,
        Bold = s.Bold,
        Italic = s.Italic,
        Underline = s.Underline,
        Strikethrough = s.Strikethrough,
        Overline = s.Overline,
        Superscript = s.Superscript,
        Subscript = s.Subscript,
        Foreground = s.Foreground,
        Background = s.Background,
        WidthPercent = s.WidthPercent,
        LetterSpacingPx = s.LetterSpacingPx,
    };

    public static ParagraphStyle CloneParagraphStyle(ParagraphStyle s) => new()
    {
        Alignment = s.Alignment,
        LineHeightFactor = s.LineHeightFactor,
        SpaceBeforePt = s.SpaceBeforePt,
        SpaceAfterPt = s.SpaceAfterPt,
        IndentFirstLineMm = s.IndentFirstLineMm,
        IndentLeftMm = s.IndentLeftMm,
        IndentRightMm = s.IndentRightMm,
        Outline = s.Outline,
        ListMarker = s.ListMarker is null ? null : new ListMarker
        {
            Kind = s.ListMarker.Kind,
            Level = s.ListMarker.Level,
            OrderedNumber = s.ListMarker.OrderedNumber,
        },
    };
}

/// <summary>HWPX style 정의 — paragraph 의 styleIDRef 가 가리킨다.</summary>
internal sealed class HwpxStyleDef
{
    public int? ParaPrIdRef { get; init; }
    public int? CharPrIdRef { get; init; }
    public string? Name { get; init; }
    public string? EngName { get; init; }
    public OutlineLevel Outline { get; init; } = OutlineLevel.Body;
}

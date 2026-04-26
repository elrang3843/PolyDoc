namespace PolyDoc.Core;

public sealed class Paragraph : Block
{
    public string? StyleId { get; set; }
    public ParagraphStyle Style { get; set; } = new();
    public IList<Run> Runs { get; set; } = new List<Run>();

    public string GetPlainText() => string.Concat(Runs.Select(r => r.Text));

    public Paragraph AddText(string text, RunStyle? style = null)
    {
        Runs.Add(new Run { Text = text, Style = style ?? new RunStyle() });
        return this;
    }

    public static Paragraph Of(string text, RunStyle? style = null)
    {
        var p = new Paragraph();
        p.AddText(text, style);
        return p;
    }
}

public sealed class ParagraphStyle
{
    public Alignment Alignment { get; set; } = Alignment.Left;
    public double LineHeightFactor { get; set; } = 1.2;
    public double SpaceBeforePt { get; set; }
    public double SpaceAfterPt { get; set; }
    public double IndentFirstLineMm { get; set; }
    public double IndentLeftMm { get; set; }
    public double IndentRightMm { get; set; }
    public OutlineLevel Outline { get; set; } = OutlineLevel.Body;
    public ListMarker? ListMarker { get; set; }
}

public enum Alignment
{
    Left,
    Center,
    Right,
    Justify,
    Distributed,
}

/// <summary>개요 수준. Body 는 본문, H1~H6 는 제목 단계.</summary>
public enum OutlineLevel
{
    Body = 0,
    H1 = 1,
    H2 = 2,
    H3 = 3,
    H4 = 4,
    H5 = 5,
    H6 = 6,
}

public sealed class ListMarker
{
    public ListKind Kind { get; set; } = ListKind.Bullet;
    public int Level { get; set; }
    public int? OrderedNumber { get; set; }
}

public enum ListKind
{
    Bullet,
    OrderedDecimal,
    OrderedAlpha,
    OrderedRoman,
}

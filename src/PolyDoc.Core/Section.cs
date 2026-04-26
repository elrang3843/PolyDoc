namespace PolyDoc.Core;

public sealed class Section
{
    public string? Id { get; set; }
    public PageSettings Page { get; set; } = new();
    public IList<Block> Blocks { get; set; } = new List<Block>();
}

/// <summary>편집 용지·여백·다단 등 페이지 기본 속성. 단위는 mm.</summary>
public sealed class PageSettings
{
    public double WidthMm { get; set; } = 210;        // A4
    public double HeightMm { get; set; } = 297;       // A4
    public double MarginTopMm { get; set; } = 25;
    public double MarginRightMm { get; set; } = 25;
    public double MarginBottomMm { get; set; } = 25;
    public double MarginLeftMm { get; set; } = 25;
    public int Columns { get; set; } = 1;
    public PageOrientation Orientation { get; set; } = PageOrientation.Portrait;
}

public enum PageOrientation
{
    Portrait,
    Landscape,
}

namespace PolyDoc.Core;

/// <summary>이름이 부여된 문단/문자 스타일 모음.</summary>
public sealed class StyleSheet
{
    public IDictionary<string, ParagraphStyle> ParagraphStyles { get; set; } = new Dictionary<string, ParagraphStyle>();
    public IDictionary<string, RunStyle> RunStyles { get; set; } = new Dictionary<string, RunStyle>();
}

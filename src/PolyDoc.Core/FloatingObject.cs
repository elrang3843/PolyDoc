namespace PolyDoc.Core;

/// <summary>
/// 텍스트 흐름 외부에 자유 위치로 배치되는 객체 (글상자·도형·이미지 캡션 등) 의 추상 베이스.
/// 좌표는 페이지 좌상단 원점, 단위는 mm. <see cref="ZOrder"/> 가 큰 값이 위에 그려진다.
/// 다형 직렬화는 <see cref="FloatingObjectJsonConverter"/> 가 처리.
/// </summary>
public abstract class FloatingObject
{
    public string? Id { get; set; }
    public double XMm { get; set; }
    public double YMm { get; set; }
    public double WidthMm { get; set; }
    public double HeightMm { get; set; }
    public int ZOrder { get; set; }
    public NodeStatus Status { get; set; } = NodeStatus.Clean;
}

/// <summary>글상자 모양.</summary>
public enum TextBoxShape
{
    Rectangle = 0,
    Speech    = 1,
    Cloud     = 2,
    Spiky     = 3,
    Lightning = 4,
}

/// <summary>글상자 내 텍스트 가로 정렬.</summary>
public enum TextBoxHAlign { Left = 0, Center = 1, Right = 2, Justify = 3 }

/// <summary>글상자 내 텍스트 세로 정렬.</summary>
public enum TextBoxVAlign { Top = 0, Middle = 1, Bottom = 2 }

public sealed class TextBoxObject : FloatingObject
{
    public TextBoxShape Shape { get; set; } = TextBoxShape.Rectangle;

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
    public TextBoxHAlign HAlign { get; set; } = TextBoxHAlign.Left;
    public TextBoxVAlign VAlign { get; set; } = TextBoxVAlign.Top;

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

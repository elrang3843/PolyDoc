namespace PolyDonky.Core;

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

    public Paragraph Clone() => new()
    {
        StyleId = StyleId,
        Style   = Style.Clone(),
        Runs    = Runs.Select(r => r.Clone()).ToList(),
    };
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

    /// <summary>인용 깊이. 0 = 일반 단락, ≥1 = 인용(blockquote) — Markdown 의 `>` 깊이.</summary>
    public int QuoteLevel { get; set; }

    /// <summary>코드 블록 언어 힌트. null = 일반 단락, "" = 언어 미지정 코드, "python"/"cs" 등 = 언어 코드.
    /// Markdown 펜스드 코드 블록의 info string 에 대응. non-null 이면 단락 전체가 코드 블록.</summary>
    public string? CodeLanguage { get; set; }

    /// <summary>코드 블록 줄 번호 표시. <see cref="CodeLanguage"/> 가 non-null 인 경우에만 유의미.</summary>
    public bool ShowLineNumbers { get; set; }

    /// <summary>강제 페이지 나누기. true 이면 이 단락 앞에 페이지 브레이크를 삽입한다.
    /// DOCX: w:pageBreakBefore, HWPX: hp:p pageBreak="1"</summary>
    public bool ForcePageBreakBefore { get; set; }

    /// <summary>단락 아래 경계선 두께(pt). 0이면 없음. CSS border-bottom / OutlineStyle 경계선에 대응.</summary>
    public double BorderBottomPt { get; set; }

    /// <summary>단락 아래 경계선 색상(hex, 예: "#CCCCCC"). null이면 기본 색.</summary>
    public string? BorderBottomColor { get; set; }

    /// <summary>모든 필드를 복사한 깊은 복제본 — ListMarker 도 새 인스턴스로.</summary>
    public ParagraphStyle Clone() => new()
    {
        Alignment              = Alignment,
        LineHeightFactor       = LineHeightFactor,
        SpaceBeforePt          = SpaceBeforePt,
        SpaceAfterPt           = SpaceAfterPt,
        IndentFirstLineMm      = IndentFirstLineMm,
        IndentLeftMm           = IndentLeftMm,
        IndentRightMm          = IndentRightMm,
        Outline                = Outline,
        ListMarker             = ListMarker?.Clone(),
        QuoteLevel             = QuoteLevel,
        CodeLanguage           = CodeLanguage,
        ShowLineNumbers        = ShowLineNumbers,
        ForcePageBreakBefore   = ForcePageBreakBefore,
        BorderBottomPt         = BorderBottomPt,
        BorderBottomColor      = BorderBottomColor,
    };
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

    /// <summary>중첩 깊이. 0 = 최상위. Markdown 들여쓰기 / 트리 형 리스트에 사용.</summary>
    public int Level { get; set; }

    public int? OrderedNumber { get; set; }

    /// <summary>GFM 작업 목록(task list) 체크 상태. null = 작업 목록 아님, true = `[x]`, false = `[ ]`.</summary>
    public bool? Checked { get; set; }

    /// <summary>알파벳/로마자 대소문자. null = 자동(L0=대문자, L≥1=소문자),
    /// true = 항상 대문자, false = 항상 소문자. Decimal/Bullet 에서는 무시.
    /// HTML <c>&lt;ol type="A"/"a"/"I"/"i"&gt;</c> 의 type 속성을 보존한다.</summary>
    public bool? UpperCase { get; set; }

    /// <summary>모든 필드를 복사한 깊은 복제본.</summary>
    public ListMarker Clone() => new()
    {
        Kind          = Kind,
        Level         = Level,
        OrderedNumber = OrderedNumber,
        Checked       = Checked,
        UpperCase     = UpperCase,
    };
}

public enum ListKind
{
    Bullet,
    OrderedDecimal,
    OrderedAlpha,
    OrderedRoman,
}

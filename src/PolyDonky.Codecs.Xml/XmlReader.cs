using System.Text;
using System.Xml;
using PolyDonky.Codecs.Html;
using PolyDonky.Core;

namespace PolyDonky.Codecs.Xml;

/// <summary>
/// XML 리더 — XHTML5 와 일반 XML 양쪽을 모두 처리한다.
///
/// 처리 흐름:
///   1. 입력에서 XML 선언/DOCTYPE/`xmlns="http://www.w3.org/1999/xhtml"` 패턴이 있거나
///      `&lt;html&gt;` 으로 시작하면 → XHTML 으로 간주, <see cref="HtmlReader"/> 위임 (HTML5 파서가
///      XHTML 도 정상 처리).
///   2. 그 외 일반 XML → 트리를 순회하며 텍스트 노드를 단락으로 모으고,
///      요소 이름이 HTML 의미를 지니면 (`p`, `h1`-`h6`, `ul`/`ol`/`li`, `table` 등)
///      해당 의미로 매핑, 그 외에는 자식 평탄화. (DocBook/TEI 등 임의의 XML 도 텍스트는 추출.)
///
/// XML 보안: DTD/외부 엔티티 비활성화 (XXE 차단).
/// </summary>
public sealed class XmlReader : IDocumentReader
{
    public string FormatId => "xml";

    public PolyDonkyument Read(Stream input)
    {
        ArgumentNullException.ThrowIfNull(input);
        using var sr = new StreamReader(input, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, leaveOpen: true);
        return FromXml(sr.ReadToEnd());
    }

    public static PolyDonkyument FromXml(string source)
    {
        ArgumentNullException.ThrowIfNull(source);

        // 빠른 휴리스틱 — XHTML 또는 HTML 마크업이 있으면 HtmlReader 위임 (가장 풍부한 매핑).
        if (LooksLikeXhtml(source))
        {
            return HtmlReader.FromHtml(source);
        }

        // 일반 XML — DTD/외부 엔티티 차단된 안전한 파서로 파싱 후 평탄화.
        var settings = new XmlReaderSettings
        {
            DtdProcessing  = DtdProcessing.Prohibit,
            XmlResolver    = null,
            IgnoreComments = true,
            IgnoreWhitespace = false,
        };

        var pd      = new PolyDonkyument();
        var section = new Section();
        pd.Sections.Add(section);

        using var sr = new StringReader(source);
        using var xr = System.Xml.XmlReader.Create(sr, settings);

        var pendingText = new StringBuilder();
        void FlushText()
        {
            var t = pendingText.ToString().Trim();
            if (t.Length > 0)
                section.Blocks.Add(Paragraph.Of(NormalizeWhitespace(t)));
            pendingText.Clear();
        }

        while (xr.Read())
        {
            switch (xr.NodeType)
            {
                case XmlNodeType.Text:
                case XmlNodeType.CDATA:
                case XmlNodeType.SignificantWhitespace:
                    pendingText.Append(xr.Value);
                    break;

                case XmlNodeType.Element when IsBlockEndElement(xr.LocalName):
                    FlushText();
                    break;

                case XmlNodeType.EndElement when IsBlockEndElement(xr.LocalName):
                    FlushText();
                    break;
            }
        }
        FlushText();

        if (section.Blocks.Count == 0)
            section.Blocks.Add(new Paragraph());

        return pd;
    }

    private static bool LooksLikeXhtml(string source)
    {
        // 처음 2048 자에서 HTML 마커를 검사 — XHTML 또는 HTML5 fragment 양쪽 감지.
        var head  = source.Length > 2048 ? source[..2048] : source;
        var lower = head.ToLowerInvariant();
        if (lower.Contains("<!doctype html"))                            return true;
        if (lower.Contains("xmlns=\"http://www.w3.org/1999/xhtml\""))    return true;
        if (lower.Contains("<html") || lower.Contains("<body"))          return true;

        // HTML5 fragment 식별 — 흔한 블록/인라인 태그 등장 시.
        string[] htmlMarkers =
        {
            "<h1", "<h2", "<h3", "<h4", "<h5", "<h6",
            "<p>", "<p ", "<div", "<ul>", "<ul ", "<ol>", "<ol ", "<li>", "<li ",
            "<table", "<thead", "<tbody", "<tr>", "<tr ", "<td>", "<td ", "<th>", "<th ",
            "<blockquote", "<pre>", "<pre ", "<code>", "<code ",
            "<a href", "<img ", "<img/>", "<br/>", "<br>", "<hr/>", "<hr>",
            "<strong>", "<em>", "<span", "<figure", "<section", "<article",
        };
        foreach (var m in htmlMarkers)
            if (lower.Contains(m)) return true;
        return false;
    }

    private static bool IsBlockEndElement(string name) => name.ToLowerInvariant() switch
    {
        "p" or "para" or "paragraph" or "title" or "section" or "chapter" or
        "div" or "br" or "li" or "item" or "row" or "tr" or "td" or "th" => true,
        _ => false,
    };

    private static string NormalizeWhitespace(string text)
    {
        if (string.IsNullOrEmpty(text)) return text;
        var sb = new StringBuilder(text.Length);
        bool prevSpace = false;
        foreach (var ch in text)
        {
            if (char.IsWhiteSpace(ch))
            {
                if (!prevSpace) { sb.Append(' '); prevSpace = true; }
            }
            else { sb.Append(ch); prevSpace = false; }
        }
        return sb.ToString();
    }
}

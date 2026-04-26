using System.Globalization;
using System.Xml.Linq;
using PolyDoc.Core;

namespace PolyDoc.Codecs.Hwpx;

/// <summary>
/// HWPX 의 header.xml 을 PolyDoc 모델로 매핑하는 파서.
/// 한컴 오피스가 만든 hwpx 의 임의 ID 매핑과 우리 codec 의 0~5 약속 매핑 모두 같은 경로로 동작.
///
/// 단위 변환 메모:
///   - charPr.height: 0.01 pt 단위 (예: 1000 = 10pt)
///   - paraPr.margin: hwpunit (1/7200 inch) — left/right/intent 를 mm 로 환산
///   - 색상: "#RRGGBB" 또는 "RRGGBB"
/// </summary>
internal static class HwpxHeaderReader
{
    private const double HwpUnitToMm = 25.4 / 7200.0;

    public static HwpxHeader Parse(XDocument? doc)
    {
        var header = new HwpxHeader();
        if (doc?.Root is null)
        {
            return header;
        }

        // 한컴 hwpx 의 element 위치는 변종이 있어 LocalName 으로 descendants 전수 매칭.
        foreach (var elem in doc.Root.Descendants())
        {
            switch (elem.Name.LocalName)
            {
                case "font":
                    ParseFont(elem, header);
                    break;
                case "charPr":
                    ParseCharPr(elem, header);
                    break;
                case "paraPr":
                    ParseParaPr(elem, header);
                    break;
                case "style":
                    ParseStyle(elem, header);
                    break;
            }
        }

        return header;
    }

    private static void ParseFont(XElement elem, HwpxHeader header)
    {
        if (TryGetIntAttr(elem, "id", out var id)
            && elem.Attribute("face")?.Value is { Length: > 0 } face)
        {
            header.Fonts[id] = face;
        }
    }

    private static void ParseCharPr(XElement elem, HwpxHeader header)
    {
        if (!TryGetIntAttr(elem, "id", out var id))
        {
            return;
        }

        var rs = new RunStyle();

        // height: 0.01 pt 단위
        if (TryGetIntAttr(elem, "height", out var height) && height > 0)
        {
            rs.FontSizePt = height / 100.0;
        }

        TryParseColor(elem.Attribute("textColor")?.Value, c => rs.Foreground = c);
        TryParseColor(elem.Attribute("shadeColor")?.Value, c => rs.Background = c);

        // 자식 element 로 강조/밑줄/취소선 등
        foreach (var child in elem.Elements())
        {
            switch (child.Name.LocalName)
            {
                case "bold":
                    rs.Bold = true;
                    break;
                case "italic":
                    rs.Italic = true;
                    break;
                case "underline":
                    var ut = child.Attribute("type")?.Value;
                    if (string.IsNullOrEmpty(ut) || !string.Equals(ut, "NONE", StringComparison.OrdinalIgnoreCase))
                    {
                        rs.Underline = true;
                    }
                    break;
                case "strikeout":
                    var ss = child.Attribute("shape")?.Value;
                    if (string.IsNullOrEmpty(ss) || !string.Equals(ss, "NONE", StringComparison.OrdinalIgnoreCase))
                    {
                        rs.Strikethrough = true;
                    }
                    break;
                case "supscript":     // HWPX 사양 표기 (super 가 아닌 sup)
                case "supScript":
                case "superscript":
                    rs.Superscript = true;
                    break;
                case "subscript":
                case "subScript":
                    rs.Subscript = true;
                    break;
                case "fontRef":
                    // 한글 우선, 다음 라틴/한자 순으로 폰트 매핑.
                    if (TryResolveFontFamily(child, header, out var family))
                    {
                        rs.FontFamily = family;
                    }
                    break;
            }
        }

        header.CharProperties[id] = rs;
    }

    private static bool TryResolveFontFamily(XElement fontRef, HwpxHeader header, out string family)
    {
        foreach (var attr in new[] { "hangul", "latin", "hanja", "japanese", "other", "symbol", "user" })
        {
            if (TryGetIntAttr(fontRef, attr, out var fid)
                && header.Fonts.TryGetValue(fid, out var face))
            {
                family = face;
                return true;
            }
        }
        family = string.Empty;
        return false;
    }

    private static void ParseParaPr(XElement elem, HwpxHeader header)
    {
        if (!TryGetIntAttr(elem, "id", out var id))
        {
            return;
        }

        var ps = new ParagraphStyle();

        var align = elem.Elements().FirstOrDefault(c => c.Name.LocalName == "align");
        if (align?.Attribute("horizontal")?.Value is { } ha)
        {
            ps.Alignment = ha.ToUpperInvariant() switch
            {
                "CENTER" => Alignment.Center,
                "RIGHT" => Alignment.Right,
                "JUSTIFY" => Alignment.Justify,
                "DISTRIBUTE" => Alignment.Distributed,
                _ => Alignment.Left,
            };
        }

        var margin = elem.Elements().FirstOrDefault(c => c.Name.LocalName == "margin");
        if (margin is not null)
        {
            if (TryGetDoubleAttr(margin, "left", out var lm)) ps.IndentLeftMm = lm * HwpUnitToMm;
            if (TryGetDoubleAttr(margin, "right", out var rm)) ps.IndentRightMm = rm * HwpUnitToMm;
            if (TryGetDoubleAttr(margin, "intent", out var im)) ps.IndentFirstLineMm = im * HwpUnitToMm;
            if (TryGetDoubleAttr(margin, "prev", out var prev)) ps.SpaceBeforePt = prev / 100.0;     // hwpunit 가 아닌 pt*100 (사양 변종)
            if (TryGetDoubleAttr(margin, "next", out var next)) ps.SpaceAfterPt = next / 100.0;
        }

        var lineSpacing = elem.Elements().FirstOrDefault(c => c.Name.LocalName == "lineSpacing");
        if (lineSpacing is not null
            && string.Equals(lineSpacing.Attribute("type")?.Value, "PERCENT", StringComparison.OrdinalIgnoreCase)
            && TryGetDoubleAttr(lineSpacing, "value", out var lsv) && lsv > 0)
        {
            ps.LineHeightFactor = lsv / 100.0;
        }

        header.ParaProperties[id] = ps;
    }

    private static void ParseStyle(XElement elem, HwpxHeader header)
    {
        if (!TryGetIntAttr(elem, "id", out var id))
        {
            return;
        }

        var name = elem.Attribute("name")?.Value;
        var engName = elem.Attribute("engName")?.Value;

        // engName 또는 name 이 "Heading{N}" / "개요{N}" 인 경우 outline 레벨 추정.
        var outline = OutlineLevel.Body;
        var key = engName ?? name ?? string.Empty;
        if (key.StartsWith("Heading", StringComparison.OrdinalIgnoreCase)
            && int.TryParse(key.AsSpan("Heading".Length), out var lvl)
            && lvl is >= 1 and <= 6)
        {
            outline = (OutlineLevel)lvl;
        }
        else if (key.StartsWith("개요", StringComparison.Ordinal)
            && int.TryParse(key.AsSpan("개요".Length), out lvl)
            && lvl is >= 1 and <= 6)
        {
            outline = (OutlineLevel)lvl;
        }

        header.Styles[id] = new HwpxStyleDef
        {
            ParaPrIdRef = TryGetIntAttr(elem, "paraPrIDRef", out var pp) ? pp : null,
            CharPrIdRef = TryGetIntAttr(elem, "charPrIDRef", out var cp) ? cp : null,
            Name = name,
            EngName = engName,
            Outline = outline,
        };
    }

    private static bool TryGetIntAttr(XElement elem, string name, out int value)
    {
        var raw = elem.Attribute(name)?.Value;
        if (!string.IsNullOrEmpty(raw)
            && int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out value))
        {
            return true;
        }
        value = 0;
        return false;
    }

    private static bool TryGetDoubleAttr(XElement elem, string name, out double value)
    {
        var raw = elem.Attribute(name)?.Value;
        if (!string.IsNullOrEmpty(raw)
            && double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out value))
        {
            return true;
        }
        value = 0;
        return false;
    }

    private static void TryParseColor(string? raw, Action<Color> assign)
    {
        if (string.IsNullOrEmpty(raw)) return;
        var hex = raw.StartsWith('#') ? raw : "#" + raw;
        if (hex.Length is not (7 or 9)) return;
        try
        {
            assign(Color.FromHex(hex));
        }
        catch (FormatException)
        {
            // 잘못된 색상 표기는 무시.
        }
    }
}

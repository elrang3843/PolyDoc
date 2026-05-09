using System.Globalization;
using System.Text;
using PolyDonky.Core;
using System.Linq;

namespace PolyDonky.Codecs.Xml;

/// <summary>
/// XML 작성기 — XHTML5 (Polyglot Markup) 직렬화.
///
/// HTML5 와 XHTML5 양쪽 파서가 동일하게 해석할 수 있는 마크업으로 출력한다:
///   - XML 선언: &lt;?xml version="1.0" encoding="utf-8"?&gt;
///   - DOCTYPE: &lt;!DOCTYPE html&gt;
///   - 루트 namespace: &lt;html xmlns="http://www.w3.org/1999/xhtml" lang="ko"&gt;
///   - void 요소(br, hr, img, meta, input)는 self-closing &lt;br/&gt; 형태
///   - 모든 속성 값은 큰따옴표 인용
///   - 모든 태그/속성 이름은 소문자
///   - 텍스트는 &amp; &lt; &gt; 이스케이프, 속성은 추가로 " ' 이스케이프
///
/// HtmlWriter 와 같은 매핑(헤딩, 리스트, 표, 링크, 이미지, 인용, 코드, 작업 목록…) 을 사용하되
/// 출력만 XML 규정에 맞게 엄격화한다. PolyDonky 의 XML 형식은 IWPF 의 인간 친화적 단일 파일 대안이다.
/// </summary>
public sealed class XmlWriter : IDocumentWriter
{
    public string FormatId => "xml";

    public Encoding Encoding { get; init; } = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

    /// <summary>완전한 XHTML 문서로 출력할지 (기본 true) — false 면 fragment.</summary>
    public bool FullDocument { get; init; } = true;

    /// <summary>문서 제목. null 이면 첫 H1 텍스트 사용.</summary>
    public string? DocumentTitle { get; init; }

    public void Write(PolyDonkyument document, Stream output)
    {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(output);

        using var w = new StreamWriter(output, Encoding, leaveOpen: true) { NewLine = "\n" };
        w.Write(ToXml(document, FullDocument, DocumentTitle));
    }

    public static string ToXml(PolyDonkyument document, bool fullDocument = true, string? title = null)
    {
        ArgumentNullException.ThrowIfNull(document);
        var sb = new StringBuilder();
        var notes = BuildNoteNums(document);
        var indent = fullDocument ? "  " : "";

        if (fullDocument)
        {
            var docTitle = title ?? document.EnumerateParagraphs()
                .FirstOrDefault(p => p.Style.Outline == OutlineLevel.H1)?.GetPlainText()
                ?? "PolyDonky 문서";

            var page = document.Sections.Count > 0 ? document.Sections[0].Page : new PageSettings();

            sb.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>\n");
            sb.Append("<!DOCTYPE html>\n");
            sb.Append("<html xmlns=\"http://www.w3.org/1999/xhtml\" lang=\"ko\">\n");
            sb.Append("<head>\n");
            sb.Append("  <meta charset=\"utf-8\"/>\n");
            sb.Append("  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\"/>\n");
            sb.Append("  <meta name=\"generator\" content=\"PolyDonky\"/>\n");
            sb.Append("  <title>").Append(EscapeText(docTitle)).Append("</title>\n");
            // 편집용지 메타 태그.
            sb.Append("  <meta name=\"pd-page-size\" content=\"").Append(page.SizeKind).Append("\"/>\n");
            sb.Append("  <meta name=\"pd-page-orientation\" content=\"")
              .Append(page.Orientation == PageOrientation.Landscape ? "landscape" : "portrait")
              .Append("\"/>\n");
            if (page.SizeKind == PaperSizeKind.Custom)
            {
                sb.Append("  <meta name=\"pd-page-width\" content=\"")
                  .Append(page.WidthMm.ToString("0.##", CultureInfo.InvariantCulture)).Append("mm\"/>\n");
                sb.Append("  <meta name=\"pd-page-height\" content=\"")
                  .Append(page.HeightMm.ToString("0.##", CultureInfo.InvariantCulture)).Append("mm\"/>\n");
            }
            WriteStyleBlock(sb, document.Styles, page);
            sb.Append("</head>\n");
            sb.Append("<body>\n");
        }

        foreach (var section in document.Sections)
            WriteBlocks(sb, section.Blocks, indent, notes);

        if (fullDocument)
        {
            if (notes.HasNotes)
                WriteNoteSections(sb, document, notes, indent);
            sb.Append("</body>\n</html>\n");
        }

        return sb.ToString();
    }

    private sealed record NoteNums(
        IReadOnlyDictionary<string, int> Footnotes,
        IReadOnlyDictionary<string, int> Endnotes)
    {
        public bool HasNotes => Footnotes.Count > 0 || Endnotes.Count > 0;

        public static readonly NoteNums Empty = new(
            new Dictionary<string, int>(),
            new Dictionary<string, int>());
    }

    private static NoteNums BuildNoteNums(PolyDonkyument doc)
    {
        if (doc.Footnotes.Count == 0 && doc.Endnotes.Count == 0) return NoteNums.Empty;
        return new NoteNums(
            doc.Footnotes.Select((f, i) => (f.Id, i + 1)).ToDictionary(x => x.Id, x => x.Item2),
            doc.Endnotes.Select((e, i) => (e.Id, i + 1)).ToDictionary(x => x.Id, x => x.Item2));
    }

    private static void WriteNoteSections(StringBuilder sb, PolyDonkyument doc, NoteNums notes, string indent)
    {
        if (doc.Footnotes.Count > 0)
        {
            sb.Append(indent).Append("<section class=\"footnotes\">\n");
            sb.Append(indent).Append("  <hr/>\n");
            sb.Append(indent).Append("  <ol>\n");
            foreach (var entry in doc.Footnotes)
            {
                if (!notes.Footnotes.TryGetValue(entry.Id, out var num)) continue;
                sb.Append(indent).Append("    <li id=\"fn-").Append(num).Append("\">");
                sb.Append(RenderBlocks(entry.Blocks, notes));
                sb.Append(" <a href=\"#fnref-").Append(num).Append("\">&#8617;</a>");
                sb.Append("</li>\n");
            }
            sb.Append(indent).Append("  </ol>\n");
            sb.Append(indent).Append("</section>\n");
        }

        if (doc.Endnotes.Count > 0)
        {
            sb.Append(indent).Append("<section class=\"endnotes\">\n");
            sb.Append(indent).Append("  <hr/>\n");
            sb.Append(indent).Append("  <ol>\n");
            foreach (var entry in doc.Endnotes)
            {
                if (!notes.Endnotes.TryGetValue(entry.Id, out var num)) continue;
                sb.Append(indent).Append("    <li id=\"en-").Append(num).Append("\">");
                sb.Append(RenderBlocks(entry.Blocks, notes));
                sb.Append(" <a href=\"#enref-").Append(num).Append("\">&#8617;</a>");
                sb.Append("</li>\n");
            }
            sb.Append(indent).Append("  </ol>\n");
            sb.Append(indent).Append("</section>\n");
        }
    }

    private static string RenderBlocks(IList<Block> blocks, NoteNums notes)
    {
        var sb = new StringBuilder();
        bool first = true;
        foreach (var b in blocks)
        {
            if (b is Paragraph p)
            {
                if (!first) sb.Append(' ');
                sb.Append(RenderRuns(p.Runs, notes));
                first = false;
            }
        }
        return sb.ToString();
    }

    // ── 블록 렌더링 ───────────────────────────────────────────────────

    private static void WriteBlocks(StringBuilder sb, IList<Block> blocks, string indent, NoteNums? notes = null)
    {
        int i = 0;
        while (i < blocks.Count)
        {
            var b = blocks[i];

            int qLvl = QuoteLevelOf(b);
            if (qLvl > 0)
            {
                int j = i;
                while (j < blocks.Count && QuoteLevelOf(blocks[j]) >= qLvl) j++;
                var inner = blocks.Skip(i).Take(j - i).Select(StripQuoteLevel).ToList();
                sb.Append(indent).Append("<blockquote>\n");
                WriteBlocks(sb, inner, indent + "  ", notes);
                sb.Append(indent).Append("</blockquote>\n");
                i = j;
                continue;
            }

            if (b is Paragraph p && p.Style.ListMarker is { } lm0)
            {
                int j = i;
                while (j < blocks.Count
                       && blocks[j] is Paragraph pj
                       && pj.Style.ListMarker is { } lmj
                       && lmj.Kind == lm0.Kind
                       && lmj.Level == lm0.Level)
                    j++;
                WriteListGroup(sb, blocks, i, j, indent, notes);
                i = j;
                continue;
            }

            switch (b)
            {
                case ThematicBreakBlock thb:
                {
                    var hrStyle = new List<string>();
                    if (thb.LineColor is not null)
                        hrStyle.Add($"border-top:1px solid {thb.LineColor}");
                    if (thb.MarginPt > 0)
                    {
                        var marginPx = thb.MarginPt * 96.0 / 72.0;
                        hrStyle.Add($"margin:{marginPx:F0}px 0");
                    }
                    var styleAttr = hrStyle.Count > 0 ? $" style=\"{string.Join(';', hrStyle)}\"" : "";
                    sb.Append(indent).Append($"<hr{styleAttr}/>\n");
                    break;
                }
                case Paragraph para:     WriteParagraph(sb, para, indent, notes); break;
                case Table table:        WriteTable(sb, table, indent, notes);    break;
                case ImageBlock img:     WriteImage(sb, img, indent);             break;
                case TocBlock toc:       WriteToc(sb, toc, indent);               break;
                case ContainerBlock box: WriteContainer(sb, box, indent, notes);  break;
                case ShapeObject shape:  WriteShape(sb, shape, indent);           break;
                case TextBoxObject tbox: WriteTextBox(sb, tbox, indent, notes);   break;
                case OpaqueBlock opq:    WriteOpaque(sb, opq, indent);            break;
            }
            i++;
        }
    }

    private static int QuoteLevelOf(Block b) => b is Paragraph p ? p.Style.QuoteLevel : 0;

    private static Block StripQuoteLevel(Block b)
    {
        if (b is not Paragraph p) return b;
        var style = p.Style.Clone();
        style.QuoteLevel = Math.Max(0, p.Style.QuoteLevel - 1);
        return new Paragraph { StyleId = p.StyleId, Style = style, Runs = p.Runs };
    }

    private static void WriteParagraph(StringBuilder sb, Paragraph p, string indent, NoteNums? notes = null)
    {
        if (p.Style.CodeLanguage is not null)
        {
            var langAttr = p.Style.CodeLanguage.Length > 0
                ? $" class=\"language-{EscapeAttr(p.Style.CodeLanguage)}\""
                : "";
            sb.Append(indent).Append("<pre><code").Append(langAttr).Append('>')
              .Append(EscapeText(p.GetPlainText()))
              .Append("</code></pre>\n");
            return;
        }

        var classAttr = BuildClassAttr(p.StyleId);

        if (p.Style.Outline > OutlineLevel.Body)
        {
            int lvl = (int)p.Style.Outline;
            sb.Append(indent).Append('<').Append('h').Append(lvl).Append(classAttr).Append(ParagraphStyleAttr(p.Style)).Append('>');
            sb.Append(RenderRuns(p.Runs, notes));
            sb.Append("</h").Append(lvl).Append(">\n");
            return;
        }

        sb.Append(indent).Append("<p").Append(classAttr).Append(ParagraphStyleAttr(p.Style)).Append('>');
        sb.Append(RenderRuns(p.Runs, notes));
        sb.Append("</p>\n");
    }

    private static List<string> BuildParagraphCssParts(ParagraphStyle s)
    {
        var parts = new List<string>(6);
        switch (s.Alignment)
        {
            case Alignment.Center:  parts.Add("text-align:center");  break;
            case Alignment.Right:   parts.Add("text-align:right");   break;
            case Alignment.Justify: parts.Add("text-align:justify"); break;
        }
        if (s.LineHeightFactor > 0 && Math.Abs(s.LineHeightFactor - 1.0) > 0.001)
            parts.Add($"line-height:{s.LineHeightFactor.ToString("0.##", CultureInfo.InvariantCulture)}");
        if (s.SpaceBeforePt > 0)
            parts.Add($"margin-top:{s.SpaceBeforePt.ToString("0.##", CultureInfo.InvariantCulture)}pt");
        if (s.SpaceAfterPt > 0)
            parts.Add($"margin-bottom:{s.SpaceAfterPt.ToString("0.##", CultureInfo.InvariantCulture)}pt");
        if (s.IndentFirstLineMm > 0)
            parts.Add($"text-indent:{FmtMm(s.IndentFirstLineMm)}");
        if (s.IndentLeftMm > 0)
            parts.Add($"padding-left:{FmtMm(s.IndentLeftMm)}");
        if (s.IndentRightMm > 0)
            parts.Add($"padding-right:{FmtMm(s.IndentRightMm)}");

        if (s.ForcePageBreakBefore)
            parts.Add("page-break-before:always");

        return parts;
    }

    private static string ParagraphStyleAttr(ParagraphStyle s)
    {
        var parts = BuildParagraphCssParts(s);
        return parts.Count == 0 ? "" : $" style=\"{string.Join(';', parts)}\"";
    }

    private static void WriteStyleBlock(StringBuilder sb, StyleSheet styles, PageSettings? page = null)
    {
        bool hasStyles = styles.ParagraphStyles.Count > 0;
        if (!hasStyles && page is null) return;

        sb.Append("  <style>\n");

        if (page is not null)
        {
            var w  = page.EffectiveWidthMm.ToString("0.##",  CultureInfo.InvariantCulture);
            var h  = page.EffectiveHeightMm.ToString("0.##", CultureInfo.InvariantCulture);
            var mt = FmtMm(page.MarginTopMm);
            var mr = FmtMm(page.MarginRightMm);
            var mb = FmtMm(page.MarginBottomMm);
            var ml = FmtMm(page.MarginLeftMm);
            sb.Append("    @page {\n");
            sb.Append("      size: ").Append(w).Append("mm ").Append(h).Append("mm;\n");
            sb.Append("      margin: ").Append(mt).Append(' ').Append(mr).Append(' ').Append(mb).Append(' ').Append(ml).Append(";\n");
            sb.Append("    }\n");
        }

        foreach (var (id, ps) in styles.ParagraphStyles)
        {
            var parts = BuildParagraphCssParts(ps);
            if (parts.Count > 0)
                sb.Append("    .pd-").Append(EscapeCssIdent(id))
                  .Append(" { ").Append(string.Join("; ", parts)).Append("; }\n");
        }
        sb.Append("  </style>\n");
    }

    private static string BuildClassAttr(string? styleId)
        => styleId is { Length: > 0 } sid ? $" class=\"pd-{EscapeCssIdent(sid)}\"" : "";

    private static string EscapeCssIdent(string id)
    {
        if (string.IsNullOrEmpty(id)) return "x";
        var sb = new StringBuilder(id.Length + 1);
        foreach (var ch in id)
        {
            if (char.IsLetterOrDigit(ch) || ch == '-') sb.Append(ch);
            else sb.Append('_');
        }
        if (char.IsDigit(sb[0])) sb.Insert(0, '_');
        return sb.ToString();
    }

    private static string FmtMm(double mm) =>
        mm > 0
            ? $"{mm.ToString("0.##", CultureInfo.InvariantCulture)}mm"
            : "0";

    private static void WriteListGroup(StringBuilder sb, IList<Block> blocks, int from, int to, string indent, NoteNums? notes = null)
    {
        var p0  = (Paragraph)blocks[from];
        var lm  = p0.Style.ListMarker!;
        var tag = lm.Kind == ListKind.Bullet ? "ul" : "ol";
        var startAttr = "";
        if (lm.Kind != ListKind.Bullet && lm.OrderedNumber is { } start && start != 1)
            startAttr = $" start=\"{start}\"";

        sb.Append(indent).Append('<').Append(tag).Append(startAttr).Append(">\n");
        var inner = indent + "  ";

        for (int k = from; k < to; k++)
        {
            var p = (Paragraph)blocks[k];
            var marker = p.Style.ListMarker!;
            sb.Append(inner).Append("<li>");
            if (marker.Checked.HasValue)
            {
                var ck = marker.Checked.Value ? " checked=\"checked\"" : "";
                sb.Append("<input type=\"checkbox\" disabled=\"disabled\"").Append(ck).Append("/> ");
            }
            sb.Append(RenderRuns(p.Runs, notes));
            sb.Append("</li>\n");
        }

        sb.Append(indent).Append("</").Append(tag).Append(">\n");
    }

    private static void WriteTable(StringBuilder sb, Table t, string indent, NoteNums? notes = null)
    {
        if (t.Rows.Count == 0) return;

        var tableStyle = BuildTableStyle(t);
        sb.Append(indent).Append("<table");
        if (!string.IsNullOrEmpty(tableStyle))
            sb.Append(" style=\"").Append(tableStyle).Append('"');
        sb.Append(">\n");

        if (!string.IsNullOrEmpty(t.Caption))
            sb.Append(indent).Append("  <caption>")
              .Append(EscapeText(t.Caption))
              .Append("</caption>\n");

        // <colgroup> 출력 — 컬럼 너비 보존.
        if (t.Columns.Any(c => c.WidthMm > 0))
        {
            sb.Append(indent).Append("  <colgroup>\n");
            foreach (var col in t.Columns)
            {
                sb.Append(indent).Append("    <col");
                if (col.WidthMm > 0)
                    sb.Append(" style=\"width:").Append(FmtMm(col.WidthMm)).Append('"');
                sb.Append("/>\n");
            }
            sb.Append(indent).Append("  </colgroup>\n");
        }

        var headerRows = t.Rows.Where(r => r.IsHeader).ToList();
        var bodyRows   = t.Rows.Where(r => !r.IsHeader).ToList();

        if (headerRows.Count > 0)
        {
            sb.Append(indent).Append("  <thead>\n");
            foreach (var r in headerRows) WriteRow(sb, r, t, indent + "    ", isHeader: true, notes);
            sb.Append(indent).Append("  </thead>\n");
        }
        if (bodyRows.Count > 0)
        {
            sb.Append(indent).Append("  <tbody>\n");
            foreach (var r in bodyRows) WriteRow(sb, r, t, indent + "    ", isHeader: false, notes);
            sb.Append(indent).Append("  </tbody>\n");
        }

        sb.Append(indent).Append("</table>\n");
    }

    private static string BuildTableStyle(Table t)
    {
        var parts = new List<string>(4);
        parts.Add("border-collapse:collapse");
        if (!string.IsNullOrEmpty(t.BackgroundColor))
            parts.Add($"background-color:{t.BackgroundColor}");
        switch (t.HAlign)
        {
            case TableHAlign.Center: parts.Add("margin-left:auto");  parts.Add("margin-right:auto"); break;
            case TableHAlign.Right:  parts.Add("margin-left:auto");  break;
        }
        return string.Join(';', parts);
    }

    private static void WriteRow(StringBuilder sb, TableRow row, Table t, string indent, bool isHeader, NoteNums? notes = null)
    {
        sb.Append(indent).Append("<tr>\n");
        foreach (var cell in row.Cells)
        {
            var tag = isHeader ? "th" : "td";
            var attrs = new StringBuilder();
            if (cell.ColumnSpan > 1) attrs.Append(" colspan=\"").Append(cell.ColumnSpan).Append('"');
            if (cell.RowSpan    > 1) attrs.Append(" rowspan=\"").Append(cell.RowSpan).Append('"');

            var cellStyle = BuildCellStyle(cell, t);
            if (!string.IsNullOrEmpty(cellStyle))
                attrs.Append(" style=\"").Append(cellStyle).Append('"');

            sb.Append(indent).Append("  <").Append(tag).Append(attrs).Append('>');
            bool first = true;
            foreach (var b in cell.Blocks)
            {
                if (b is Paragraph p)
                {
                    if (!first) sb.Append("<br/>");
                    sb.Append(RenderRuns(p.Runs, notes));
                    first = false;
                }
            }
            sb.Append("</").Append(tag).Append(">\n");
        }
        sb.Append(indent).Append("</tr>\n");
    }

    private static string BuildCellStyle(TableCell cell, Table t)
    {
        var parts = new List<string>(8);
        switch (cell.TextAlign)
        {
            case CellTextAlign.Center:  parts.Add("text-align:center");  break;
            case CellTextAlign.Right:   parts.Add("text-align:right");   break;
            case CellTextAlign.Justify: parts.Add("text-align:justify"); break;
        }
        if (!string.IsNullOrEmpty(cell.BackgroundColor))
            parts.Add($"background-color:{cell.BackgroundColor}");

        // 패딩 — 4면이 모두 같으면 단축형, 아니면 개별 출력.
        var pt = cell.PaddingTopMm;    var pb = cell.PaddingBottomMm;
        var pl = cell.PaddingLeftMm;   var pr = cell.PaddingRightMm;
        bool sameAll = Math.Abs(pt - pb) < 0.01 && Math.Abs(pt - pl) < 0.01 && Math.Abs(pt - pr) < 0.01;
        if (pt > 0 || pb > 0 || pl > 0 || pr > 0)
        {
            if (sameAll && pt > 0)
                parts.Add($"padding:{FmtMm(pt)}");
            else
            {
                if (pt > 0) parts.Add($"padding-top:{FmtMm(pt)}");
                if (pb > 0) parts.Add($"padding-bottom:{FmtMm(pb)}");
                if (pl > 0) parts.Add($"padding-left:{FmtMm(pl)}");
                if (pr > 0) parts.Add($"padding-right:{FmtMm(pr)}");
            }
        }

        // per-side border (BorderTop/Bottom/Left/Right). null 이면 공통값 fallback.
        var topCss = BorderCss(cell.BorderTop,    cell.BorderThicknessPt, cell.BorderColor);
        var btmCss = BorderCss(cell.BorderBottom, cell.BorderThicknessPt, cell.BorderColor);
        var lftCss = BorderCss(cell.BorderLeft,   cell.BorderThicknessPt, cell.BorderColor);
        var rgtCss = BorderCss(cell.BorderRight,  cell.BorderThicknessPt, cell.BorderColor);

        if (!string.IsNullOrEmpty(topCss) && topCss == btmCss && topCss == lftCss && topCss == rgtCss)
        {
            parts.Add($"border:{topCss}");
        }
        else
        {
            if (!string.IsNullOrEmpty(topCss)) parts.Add($"border-top:{topCss}");
            if (!string.IsNullOrEmpty(btmCss)) parts.Add($"border-bottom:{btmCss}");
            if (!string.IsNullOrEmpty(lftCss)) parts.Add($"border-left:{lftCss}");
            if (!string.IsNullOrEmpty(rgtCss)) parts.Add($"border-right:{rgtCss}");
        }

        return string.Join(';', parts);
    }

    /// <summary>per-side border CSS — `{pt}pt solid {color}` 형식. side 가 null 이면 공통값 사용.</summary>
    private static string BorderCss(CellBorderSide? side, double defPt, string? defColor)
    {
        var pt    = side.HasValue && side.Value.ThicknessPt > 0 ? side.Value.ThicknessPt : defPt;
        var color = side.HasValue && !string.IsNullOrEmpty(side.Value.Color) ? side.Value.Color! : defColor;
        if (pt <= 0) return "";
        var c = string.IsNullOrEmpty(color) ? "#C8C8C8" : color;
        return $"{pt.ToString("0.##", CultureInfo.InvariantCulture)}pt solid {c}";
    }

    private static void WriteImage(StringBuilder sb, ImageBlock img, string indent)
    {
        var styleAttr = BuildImageStyle(img);
        var styleStr  = string.IsNullOrEmpty(styleAttr) ? "" : $" style=\"{styleAttr}\"";

        // SVG ImageBlock → inline <svg> (XHTML 구조 유지, 재임포트 가능).
        if (img.MediaType == "image/svg+xml" && img.Data.Length > 0)
        {
            var svgContent = Encoding.UTF8.GetString(img.Data);
            if (img.ShowTitle && !string.IsNullOrEmpty(img.Title))
            {
                sb.Append(indent).Append("<figure>\n");
                sb.Append(indent).Append("  ").Append(svgContent).Append('\n');
                sb.Append(indent).Append("  <figcaption>").Append(EscapeText(img.Title!)).Append("</figcaption>\n");
                sb.Append(indent).Append("</figure>\n");
            }
            else
            {
                sb.Append(indent).Append(svgContent).Append('\n');
            }
            return;
        }

        var src = img.ResourcePath ?? BuildDataUri(img);
        var alt = EscapeAttr(img.Description ?? "");
        var size = new StringBuilder();
        if (img.WidthMm  > 0) size.Append(" width=\"") .Append(MmToPx(img.WidthMm) .ToString("0", CultureInfo.InvariantCulture)).Append('"');
        if (img.HeightMm > 0) size.Append(" height=\"").Append(MmToPx(img.HeightMm).ToString("0", CultureInfo.InvariantCulture)).Append('"');

        if (img.ShowTitle && !string.IsNullOrEmpty(img.Title))
        {
            sb.Append(indent).Append("<figure>\n");
            sb.Append(indent).Append("  <img src=\"").Append(EscapeAttr(src))
              .Append("\" alt=\"").Append(alt).Append('"').Append(size).Append(styleStr).Append("/>\n");
            sb.Append(indent).Append("  <figcaption>").Append(EscapeText(img.Title!)).Append("</figcaption>\n");
            sb.Append(indent).Append("</figure>\n");
        }
        else
        {
            sb.Append(indent).Append("<img src=\"").Append(EscapeAttr(src))
              .Append("\" alt=\"").Append(alt).Append('"').Append(size).Append(styleStr).Append("/>\n");
        }
    }

    /// <summary>ContainerBlock 을 XHTML 의 &lt;div&gt; 로 직렬화 — HTML 코덱과 동일 패턴.</summary>
    private static void WriteContainer(StringBuilder sb, ContainerBlock box, string indent, NoteNums? notes)
    {
        var parts = new List<string>(8);
        if (box.BorderTopPt    > 0) parts.Add($"border-top:{box.BorderTopPt.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture)}pt solid {box.BorderTopColor ?? "#CCCCCC"}");
        if (box.BorderRightPt  > 0) parts.Add($"border-right:{box.BorderRightPt.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture)}pt solid {box.BorderRightColor ?? "#CCCCCC"}");
        if (box.BorderBottomPt > 0) parts.Add($"border-bottom:{box.BorderBottomPt.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture)}pt solid {box.BorderBottomColor ?? "#CCCCCC"}");
        if (box.BorderLeftPt   > 0) parts.Add($"border-left:{box.BorderLeftPt.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture)}pt solid {box.BorderLeftColor ?? "#CCCCCC"}");
        if (box.BackgroundColor is { Length: > 0 } bg) parts.Add($"background-color:{bg}");
        if (box.PaddingTopMm    > 0) parts.Add($"padding-top:{box.PaddingTopMm.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture)}mm");
        if (box.PaddingRightMm  > 0) parts.Add($"padding-right:{box.PaddingRightMm.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture)}mm");
        if (box.PaddingBottomMm > 0) parts.Add($"padding-bottom:{box.PaddingBottomMm.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture)}mm");
        if (box.PaddingLeftMm   > 0) parts.Add($"padding-left:{box.PaddingLeftMm.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture)}mm");
        if (box.MarginTopMm     > 0) parts.Add($"margin-top:{box.MarginTopMm.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture)}mm");
        if (box.MarginBottomMm  > 0) parts.Add($"margin-bottom:{box.MarginBottomMm.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture)}mm");
        var styleAttr = parts.Count > 0 ? $" style=\"{string.Join(';', parts)}\"" : "";
        var classAttr = string.IsNullOrEmpty(box.ClassNames) ? "" : $" class=\"{System.Net.WebUtility.HtmlEncode(box.ClassNames!)}\"";
        sb.Append(indent).Append("<div").Append(classAttr).Append(styleAttr).Append(">\n");
        WriteBlocks(sb, box.Children, indent + "  ", notes);
        sb.Append(indent).Append("</div>\n");
    }

    private static void WriteToc(StringBuilder sb, TocBlock toc, string indent)
    {
        sb.Append(indent).Append("<nav class=\"pd-toc\">\n");
        sb.Append(indent).Append("  <p class=\"pd-toc-title\"><strong>목차</strong></p>\n");
        foreach (var entry in toc.Entries)
        {
            var lvl = Math.Clamp(entry.Level, 1, 6);
            var pad = ((lvl - 1) * 4).ToString(CultureInfo.InvariantCulture);
            sb.Append(indent).Append("  <p class=\"pd-toc-l").Append(lvl)
              .Append("\" style=\"padding-left:").Append(pad).Append("mm\">");
            sb.Append(EscapeText(entry.Text));
            if (entry.PageNumber.HasValue)
                sb.Append("\t<span class=\"pd-toc-page\">").Append(entry.PageNumber.Value).Append("</span>");
            sb.Append("</p>\n");
        }
        sb.Append(indent).Append("</nav>\n");
    }

    private static void WriteShape(StringBuilder sb, ShapeObject shape, string indent)
    {
        var wPx = MmToPx(shape.WidthMm  > 0 ? shape.WidthMm  : 40);
        var hPx = MmToPx(shape.HeightMm > 0 ? shape.HeightMm : 30);

        var alignStyle = shape.HAlign switch
        {
            ImageHAlign.Center => " style=\"display:block;margin-left:auto;margin-right:auto\"",
            ImageHAlign.Right  => " style=\"display:block;margin-left:auto\"",
            _                  => "",
        };

        var svgBody = BuildShapeSvgBody(shape, wPx, hPx);
        sb.Append(indent).Append("<figure class=\"pd-shape\"").Append(alignStyle).Append(">\n");
        sb.Append(indent).Append("  <svg xmlns=\"http://www.w3.org/2000/svg\" width=\"")
          .Append(wPx.ToString("0.#", CultureInfo.InvariantCulture))
          .Append("\" height=\"").Append(hPx.ToString("0.#", CultureInfo.InvariantCulture)).Append("\">");
        sb.Append(svgBody).Append("</svg>\n");
        if (!string.IsNullOrEmpty(shape.LabelText))
            sb.Append(indent).Append("  <figcaption>").Append(EscapeText(shape.LabelText)).Append("</figcaption>\n");
        sb.Append(indent).Append("</figure>\n");
    }

    // XHTML5 polyglot — void SVG elements use self-closing />.
    private static string BuildShapeSvgBody(ShapeObject shape, double wPx, double hPx)
    {
        var stroke = EscapeAttr(shape.StrokeColor);
        var fill   = shape.FillColor is { Length: > 0 } fc ? EscapeAttr(fc) : "none";
        var sw     = shape.StrokeThicknessPt.ToString("0.##", CultureInfo.InvariantCulture);

        return shape.Kind switch
        {
            ShapeKind.Rectangle =>
                $"<rect x=\"0.5\" y=\"0.5\" width=\"{wPx - 1:0.#}\" height=\"{hPx - 1:0.#}\" stroke=\"{stroke}\" stroke-width=\"{sw}\" fill=\"{fill}\"/>",

            ShapeKind.RoundedRect =>
                $"<rect x=\"0.5\" y=\"0.5\" width=\"{wPx - 1:0.#}\" height=\"{hPx - 1:0.#}\" rx=\"{MmToPx(shape.CornerRadiusMm):0.#}\" stroke=\"{stroke}\" stroke-width=\"{sw}\" fill=\"{fill}\"/>",

            ShapeKind.Ellipse =>
                $"<ellipse cx=\"{wPx / 2:0.#}\" cy=\"{hPx / 2:0.#}\" rx=\"{(wPx - 1) / 2:0.#}\" ry=\"{(hPx - 1) / 2:0.#}\" stroke=\"{stroke}\" stroke-width=\"{sw}\" fill=\"{fill}\"/>",

            ShapeKind.Line when shape.Points.Count >= 2 =>
                $"<line x1=\"{MmToPx(shape.Points[0].X):0.#}\" y1=\"{MmToPx(shape.Points[0].Y):0.#}\" x2=\"{MmToPx(shape.Points[^1].X):0.#}\" y2=\"{MmToPx(shape.Points[^1].Y):0.#}\" stroke=\"{stroke}\" stroke-width=\"{sw}\" fill=\"none\"/>",

            ShapeKind.Polyline when shape.Points.Count >= 2 =>
                $"<polyline points=\"{PointsToSvg(shape.Points)}\" stroke=\"{stroke}\" stroke-width=\"{sw}\" fill=\"none\"/>",

            ShapeKind.Polygon or ShapeKind.Triangle when shape.Points.Count >= 3 =>
                $"<polygon points=\"{PointsToSvg(shape.Points)}\" stroke=\"{stroke}\" stroke-width=\"{sw}\" fill=\"{fill}\"/>",

            ShapeKind.Spline or ShapeKind.ClosedSpline when shape.Points.Count >= 2 =>
                BuildSplinePath(shape, stroke, fill, sw),

            ShapeKind.RegularPolygon =>
                BuildRegularPolygonSvg(shape, wPx, hPx, stroke, fill, sw),

            ShapeKind.Star =>
                BuildStarSvg(shape, wPx, hPx, stroke, fill, sw),

            _ =>
                $"<rect x=\"0.5\" y=\"0.5\" width=\"{wPx - 1:0.#}\" height=\"{hPx - 1:0.#}\" stroke=\"{stroke}\" stroke-width=\"{sw}\" fill=\"{fill}\"/>",
        };
    }

    private static string PointsToSvg(IList<ShapePoint> pts)
        => string.Join(" ", pts.Select(p => $"{MmToPx(p.X):0.#},{MmToPx(p.Y):0.#}"));

    private static string BuildSplinePath(ShapeObject shape, string stroke, string fill, string sw)
    {
        var pts = shape.Points;
        bool closed = shape.Kind == ShapeKind.ClosedSpline;
        var sb = new StringBuilder();
        sb.Append($"<path d=\"M{MmToPx(pts[0].X):0.#},{MmToPx(pts[0].Y):0.#}");
        int count = closed ? pts.Count : pts.Count - 1;
        for (int i = 0; i < count; i++)
        {
            int j    = (i + 1) % pts.Count;
            int prev = (i - 1 + pts.Count) % pts.Count;
            int next = (j + 1) % pts.Count;
            var p0 = pts[i]; var p1 = pts[j];
            if (p0.OutCtrlX.HasValue && p1.InCtrlX.HasValue)
            {
                sb.Append($" C{MmToPx(p0.OutCtrlX.Value):0.#},{MmToPx(p0.OutCtrlY!.Value):0.#} " +
                          $"{MmToPx(p1.InCtrlX.Value):0.#},{MmToPx(p1.InCtrlY!.Value):0.#} " +
                          $"{MmToPx(p1.X):0.#},{MmToPx(p1.Y):0.#}");
            }
            else
            {
                var cp0x = pts[i].X + (pts[j].X - pts[prev].X) / 6.0;
                var cp0y = pts[i].Y + (pts[j].Y - pts[prev].Y) / 6.0;
                var cp1x = pts[j].X - (pts[next].X - pts[i].X) / 6.0;
                var cp1y = pts[j].Y - (pts[next].Y - pts[i].Y) / 6.0;
                sb.Append($" C{MmToPx(cp0x):0.#},{MmToPx(cp0y):0.#} " +
                          $"{MmToPx(cp1x):0.#},{MmToPx(cp1y):0.#} " +
                          $"{MmToPx(p1.X):0.#},{MmToPx(p1.Y):0.#}");
            }
        }
        if (closed) sb.Append(" Z");
        sb.Append($"\" stroke=\"{stroke}\" stroke-width=\"{sw}\" fill=\"{fill}\"/>");
        return sb.ToString();
    }

    private static string BuildRegularPolygonSvg(ShapeObject shape, double wPx, double hPx, string stroke, string fill, string sw)
    {
        int sides = Math.Max(3, shape.SideCount);
        double cx = wPx / 2, cy = hPx / 2, rx = (wPx - 1) / 2, ry = (hPx - 1) / 2;
        var pts = string.Join(" ", Enumerable.Range(0, sides).Select(i =>
        {
            double a = 2 * Math.PI * i / sides - Math.PI / 2;
            return $"{(cx + rx * Math.Cos(a)):0.#},{(cy + ry * Math.Sin(a)):0.#}";
        }));
        return $"<polygon points=\"{pts}\" stroke=\"{stroke}\" stroke-width=\"{sw}\" fill=\"{fill}\"/>";
    }

    private static string BuildStarSvg(ShapeObject shape, double wPx, double hPx, string stroke, string fill, string sw)
    {
        int spikes = Math.Max(3, shape.SideCount);
        double cx = wPx / 2, cy = hPx / 2;
        double orx = (wPx - 1) / 2, ory = (hPx - 1) / 2;
        double irx = orx * shape.InnerRadiusRatio, iry = ory * shape.InnerRadiusRatio;
        var pts = string.Join(" ", Enumerable.Range(0, spikes * 2).Select(i =>
        {
            double a  = Math.PI * i / spikes - Math.PI / 2;
            double rx = i % 2 == 0 ? orx : irx;
            double ry = i % 2 == 0 ? ory : iry;
            return $"{(cx + rx * Math.Cos(a)):0.#},{(cy + ry * Math.Sin(a)):0.#}";
        }));
        return $"<polygon points=\"{pts}\" stroke=\"{stroke}\" stroke-width=\"{sw}\" fill=\"{fill}\"/>";
    }

    private static void WriteTextBox(StringBuilder sb, TextBoxObject tbox, string indent, NoteNums? notes = null)
    {
        var parts = new List<string>(8);
        if (tbox.WidthMm  > 0) parts.Add($"width:{FmtMm(tbox.WidthMm)}");
        if (tbox.HeightMm > 0) parts.Add($"min-height:{FmtMm(tbox.HeightMm)}");
        if (!string.IsNullOrEmpty(tbox.BackgroundColor)) parts.Add($"background-color:{tbox.BackgroundColor}");
        if (tbox.BorderThicknessPt > 0)
            parts.Add($"border:{tbox.BorderThicknessPt.ToString("0.##", CultureInfo.InvariantCulture)}pt solid {(tbox.BorderColor ?? "#888888")}");
        double pt = tbox.PaddingTopMm, pb = tbox.PaddingBottomMm, pl = tbox.PaddingLeftMm, pr = tbox.PaddingRightMm;
        if (pt > 0 || pb > 0 || pl > 0 || pr > 0)
            parts.Add($"padding:{FmtMm(pt)} {FmtMm(pr)} {FmtMm(pb)} {FmtMm(pl)}");
        if (tbox.RotationAngleDeg != 0)
            parts.Add($"transform:rotate({tbox.RotationAngleDeg.ToString("0.##", CultureInfo.InvariantCulture)}deg)");

        var styleAttr = parts.Count > 0 ? $" style=\"{string.Join(";", parts)}\"" : "";
        sb.Append(indent).Append("<div class=\"pd-textbox\"").Append(styleAttr).Append(">\n");
        WriteBlocks(sb, tbox.Content, indent + "  ", notes);
        sb.Append(indent).Append("</div>\n");
    }

    private static void WriteOpaque(StringBuilder sb, OpaqueBlock opq, string indent)
    {
        sb.Append(indent)
          .Append("<div class=\"pd-opaque\" data-pd-format=\"").Append(EscapeAttr(opq.Format)).Append("\">")
          .Append(EscapeText(opq.DisplayLabel))
          .Append("</div>\n");
    }

    private static string BuildImageStyle(ImageBlock img)
    {
        var parts = new List<string>(4);

        // WrapMode → CSS float (텍스트 감싸기).
        // WrapLeft = 이미지 오른쪽에 텍스트 → float:right
        // WrapRight = 이미지 왼쪽에 텍스트 → float:left
        switch (img.WrapMode)
        {
            case ImageWrapMode.WrapLeft:  parts.Add("float:right"); break;
            case ImageWrapMode.WrapRight: parts.Add("float:left");  break;
        }

        // HAlign (WrapMode=Inline 일 때) → display:block + margin auto.
        if (img.WrapMode == ImageWrapMode.Inline)
        {
            switch (img.HAlign)
            {
                case ImageHAlign.Center: parts.Add("display:block;margin-left:auto;margin-right:auto"); break;
                case ImageHAlign.Right:  parts.Add("display:block;margin-left:auto");                   break;
            }
        }

        if (img.MarginTopMm    > 0) parts.Add($"margin-top:{FmtMm(img.MarginTopMm)}");
        if (img.MarginBottomMm > 0) parts.Add($"margin-bottom:{FmtMm(img.MarginBottomMm)}");

        return string.Join(';', parts);
    }

    private static string BuildDataUri(ImageBlock img)
    {
        if (img.Data.Length == 0) return "";
        var b64 = Convert.ToBase64String(img.Data);
        return $"data:{img.MediaType};base64,{b64}";
    }

    private static double MmToPx(double mm) => mm * 96.0 / 25.4;

    // ── 인라인 (Run) 렌더링 ──────────────────────────────────────────

    private static string RenderRuns(IList<Run> runs, NoteNums? notes = null)
    {
        var sb = new StringBuilder();
        foreach (var r in runs) sb.Append(RenderRun(r, notes));
        return sb.ToString();
    }

    private static string RenderRun(Run run, NoteNums? notes = null)
    {
        // 각주/미주 참조 런 — Pandoc 스타일 superscript 링크로 직렬화.
        if (run.FootnoteId is { Length: > 0 } fnId
            && notes is not null && notes.Footnotes.TryGetValue(fnId, out var fnNum))
        {
            return $"<sup id=\"fnref-{fnNum}\"><a href=\"#fn-{fnNum}\">{fnNum}</a></sup>";
        }
        if (run.EndnoteId is { Length: > 0 } enId
            && notes is not null && notes.Endnotes.TryGetValue(enId, out var enNum))
        {
            return $"<sup id=\"enref-{enNum}\"><a href=\"#en-{enNum}\">{enNum}</a></sup>";
        }
        if (run.FootnoteId is { Length: > 0 } || run.EndnoteId is { Length: > 0 })
            return string.Empty;

        // LaTeX 수식
        if (run.LatexSource is { Length: > 0 } latex)
        {
            var escaped = EscapeText(latex);
            return run.IsDisplayEquation
                ? $"<span class=\"pd-math pd-math-display\">\\[{escaped}\\]</span>"
                : $"<span class=\"pd-math\">\\({escaped}\\)</span>";
        }

        // 이모지
        if (run.EmojiKey is { Length: > 0 } emojiKey)
        {
            var parts = emojiKey.Split('_', 2);
            var name  = parts.Length == 2 ? parts[1] : emojiKey;
            return $"<span class=\"pd-emoji\" data-pd-emoji=\"{EscapeAttr(emojiKey)}\" title=\"{EscapeAttr(name)}\">{EscapeText(name)}</span>";
        }

        // 인라인 필드
        if (run.Field.HasValue)
        {
            var (cls, placeholder) = run.Field.Value switch
            {
                FieldType.Page     => ("page",     "1"),
                FieldType.NumPages => ("numpages", "1"),
                FieldType.Date     => ("date",     DateTime.Today.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)),
                FieldType.Time     => ("time",     DateTime.Now.ToString("HH:mm", CultureInfo.InvariantCulture)),
                FieldType.Author   => ("author",   ""),
                FieldType.Title    => ("title",    ""),
                _                  => ("field",    ""),
            };
            return $"<span class=\"pd-field pd-field-{cls}\">{EscapeText(placeholder)}</span>";
        }

        var s    = run.Style;
        var text = EscapeText(run.Text).Replace("\n", "<br/>");

        bool isMono = !string.IsNullOrEmpty(s.FontFamily) &&
                      s.FontFamily.Contains("monospace", StringComparison.OrdinalIgnoreCase);

        var styleAttr = BuildSpanStyle(s, includeMono: false);
        if (!string.IsNullOrEmpty(styleAttr))
            text = $"<span style=\"{styleAttr}\">{text}</span>";

        if (isMono)            text = $"<code>{text}</code>";
        if (s.Subscript)       text = $"<sub>{text}</sub>";
        if (s.Superscript)     text = $"<sup>{text}</sup>";
        if (s.Underline)       text = $"<u>{text}</u>";
        if (s.Strikethrough)   text = $"<s>{text}</s>";
        if (s.Italic)          text = $"<em>{text}</em>";
        if (s.Bold)            text = $"<strong>{text}</strong>";

        if (run.Url is { Length: > 0 } href)
            text = $"<a href=\"{EscapeAttr(href)}\">{text}</a>";

        return text;
    }

    private static string BuildSpanStyle(RunStyle s, bool includeMono)
    {
        var parts = new List<string>(5);
        if (!string.IsNullOrEmpty(s.FontFamily) &&
            (includeMono || !s.FontFamily.Contains("monospace", StringComparison.OrdinalIgnoreCase)))
        {
            parts.Add($"font-family:{s.FontFamily}");
        }
        if (Math.Abs(s.FontSizePt - 11) > 0.01 && s.FontSizePt > 0)
            parts.Add($"font-size:{s.FontSizePt.ToString("0.##", CultureInfo.InvariantCulture)}pt");
        if (s.Foreground is { } fg) parts.Add($"color:{ColorHex(fg)}");
        if (s.Background is { } bg) parts.Add($"background-color:{ColorHex(bg)}");
        // Overline 은 semantic HTML 태그 없음 → CSS로만 표현.
        if (s.Overline) parts.Add("text-decoration:overline");
        if (Math.Abs(s.WidthPercent - 100) > 0.5)
            parts.Add($"transform:scaleX({(s.WidthPercent / 100.0).ToString("0.###", CultureInfo.InvariantCulture)});display:inline-block");
        if (Math.Abs(s.LetterSpacingPx) > 0.01)
            parts.Add($"letter-spacing:{s.LetterSpacingPx.ToString("0.##", CultureInfo.InvariantCulture)}px");
        return string.Join(';', parts);
    }

    private static string ColorHex(Color c) => $"#{c.R:X2}{c.G:X2}{c.B:X2}";

    // ── 이스케이프 (XML 규정) ────────────────────────────────────────

    private static string EscapeText(string text)
    {
        if (string.IsNullOrEmpty(text)) return text;
        var sb = new StringBuilder(text.Length);
        foreach (var ch in text)
        {
            switch (ch)
            {
                case '&': sb.Append("&amp;");  break;
                case '<': sb.Append("&lt;");   break;
                case '>': sb.Append("&gt;");   break;
                default:
                    // XML 1.0 invalid characters 제거 — \x00-\x08, \x0B, \x0C, \x0E-\x1F.
                    if (ch < 0x20 && ch != '\t' && ch != '\n' && ch != '\r') break;
                    sb.Append(ch);
                    break;
            }
        }
        return sb.ToString();
    }

    private static string EscapeAttr(string text)
    {
        if (string.IsNullOrEmpty(text)) return text;
        var sb = new StringBuilder(text.Length);
        foreach (var ch in text)
        {
            switch (ch)
            {
                case '&':  sb.Append("&amp;");  break;
                case '<':  sb.Append("&lt;");   break;
                case '>':  sb.Append("&gt;");   break;
                case '"':  sb.Append("&quot;"); break;
                case '\'': sb.Append("&apos;"); break;
                default:
                    if (ch < 0x20 && ch != '\t' && ch != '\n' && ch != '\r') break;
                    sb.Append(ch);
                    break;
            }
        }
        return sb.ToString();
    }
}

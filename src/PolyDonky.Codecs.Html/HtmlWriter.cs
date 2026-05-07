using System.Globalization;
using System.Text;
using System.Linq;
using PolyDonky.Core;

namespace PolyDonky.Codecs.Html;

/// <summary>
/// PolyDonkyument 를 HTML5 로 직렬화한다 — UTF-8, &lt;!DOCTYPE html&gt;, semantic 마크업.
///
/// 매핑:
///   - OutlineLevel.H1~H6   → &lt;h1&gt;~&lt;h6&gt;
///   - 일반 단락             → &lt;p&gt;
///   - QuoteLevel ≥ 1       → 중첩 &lt;blockquote&gt;
///   - IsThematicBreak      → &lt;hr&gt;
///   - CodeLanguage non-null → &lt;pre&gt;&lt;code class="language-xxx"&gt;...&lt;/code&gt;&lt;/pre&gt;
///   - ListMarker (bullet/ordered, nested by Level) → &lt;ul&gt;/&lt;ol&gt; + &lt;li&gt;
///   - ListMarker.Checked    → &lt;input type="checkbox" disabled checked?&gt; 접두
///   - Run.Bold/Italic/Strike/Sub/Super/Underline/Overline → &lt;strong&gt;&lt;em&gt;&lt;s&gt;&lt;sub&gt;&lt;sup&gt;&lt;u&gt; + span CSS
///   - 모노스페이스 FontFamily → &lt;code&gt;
///   - Run.Url               → &lt;a href="..."&gt;
///   - Run.Foreground/Background/FontSizePt/FontFamily → &lt;span style="..."&gt;
///   - Table                 → &lt;table&gt;&lt;colgroup&gt;&lt;thead&gt;/&lt;tbody&gt; + 셀 정렬·배경·패딩·테두리 style
///   - ImageBlock            → &lt;img src alt width height style&gt; (또는 &lt;figure&gt; + &lt;figcaption&gt;)
///   - ParagraphStyle.LineHeightFactor/SpaceBeforePt/SpaceAfterPt/IndentFirstLineMm/IndentLeftMm/IndentRightMm → CSS
/// </summary>
public sealed class HtmlWriter : IDocumentWriter
{
    public string FormatId => "html";

    public Encoding Encoding { get; init; } = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

    /// <summary>완전한 HTML5 문서로 출력할지 (기본 true) — false 면 fragment.</summary>
    public bool FullDocument { get; init; } = true;

    /// <summary>문서 제목 — &lt;title&gt; 에 사용. null 이면 첫 H1 텍스트 사용.</summary>
    public string? DocumentTitle { get; init; }

    public void Write(PolyDonkyument document, Stream output)
    {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(output);

        using var writer = new StreamWriter(output, Encoding, leaveOpen: true) { NewLine = "\n" };
        writer.Write(ToHtml(document, FullDocument, DocumentTitle));
    }

    public static string ToHtml(PolyDonkyument document, bool fullDocument = true, string? title = null)
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

            sb.Append("<!DOCTYPE html>\n");
            sb.Append("<html lang=\"ko\">\n");
            sb.Append("<head>\n");
            sb.Append("  <meta charset=\"utf-8\">\n");
            sb.Append("  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n");
            sb.Append("  <meta name=\"generator\" content=\"PolyDonky\">\n");
            sb.Append("  <title>").Append(EscapeHtml(docTitle)).Append("</title>\n");
            // 편집용지 메타 태그 — HtmlReader 가 Section.Page 로 복원한다.
            sb.Append("  <meta name=\"pd-page-size\" content=\"").Append(page.SizeKind).Append("\">\n");
            sb.Append("  <meta name=\"pd-page-orientation\" content=\"")
              .Append(page.Orientation == PageOrientation.Landscape ? "landscape" : "portrait")
              .Append("\">\n");
            if (page.SizeKind == PaperSizeKind.Custom)
            {
                sb.Append("  <meta name=\"pd-page-width\" content=\"")
                  .Append(page.WidthMm.ToString("0.##", CultureInfo.InvariantCulture)).Append("mm\">\n");
                sb.Append("  <meta name=\"pd-page-height\" content=\"")
                  .Append(page.HeightMm.ToString("0.##", CultureInfo.InvariantCulture)).Append("mm\">\n");
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
            sb.Append(indent).Append("  <hr>\n");
            sb.Append(indent).Append("  <ol>\n");
            foreach (var entry in doc.Footnotes)
            {
                if (!notes.Footnotes.TryGetValue(entry.Id, out var num)) continue;
                sb.Append(indent).Append("    <li id=\"fn-").Append(num).Append("\">");
                sb.Append(RenderBlocks(entry.Blocks, notes));
                sb.Append(" <a href=\"#fnref-").Append(num).Append("\">↩</a>");
                sb.Append("</li>\n");
            }
            sb.Append(indent).Append("  </ol>\n");
            sb.Append(indent).Append("</section>\n");
        }

        if (doc.Endnotes.Count > 0)
        {
            sb.Append(indent).Append("<section class=\"endnotes\">\n");
            sb.Append(indent).Append("  <hr>\n");
            sb.Append(indent).Append("  <ol>\n");
            foreach (var entry in doc.Endnotes)
            {
                if (!notes.Endnotes.TryGetValue(entry.Id, out var num)) continue;
                sb.Append(indent).Append("    <li id=\"en-").Append(num).Append("\">");
                sb.Append(RenderBlocks(entry.Blocks, notes));
                sb.Append(" <a href=\"#enref-").Append(num).Append("\">↩</a>");
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

    // ── 블록 렌더링 ─────────────────────────────────────────────────────

    private static void WriteBlocks(StringBuilder sb, IList<Block> blocks, string indent, NoteNums? notes = null)
    {
        // 인접 리스트·인용은 묶어서 처리한다.
        int i = 0;
        while (i < blocks.Count)
        {
            var b = blocks[i];

            // 연속된 같은 인용 깊이의 블록을 <blockquote> 로 감싼다.
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

            // 연속된 리스트 단락 묶음.
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
                case Paragraph para:     WriteParagraph(sb, para, indent, notes); break;
                case Table table:        WriteTable(sb, table, indent, notes);    break;
                case ImageBlock img:     WriteImage(sb, img, indent);             break;
                case TocBlock toc:       WriteToc(sb, toc, indent);               break;
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
        var copy = new Paragraph
        {
            StyleId = p.StyleId,
            Style   = new ParagraphStyle
            {
                Alignment         = p.Style.Alignment,
                LineHeightFactor  = p.Style.LineHeightFactor,
                SpaceBeforePt     = p.Style.SpaceBeforePt,
                SpaceAfterPt      = p.Style.SpaceAfterPt,
                IndentFirstLineMm = p.Style.IndentFirstLineMm,
                IndentLeftMm      = p.Style.IndentLeftMm,
                IndentRightMm     = p.Style.IndentRightMm,
                Outline           = p.Style.Outline,
                ListMarker        = p.Style.ListMarker,
                QuoteLevel        = Math.Max(0, p.Style.QuoteLevel - 1),
                CodeLanguage      = p.Style.CodeLanguage,
                IsThematicBreak   = p.Style.IsThematicBreak,
            },
            Runs    = p.Runs,
        };
        return copy;
    }

    private static void WriteParagraph(StringBuilder sb, Paragraph p, string indent, NoteNums? notes = null)
    {
        if (p.Style.IsThematicBreak)
        {
            sb.Append(indent).Append("<hr>\n");
            return;
        }

        if (p.Style.CodeLanguage is not null)
        {
            var code = EscapeHtml(p.GetPlainText());
            var langAttr = p.Style.CodeLanguage.Length > 0
                ? $" class=\"language-{EscapeAttr(p.Style.CodeLanguage)}\""
                : "";
            sb.Append(indent).Append("<pre><code").Append(langAttr).Append('>')
              .Append(code).Append("</code></pre>\n");
            return;
        }

        var classAttr = BuildClassAttr(p.StyleId);

        if (p.Style.Outline > OutlineLevel.Body)
        {
            int lvl = (int)p.Style.Outline;
            var styleAttr = ParagraphStyleAttr(p.Style);
            sb.Append(indent).Append('<').Append('h').Append(lvl).Append(classAttr).Append(styleAttr).Append('>');
            sb.Append(RenderRuns(p.Runs, notes));
            sb.Append("</h").Append(lvl).Append(">\n");
            return;
        }

        var pStyleAttr = ParagraphStyleAttr(p.Style);
        sb.Append(indent).Append("<p").Append(classAttr).Append(pStyleAttr).Append('>');
        sb.Append(RenderRuns(p.Runs, notes));
        sb.Append("</p>\n");
    }

    private static List<string> BuildParagraphCssParts(ParagraphStyle s)
    {
        var parts = new List<string>(6);

        var ta = s.Alignment switch
        {
            Alignment.Center  => "center",
            Alignment.Right   => "right",
            Alignment.Justify => "justify",
            _                 => null,
        };
        if (ta is not null) parts.Add($"text-align:{ta}");

        // 기본값 1.2 와 다를 때만 출력 (브라우저 기본값 normal ≈ 1.2 에 해당).
        if (s.LineHeightFactor > 0 && Math.Abs(s.LineHeightFactor - 1.2) > 0.01)
            parts.Add($"line-height:{s.LineHeightFactor.ToString("0.##", CultureInfo.InvariantCulture)}");

        if (s.SpaceBeforePt > 0)
            parts.Add($"margin-top:{s.SpaceBeforePt.ToString("0.##", CultureInfo.InvariantCulture)}pt");
        if (s.SpaceAfterPt > 0)
            parts.Add($"margin-bottom:{s.SpaceAfterPt.ToString("0.##", CultureInfo.InvariantCulture)}pt");
        if (Math.Abs(s.IndentFirstLineMm) > 0.01)
            parts.Add($"text-indent:{s.IndentFirstLineMm.ToString("0.##", CultureInfo.InvariantCulture)}mm");
        if (s.IndentLeftMm > 0)
            parts.Add($"padding-left:{s.IndentLeftMm.ToString("0.##", CultureInfo.InvariantCulture)}mm");
        if (s.IndentRightMm > 0)
            parts.Add($"padding-right:{s.IndentRightMm.ToString("0.##", CultureInfo.InvariantCulture)}mm");

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
        bool hasPage   = page is not null;
        if (!hasStyles && !hasPage) return;

        sb.Append("  <style>\n");

        // @page 규칙 — 편집용지 크기·여백을 CSS 인쇄 표준으로 직렬화한다.
        if (page is not null)
        {
            var w = page.EffectiveWidthMm.ToString("0.##",  CultureInfo.InvariantCulture);
            var h = page.EffectiveHeightMm.ToString("0.##", CultureInfo.InvariantCulture);
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

    private static void WriteListGroup(StringBuilder sb, IList<Block> blocks, int from, int to, string indent, NoteNums? notes = null)
    {
        var p0 = (Paragraph)blocks[from];
        var lm = p0.Style.ListMarker!;
        string tag = lm.Kind == ListKind.Bullet ? "ul" : "ol";
        string startAttr = "";
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
                var ck = marker.Checked.Value ? " checked" : "";
                sb.Append("<input type=\"checkbox\" disabled").Append(ck).Append("> ");
            }
            sb.Append(RenderRuns(p.Runs, notes));
            sb.Append("</li>\n");
        }

        sb.Append(indent).Append("</").Append(tag).Append(">\n");
    }

    private static void WriteTable(StringBuilder sb, Table t, string indent, NoteNums? notes = null)
    {
        if (t.Rows.Count == 0) return;

        // 표 수준 style (배경색·정렬).
        var tblStyle = new List<string>(3);
        if (!string.IsNullOrEmpty(t.BackgroundColor))
            tblStyle.Add($"background-color:{t.BackgroundColor}");
        switch (t.HAlign)
        {
            case TableHAlign.Center: tblStyle.Add("margin-left:auto;margin-right:auto"); break;
            case TableHAlign.Right:  tblStyle.Add("margin-left:auto");                  break;
        }
        if (t.BorderThicknessPt > 0)
            tblStyle.Add($"border-collapse:collapse");

        var tblStyleAttr = tblStyle.Count > 0
            ? $" style=\"{string.Join(';', tblStyle)}\""
            : "";

        sb.Append(indent).Append("<table").Append(tblStyleAttr).Append(">\n");

        // <colgroup> — 열 너비가 하나라도 있을 때만 출력.
        if (t.Columns.Any(c => c.WidthMm > 0))
        {
            sb.Append(indent).Append("  <colgroup>\n");
            foreach (var col in t.Columns)
            {
                if (col.WidthMm > 0)
                    sb.Append(indent).Append("    <col style=\"width:")
                      .Append(col.WidthMm.ToString("0.##", CultureInfo.InvariantCulture))
                      .Append("mm\">\n");
                else
                    sb.Append(indent).Append("    <col>\n");
            }
            sb.Append(indent).Append("  </colgroup>\n");
        }

        // <thead> / <tbody> 모음.
        var headerRows = t.Rows.Where(r => r.IsHeader).ToList();
        var bodyRows   = t.Rows.Where(r => !r.IsHeader).ToList();

        if (headerRows.Count > 0)
        {
            sb.Append(indent).Append("  <thead>\n");
            foreach (var r in headerRows) WriteRow(sb, r, indent + "    ", isHeader: true, notes);
            sb.Append(indent).Append("  </thead>\n");
        }
        if (bodyRows.Count > 0)
        {
            sb.Append(indent).Append("  <tbody>\n");
            foreach (var r in bodyRows) WriteRow(sb, r, indent + "    ", isHeader: false, notes);
            sb.Append(indent).Append("  </tbody>\n");
        }

        sb.Append(indent).Append("</table>\n");
    }

    private static void WriteRow(StringBuilder sb, TableRow row, string indent, bool isHeader, NoteNums? notes = null)
    {
        sb.Append(indent).Append("<tr>\n");
        foreach (var cell in row.Cells)
        {
            var tag = isHeader ? "th" : "td";
            var attrs = new StringBuilder();
            if (cell.ColumnSpan > 1) attrs.Append(" colspan=\"").Append(cell.ColumnSpan).Append('"');
            if (cell.RowSpan    > 1) attrs.Append(" rowspan=\"").Append(cell.RowSpan).Append('"');

            var cellStyle = BuildCellStyle(cell);
            if (cellStyle.Length > 0)
                attrs.Append(" style=\"").Append(cellStyle).Append('"');

            sb.Append(indent).Append("  <").Append(tag).Append(attrs).Append('>');
            // 셀 안 블록을 인라인적으로 렌더 — 단락은 <br> 로 구분.
            bool first = true;
            foreach (var b in cell.Blocks)
            {
                if (b is Paragraph p)
                {
                    if (!first) sb.Append("<br>");
                    sb.Append(RenderRuns(p.Runs, notes));
                    first = false;
                }
            }
            sb.Append("</").Append(tag).Append(">\n");
        }
        sb.Append(indent).Append("</tr>\n");
    }

    private static string BuildCellStyle(TableCell cell)
    {
        var parts = new List<string>(8);

        if (cell.TextAlign != CellTextAlign.Left)
        {
            var ta = cell.TextAlign switch
            {
                CellTextAlign.Center  => "center",
                CellTextAlign.Right   => "right",
                CellTextAlign.Justify => "justify",
                _                     => null,
            };
            if (ta is not null) parts.Add($"text-align:{ta}");
        }

        if (!string.IsNullOrEmpty(cell.BackgroundColor))
            parts.Add($"background-color:{cell.BackgroundColor}");

        // 셀 안여백 (mm → CSS)
        double padT = cell.PaddingTopMm,    padB = cell.PaddingBottomMm;
        double padL = cell.PaddingLeftMm,   padR = cell.PaddingRightMm;
        if (padT > 0 || padB > 0 || padL > 0 || padR > 0)
            parts.Add($"padding:{FmtMm(padT)} {FmtMm(padR)} {FmtMm(padB)} {FmtMm(padL)}");

        // 테두리 — 면별 per-side 또는 공통값이 있을 때 출력.
        bool hasBorder = cell.BorderTop is not null || cell.BorderBottom is not null
                      || cell.BorderLeft is not null || cell.BorderRight is not null
                      || cell.BorderThicknessPt > 0 || !string.IsNullOrEmpty(cell.BorderColor);
        if (hasBorder)
        {
            parts.Add($"border-top:{BorderCss(cell.BorderTop,    cell.BorderThicknessPt, cell.BorderColor)}");
            parts.Add($"border-bottom:{BorderCss(cell.BorderBottom, cell.BorderThicknessPt, cell.BorderColor)}");
            parts.Add($"border-left:{BorderCss(cell.BorderLeft,   cell.BorderThicknessPt, cell.BorderColor)}");
            parts.Add($"border-right:{BorderCss(cell.BorderRight,  cell.BorderThicknessPt, cell.BorderColor)}");
        }

        return string.Join(';', parts);
    }

    private static string BorderCss(CellBorderSide? side, double defPt, string? defColor)
    {
        var pt  = side.HasValue && side.Value.ThicknessPt > 0 ? side.Value.ThicknessPt : defPt;
        var clr = side.HasValue && !string.IsNullOrEmpty(side.Value.Color) ? side.Value.Color! : (defColor ?? "#C8C8C8");
        if (pt <= 0) return "none";
        return $"{pt.ToString("0.##", CultureInfo.InvariantCulture)}pt solid {clr}";
    }

    private static string FmtMm(double mm)
        => mm > 0 ? mm.ToString("0.##", CultureInfo.InvariantCulture) + "mm" : "0";

    private static void WriteImage(StringBuilder sb, ImageBlock img, string indent)
    {
        var imgStyle  = BuildImageStyle(img);
        var styleAttr = imgStyle.Length > 0 ? $" style=\"{imgStyle}\"" : "";

        // SVG ImageBlock → inline <svg> (base64 대신 직접 삽입해 가독성·재임포트 유지).
        if (img.MediaType == "image/svg+xml" && img.Data.Length > 0)
        {
            var svgContent = Encoding.UTF8.GetString(img.Data);
            if (img.ShowTitle && !string.IsNullOrEmpty(img.Title))
            {
                sb.Append(indent).Append("<figure").Append(styleAttr).Append(">\n");
                sb.Append(indent).Append("  ").Append(svgContent).Append('\n');
                sb.Append(indent).Append("  <figcaption>").Append(EscapeHtml(img.Title!)).Append("</figcaption>\n");
                sb.Append(indent).Append("</figure>\n");
            }
            else
            {
                sb.Append(indent).Append(svgContent).Append('\n');
            }
            return;
        }

        var src      = img.ResourcePath ?? BuildDataUri(img);
        var alt      = EscapeAttr(img.Description ?? "");
        var sizeAttr = new StringBuilder();
        if (img.WidthMm  > 0) sizeAttr.Append(" width=\"")  .Append(MmToPx(img.WidthMm) .ToString("0", CultureInfo.InvariantCulture)).Append('"');
        if (img.HeightMm > 0) sizeAttr.Append(" height=\"") .Append(MmToPx(img.HeightMm).ToString("0", CultureInfo.InvariantCulture)).Append('"');

        if (img.ShowTitle && !string.IsNullOrEmpty(img.Title))
        {
            sb.Append(indent).Append("<figure").Append(styleAttr).Append(">\n");
            sb.Append(indent).Append("  <img src=\"").Append(EscapeAttr(src)).Append("\" alt=\"")
              .Append(alt).Append('"').Append(sizeAttr).Append(">\n");
            sb.Append(indent).Append("  <figcaption>").Append(EscapeHtml(img.Title!)).Append("</figcaption>\n");
            sb.Append(indent).Append("</figure>\n");
        }
        else
        {
            sb.Append(indent).Append("<img src=\"").Append(EscapeAttr(src)).Append("\" alt=\"")
              .Append(alt).Append('"').Append(sizeAttr).Append(styleAttr).Append(">\n");
        }
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
            sb.Append(EscapeHtml(entry.Text));
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

        var svgBody = BuildShapeSvgBody(shape, wPx, hPx, xhtml: false);
        sb.Append(indent).Append("<figure class=\"pd-shape\"").Append(alignStyle).Append(">\n");
        sb.Append(indent).Append("  <svg width=\"").Append(wPx.ToString("0.#", CultureInfo.InvariantCulture))
          .Append("\" height=\"").Append(hPx.ToString("0.#", CultureInfo.InvariantCulture)).Append("\">");
        sb.Append(svgBody).Append("</svg>\n");
        if (!string.IsNullOrEmpty(shape.LabelText))
            sb.Append(indent).Append("  <figcaption>").Append(EscapeHtml(shape.LabelText)).Append("</figcaption>\n");
        sb.Append(indent).Append("</figure>\n");
    }

    private static string BuildShapeSvgBody(ShapeObject shape, double wPx, double hPx, bool xhtml)
    {
        var stroke = EscapeAttr(shape.StrokeColor);
        var fill   = shape.FillColor is { Length: > 0 } fc ? EscapeAttr(fc) : "none";
        var sw     = shape.StrokeThicknessPt.ToString("0.##", CultureInfo.InvariantCulture);

        return shape.Kind switch
        {
            ShapeKind.Rectangle =>
                $"<rect x=\"0.5\" y=\"0.5\" width=\"{wPx - 1:0.#}\" height=\"{hPx - 1:0.#}\" stroke=\"{stroke}\" stroke-width=\"{sw}\" fill=\"{fill}\"{(xhtml ? "/" : "")}>"
                + (xhtml ? "" : "</rect>"),

            ShapeKind.RoundedRect =>
                $"<rect x=\"0.5\" y=\"0.5\" width=\"{wPx - 1:0.#}\" height=\"{hPx - 1:0.#}\" rx=\"{MmToPx(shape.CornerRadiusMm):0.#}\" stroke=\"{stroke}\" stroke-width=\"{sw}\" fill=\"{fill}\"{(xhtml ? "/" : "")}>"
                + (xhtml ? "" : "</rect>"),

            ShapeKind.Ellipse =>
                $"<ellipse cx=\"{wPx / 2:0.#}\" cy=\"{hPx / 2:0.#}\" rx=\"{(wPx - 1) / 2:0.#}\" ry=\"{(hPx - 1) / 2:0.#}\" stroke=\"{stroke}\" stroke-width=\"{sw}\" fill=\"{fill}\"{(xhtml ? "/" : "")}>"
                + (xhtml ? "" : "</ellipse>"),

            ShapeKind.Line when shape.Points.Count >= 2 =>
                $"<line x1=\"{MmToPx(shape.Points[0].X):0.#}\" y1=\"{MmToPx(shape.Points[0].Y):0.#}\" x2=\"{MmToPx(shape.Points[^1].X):0.#}\" y2=\"{MmToPx(shape.Points[^1].Y):0.#}\" stroke=\"{stroke}\" stroke-width=\"{sw}\" fill=\"none\"{(xhtml ? "/" : "")}>"
                + (xhtml ? "" : "</line>"),

            ShapeKind.Polyline when shape.Points.Count >= 2 =>
                $"<polyline points=\"{PointsToSvg(shape.Points)}\" stroke=\"{stroke}\" stroke-width=\"{sw}\" fill=\"none\"{(xhtml ? "/" : "")}>"
                + (xhtml ? "" : "</polyline>"),

            ShapeKind.Polygon or ShapeKind.Triangle when shape.Points.Count >= 3 =>
                $"<polygon points=\"{PointsToSvg(shape.Points)}\" stroke=\"{stroke}\" stroke-width=\"{sw}\" fill=\"{fill}\"{(xhtml ? "/" : "")}>"
                + (xhtml ? "" : "</polygon>"),

            ShapeKind.Spline or ShapeKind.ClosedSpline when shape.Points.Count >= 2 =>
                BuildSplinePath(shape, stroke, fill, sw, xhtml),

            ShapeKind.RegularPolygon =>
                BuildRegularPolygonSvg(shape, wPx, hPx, stroke, fill, sw, xhtml),

            ShapeKind.Star =>
                BuildStarSvg(shape, wPx, hPx, stroke, fill, sw, xhtml),

            _ =>
                $"<rect x=\"0.5\" y=\"0.5\" width=\"{wPx - 1:0.#}\" height=\"{hPx - 1:0.#}\" stroke=\"{stroke}\" stroke-width=\"{sw}\" fill=\"{fill}\"{(xhtml ? "/" : "")}>"
                + (xhtml ? "" : "</rect>"),
        };
    }

    private static string PointsToSvg(IList<ShapePoint> pts)
        => string.Join(" ", pts.Select(p => $"{MmToPx(p.X):0.#},{MmToPx(p.Y):0.#}"));

    private static string BuildSplinePath(ShapeObject shape, string stroke, string fill, string sw, bool xhtml)
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
        var sc = xhtml ? "/" : "";
        sb.Append($"\" stroke=\"{stroke}\" stroke-width=\"{sw}\" fill=\"{fill}\"{sc}>");
        if (!xhtml) sb.Append("</path>");
        return sb.ToString();
    }

    private static string BuildRegularPolygonSvg(ShapeObject shape, double wPx, double hPx, string stroke, string fill, string sw, bool xhtml)
    {
        int sides = Math.Max(3, shape.SideCount);
        double cx = wPx / 2, cy = hPx / 2, rx = (wPx - 1) / 2, ry = (hPx - 1) / 2;
        var pts = string.Join(" ", Enumerable.Range(0, sides).Select(i =>
        {
            double a = 2 * Math.PI * i / sides - Math.PI / 2;
            return $"{(cx + rx * Math.Cos(a)):0.#},{(cy + ry * Math.Sin(a)):0.#}";
        }));
        var sc = xhtml ? "/" : "";
        return $"<polygon points=\"{pts}\" stroke=\"{stroke}\" stroke-width=\"{sw}\" fill=\"{fill}\"{sc}>"
               + (xhtml ? "" : "</polygon>");
    }

    private static string BuildStarSvg(ShapeObject shape, double wPx, double hPx, string stroke, string fill, string sw, bool xhtml)
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
        var sc = xhtml ? "/" : "";
        return $"<polygon points=\"{pts}\" stroke=\"{stroke}\" stroke-width=\"{sw}\" fill=\"{fill}\"{sc}>"
               + (xhtml ? "" : "</polygon>");
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
          .Append(EscapeHtml(opq.DisplayLabel))
          .Append("</div>\n");
    }

    private static string BuildDataUri(ImageBlock img)
    {
        if (img.Data.Length == 0) return "";
        var b64 = Convert.ToBase64String(img.Data);
        return $"data:{img.MediaType};base64,{b64}";
    }

    private static double MmToPx(double mm) => mm * 96.0 / 25.4;

    // ── 인라인 렌더링 ─────────────────────────────────────────────────

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

        // LaTeX 수식 — MathJax 호환 \(...\) / \[...\] 구문.
        if (run.LatexSource is { Length: > 0 } latex)
        {
            var escaped = EscapeHtml(latex);
            return run.IsDisplayEquation
                ? $"<span class=\"pd-math pd-math-display\">\\[{escaped}\\]</span>"
                : $"<span class=\"pd-math\">\\({escaped}\\)</span>";
        }

        // 이모지 — data-pd-emoji 속성으로 키 보존, 표시 이름을 텍스트로.
        if (run.EmojiKey is { Length: > 0 } emojiKey)
        {
            var parts = emojiKey.Split('_', 2);
            var name  = parts.Length == 2 ? parts[1] : emojiKey;
            return $"<span class=\"pd-emoji\" data-pd-emoji=\"{EscapeAttr(emojiKey)}\" title=\"{EscapeAttr(name)}\">{EscapeHtml(name)}</span>";
        }

        // 인라인 필드 — pd-field-{type} 클래스 + 현재 값 placeholder.
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
            return $"<span class=\"pd-field pd-field-{cls}\">{EscapeHtml(placeholder)}</span>";
        }

        var s    = run.Style;
        var text = EscapeHtml(run.Text).Replace("\n", "<br>");

        bool isMono = !string.IsNullOrEmpty(s.FontFamily) &&
                      s.FontFamily.Contains("monospace", StringComparison.OrdinalIgnoreCase);

        // 인라인 스타일 (font-family 제외 — 모노스페이스는 <code> 로 표현).
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
        // 11 = 기본값 — 명시 변경 시에만 출력.
        if (Math.Abs(s.FontSizePt - 11) > 0.01 && s.FontSizePt > 0)
            parts.Add($"font-size:{s.FontSizePt.ToString("0.##", CultureInfo.InvariantCulture)}pt");
        if (s.Foreground is { } fg) parts.Add($"color:{ColorHex(fg)}");
        if (s.Background is { } bg) parts.Add($"background-color:{ColorHex(bg)}");
        // Overline 은 semantic HTML 태그 없음 → CSS로만 표현.
        if (s.Overline) parts.Add("text-decoration:overline");
        // 한글 조판 — 장평(WidthPercent) / 자간(LetterSpacingPx).
        if (Math.Abs(s.WidthPercent - 100) > 0.5)
            parts.Add($"transform:scaleX({(s.WidthPercent / 100.0).ToString("0.###", CultureInfo.InvariantCulture)});display:inline-block");
        if (Math.Abs(s.LetterSpacingPx) > 0.01)
            parts.Add($"letter-spacing:{s.LetterSpacingPx.ToString("0.##", CultureInfo.InvariantCulture)}px");
        return string.Join(';', parts);
    }

    private static string ColorHex(Color c) => $"#{c.R:X2}{c.G:X2}{c.B:X2}";

    // ── 이스케이프 ────────────────────────────────────────────────────

    private static string EscapeHtml(string text)
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
                default:  sb.Append(ch);       break;
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
                case '\'': sb.Append("&#39;");  break;
                default:   sb.Append(ch);       break;
            }
        }
        return sb.ToString();
    }
}

using System.Globalization;
using System.Text;
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

            sb.Append("<!DOCTYPE html>\n");
            sb.Append("<html lang=\"ko\">\n");
            sb.Append("<head>\n");
            sb.Append("  <meta charset=\"utf-8\">\n");
            sb.Append("  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n");
            sb.Append("  <meta name=\"generator\" content=\"PolyDonky\">\n");
            sb.Append("  <title>").Append(EscapeHtml(docTitle)).Append("</title>\n");
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
                case Paragraph para: WriteParagraph(sb, para, indent, notes); break;
                case Table table:    WriteTable(sb, table, indent, notes);     break;
                case ImageBlock img: WriteImage(sb, img, indent);              break;
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

        if (p.Style.Outline > OutlineLevel.Body)
        {
            int lvl = (int)p.Style.Outline;
            var styleAttr = ParagraphStyleAttr(p.Style);
            sb.Append(indent).Append('<').Append('h').Append(lvl).Append(styleAttr).Append('>');
            sb.Append(RenderRuns(p.Runs, notes));
            sb.Append("</h").Append(lvl).Append(">\n");
            return;
        }

        var pStyleAttr = ParagraphStyleAttr(p.Style);
        sb.Append(indent).Append("<p").Append(pStyleAttr).Append('>');
        sb.Append(RenderRuns(p.Runs, notes));
        sb.Append("</p>\n");
    }

    private static string ParagraphStyleAttr(ParagraphStyle s)
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

        return parts.Count == 0 ? "" : $" style=\"{string.Join(';', parts)}\"";
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
        var src      = img.ResourcePath ?? BuildDataUri(img);
        var alt      = EscapeAttr(img.Description ?? "");
        var sizeAttr = new StringBuilder();
        if (img.WidthMm  > 0) sizeAttr.Append(" width=\"")  .Append(MmToPx(img.WidthMm) .ToString("0", CultureInfo.InvariantCulture)).Append('"');
        if (img.HeightMm > 0) sizeAttr.Append(" height=\"") .Append(MmToPx(img.HeightMm).ToString("0", CultureInfo.InvariantCulture)).Append('"');

        var imgStyle = BuildImageStyle(img);
        var styleAttr = imgStyle.Length > 0 ? $" style=\"{imgStyle}\"" : "";

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
        // FootnoteId/EndnoteId 가 있지만 notes 맵이 없는 경우 — 참조만 빈 텍스트로 무시.
        if (run.FootnoteId is { Length: > 0 } || run.EndnoteId is { Length: > 0 })
            return string.Empty;

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

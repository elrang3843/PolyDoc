using System.Globalization;
using System.Text;
using PolyDonky.Core;

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

            sb.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>\n");
            sb.Append("<!DOCTYPE html>\n");
            sb.Append("<html xmlns=\"http://www.w3.org/1999/xhtml\" lang=\"ko\">\n");
            sb.Append("<head>\n");
            sb.Append("  <meta charset=\"utf-8\"/>\n");
            sb.Append("  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\"/>\n");
            sb.Append("  <meta name=\"generator\" content=\"PolyDonky\"/>\n");
            sb.Append("  <title>").Append(EscapeText(docTitle)).Append("</title>\n");
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
        var style = p.Style.Clone();
        style.QuoteLevel = Math.Max(0, p.Style.QuoteLevel - 1);
        return new Paragraph { StyleId = p.StyleId, Style = style, Runs = p.Runs };
    }

    private static void WriteParagraph(StringBuilder sb, Paragraph p, string indent, NoteNums? notes = null)
    {
        if (p.Style.IsThematicBreak)
        {
            sb.Append(indent).Append("<hr/>\n");
            return;
        }

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

        if (p.Style.Outline > OutlineLevel.Body)
        {
            int lvl = (int)p.Style.Outline;
            sb.Append(indent).Append('<').Append('h').Append(lvl).Append(ParagraphStyleAttr(p.Style)).Append('>');
            sb.Append(RenderRuns(p.Runs, notes));
            sb.Append("</h").Append(lvl).Append(">\n");
            return;
        }

        sb.Append(indent).Append("<p").Append(ParagraphStyleAttr(p.Style)).Append('>');
        sb.Append(RenderRuns(p.Runs, notes));
        sb.Append("</p>\n");
    }

    private static string ParagraphStyleAttr(ParagraphStyle s)
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

        return parts.Count == 0 ? "" : $" style=\"{string.Join(';', parts)}\"";
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
        var src = img.ResourcePath ?? BuildDataUri(img);
        var alt = EscapeAttr(img.Description ?? "");
        var size = new StringBuilder();
        if (img.WidthMm  > 0) size.Append(" width=\"") .Append(MmToPx(img.WidthMm) .ToString("0", CultureInfo.InvariantCulture)).Append('"');
        if (img.HeightMm > 0) size.Append(" height=\"").Append(MmToPx(img.HeightMm).ToString("0", CultureInfo.InvariantCulture)).Append('"');

        var styleAttr = BuildImageStyle(img);
        var styleStr  = string.IsNullOrEmpty(styleAttr) ? "" : $" style=\"{styleAttr}\"";

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

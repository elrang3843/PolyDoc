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
///   - ThematicBreakBlock   → &lt;hr&gt;
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

        var page = document.Sections.Count > 0 ? document.Sections[0].Page : new PageSettings();
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
            WriteExtraPageMeta(sb, page);
            WriteStyleBlock(sb, document.Styles, page);
            sb.Append("</head>\n");
            sb.Append("<body>\n");

            // 머리말 (Section[0].Page.Header)
            if (!page.Header.IsEmpty)
                WriteHeaderFooter(sb, page.Header, "header", indent, notes);
        }

        foreach (var section in document.Sections)
            WriteBlocks(sb, section.Blocks, indent, notes);

        if (fullDocument)
        {
            if (notes.HasNotes)
                WriteNoteSections(sb, document, notes, indent);

            // 꼬리말 (Section[0].Page.Footer)
            if (!page.Footer.IsEmpty)
                WriteHeaderFooter(sb, page.Footer, "footer", indent, notes);

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
        // 같은 레벨(같은 stacking context) 도형들에 대해 z-index 맵을 계산.
        // 컨테이너(블록쿼트·글상자 본문·표 셀)는 각자 별개 stacking context 이므로
        // 재귀 WriteBlocks 호출은 자기 레벨의 맵을 새로 계산한다.
        var shapeZ = blocks.OfType<ShapeObject>().Any()
            ? ShapeOrdering.ComputeZIndexMap(blocks.OfType<ShapeObject>())
            : null;

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
                case ThematicBreakBlock thb:
                {
                    var hrStyle = new List<string>();
                    // border-top: 두께·선종류·색상. 두께가 0 이면 1px 기본; 색상이 null 이면 #000000.
                    double thickPx = thb.ThicknessPt > 0 ? thb.ThicknessPt * 96.0 / 72.0 : 1.0;
                    string lineKw = thb.LineStyle switch
                    {
                        ThematicLineStyle.Dashed  => "dashed",
                        ThematicLineStyle.Dotted  => "dotted",
                        ThematicLineStyle.Double  => "double",
                        ThematicLineStyle.DashDot => "dashed",   // CSS 표준엔 dash-dot 없음 → dashed 로 근사
                        _                         => "solid",
                    };
                    string lineColor = thb.LineColor ?? "#000000";
                    // 두께/스타일/색상이 모두 기본이면 border-top 출력 생략 — UA 기본 hr 사용.
                    bool hasBorderInfo = thb.ThicknessPt > 0
                                      || thb.LineStyle != ThematicLineStyle.Solid
                                      || thb.LineColor is not null;
                    if (hasBorderInfo)
                    {
                        // border:0 으로 기본 hr 의 양쪽/inset 효과 제거 후 border-top 만 명시.
                        hrStyle.Add("border:0");
                        hrStyle.Add($"border-top:{thickPx.ToString("0.##", CultureInfo.InvariantCulture)}px {lineKw} {lineColor}");
                    }
                    if (thb.MarginPt > 0)
                    {
                        var marginPx = thb.MarginPt * 96.0 / 72.0;
                        hrStyle.Add($"margin:{marginPx:F0}px 0");
                    }
                    var styleAttr = hrStyle.Count > 0 ? $" style=\"{string.Join(';', hrStyle)}\"" : "";
                    sb.Append(indent).Append($"<hr{styleAttr}>\n");
                    break;
                }
                case Paragraph para:     WriteParagraph(sb, para, indent, notes); break;
                case Table table:        WriteTable(sb, table, indent, notes);    break;
                case ImageBlock img:     WriteImage(sb, img, indent);             break;
                case TocBlock toc:       WriteToc(sb, toc, indent);               break;
                case ShapeObject shape:
                    WriteShape(sb, shape, indent, shapeZ?.GetValueOrDefault(shape) ?? 0);
                    break;
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
            },
            Runs    = p.Runs,
        };
        return copy;
    }

    private static void WriteParagraph(StringBuilder sb, Paragraph p, string indent, NoteNums? notes = null)
    {
        if (p.Style.CodeLanguage is not null)
        {
            var code = EscapeHtml(p.GetPlainText());
            var preClass = p.Style.ShowLineNumbers ? " class=\"line-numbers\"" : "";
            var langAttr = p.Style.CodeLanguage.Length > 0
                ? $" class=\"language-{EscapeAttr(p.Style.CodeLanguage)}\""
                : "";
            if (p.Style.ShowLineNumbers)
            {
                // 줄 번호 재현을 위해 각 줄을 <span>으로 감싸 출력
                var lines = code.Split('\n');
                var spanLines = string.Join("\n", lines.Select(l => $"<span>{l}</span>"));
                sb.Append(indent).Append("<pre").Append(preClass).Append("><code").Append(langAttr).Append('>')
                  .Append(spanLines).Append("</code></pre>\n");
            }
            else
            {
                sb.Append(indent).Append("<pre><code").Append(langAttr).Append('>')
                  .Append(code).Append("</code></pre>\n");
            }
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

        if (s.BorderBottomPt > 0)
        {
            var bColor = s.BorderBottomColor is { Length: > 0 } ? s.BorderBottomColor : "#CCCCCC";
            parts.Add($"border-bottom:{s.BorderBottomPt.ToString("0.##", CultureInfo.InvariantCulture)}pt solid {bColor}");
        }

        return parts;
    }

    private static string ParagraphStyleAttr(ParagraphStyle s)
    {
        var parts = BuildParagraphCssParts(s);
        return parts.Count == 0 ? "" : $" style=\"{string.Join(';', parts)}\"";
    }

    /// <summary>기본값과 다른 PageSettings 필드를 pd-page-* meta 태그로 직렬화.</summary>
    private static void WriteExtraPageMeta(StringBuilder sb, PageSettings page)
    {
        if (page.PaperColor is { Length: > 0 } pc)
            sb.Append("  <meta name=\"pd-paper-color\" content=\"").Append(EscapeAttr(pc)).Append("\">\n");
        if (Math.Abs(page.MarginHeaderMm - 10) > 0.01)
            sb.Append("  <meta name=\"pd-margin-header\" content=\"").Append(FmtNum(page.MarginHeaderMm)).Append("mm\">\n");
        if (Math.Abs(page.MarginFooterMm - 10) > 0.01)
            sb.Append("  <meta name=\"pd-margin-footer\" content=\"").Append(FmtNum(page.MarginFooterMm)).Append("mm\">\n");
        if (page.ColumnCount > 1)
        {
            sb.Append("  <meta name=\"pd-column-count\" content=\"").Append(page.ColumnCount).Append("\">\n");
            if (Math.Abs(page.ColumnGapMm - 8) > 0.01)
                sb.Append("  <meta name=\"pd-column-gap\" content=\"").Append(FmtNum(page.ColumnGapMm)).Append("mm\">\n");
            if (page.ColumnWidthsMm is { Count: > 0 } cw)
            {
                var joined = string.Join(',', cw.Select(w => w.ToString("0.##", CultureInfo.InvariantCulture)));
                sb.Append("  <meta name=\"pd-column-widths\" content=\"").Append(joined).Append("mm\">\n");
            }
            if (!page.ColumnDividerVisible)
                sb.Append("  <meta name=\"pd-column-divider\" content=\"none\">\n");
            else if (page.ColumnDividerStyle != ColumnDividerStyle.Dashed
                  || page.ColumnDividerColor != "#888888"
                  || Math.Abs(page.ColumnDividerThicknessPt - 0.7) > 0.01)
            {
                sb.Append("  <meta name=\"pd-column-divider\" content=\"")
                  .Append(page.ColumnDividerStyle).Append(' ')
                  .Append(FmtNum(page.ColumnDividerThicknessPt)).Append("pt ")
                  .Append(EscapeAttr(page.ColumnDividerColor)).Append("\">\n");
            }
        }
        if (page.PageNumberStart != 1)
            sb.Append("  <meta name=\"pd-page-number-start\" content=\"").Append(page.PageNumberStart).Append("\">\n");
        if (page.TextOrientation != TextOrientation.Horizontal)
            sb.Append("  <meta name=\"pd-text-orientation\" content=\"vertical\">\n");
        if (page.TextProgression != TextProgression.Rightward)
            sb.Append("  <meta name=\"pd-text-progression\" content=\"leftward\">\n");
        if (page.DifferentFirstPage)
            sb.Append("  <meta name=\"pd-different-first-page\" content=\"true\">\n");
        if (page.DifferentOddEven)
            sb.Append("  <meta name=\"pd-different-odd-even\" content=\"true\">\n");
    }

    /// <summary>머리말/꼬리말 좌·중·우 슬롯을 &lt;header&gt;/&lt;footer&gt; 로 직렬화.</summary>
    private static void WriteHeaderFooter(StringBuilder sb, HeaderFooterContent hf, string tag, string indent, NoteNums? notes)
    {
        sb.Append(indent).Append('<').Append(tag).Append(" class=\"pd-").Append(tag).Append("\">\n");
        WriteHeaderFooterSlot(sb, hf.Left,   "left",   indent + "  ", notes);
        WriteHeaderFooterSlot(sb, hf.Center, "center", indent + "  ", notes);
        WriteHeaderFooterSlot(sb, hf.Right,  "right",  indent + "  ", notes);
        sb.Append(indent).Append("</").Append(tag).Append(">\n");
    }

    private static void WriteHeaderFooterSlot(StringBuilder sb, HeaderFooterSlot slot, string slotKey, string indent, NoteNums? notes)
    {
        if (slot.IsEmpty) return;
        sb.Append(indent).Append("<div class=\"pd-hf-").Append(slotKey).Append("\">\n");
        // 머리말/꼬리말의 단락은 일반 본문과 동일한 inline 처리만 사용 — 표/이미지 등은 비대상.
        foreach (var p in slot.Paragraphs)
        {
            sb.Append(indent).Append("  <p>");
            sb.Append(RenderRuns(p.Runs, notes));
            sb.Append("</p>\n");
        }
        sb.Append(indent).Append("</div>\n");
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

        // 묶음의 모든 항목이 마커를 숨기도록 표시돼 있으면 list-style-type:none 으로 직렬화.
        // 체크리스트(ListMarker.Checked != null) 는 별도 — input checkbox 로 표현하므로 마커는 항상 표시.
        bool allHideBullet = true;
        for (int k = from; k < to; k++)
        {
            var lmk = ((Paragraph)blocks[k]).Style.ListMarker;
            if (lmk is null || !lmk.HideBullet || lmk.Checked is not null)
            {
                allHideBullet = false;
                break;
            }
        }
        string listStyleAttr = allHideBullet ? " style=\"list-style-type:none;padding-left:0\"" : "";

        sb.Append(indent).Append('<').Append(tag).Append(startAttr).Append(listStyleAttr).Append(">\n");
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
        {
            var bColor = t.BorderColor is { Length: > 0 } ? t.BorderColor : "#C8C8C8";
            tblStyle.Add($"border:{t.BorderThicknessPt.ToString("0.##", CultureInfo.InvariantCulture)}pt solid {bColor}");
            tblStyle.Add("border-collapse:collapse");
        }
        if (t.OuterMarginTopMm > 0 || t.OuterMarginBottomMm > 0)
            tblStyle.Add($"margin:{FmtMm(t.OuterMarginTopMm)} 0 {FmtMm(t.OuterMarginBottomMm)} 0");

        var tblStyleAttr = tblStyle.Count > 0
            ? $" style=\"{string.Join(';', tblStyle)}\""
            : "";

        sb.Append(indent).Append("<table").Append(tblStyleAttr).Append(">\n");

        if (!string.IsNullOrEmpty(t.Caption))
            sb.Append(indent).Append("  <caption>")
              .Append(EscapeHtml(t.Caption))
              .Append("</caption>\n");

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
        var trStyleAttr = row.HeightMm > 0 ? $" style=\"height:{FmtMm(row.HeightMm)}\"" : "";
        sb.Append(indent).Append("<tr").Append(trStyleAttr).Append(">\n");
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
            // 셀 안의 블록을 어떻게 렌더할지 결정:
            //   - 단순 단락(헤딩/리스트/표 등 구조 없음) 만 있으면 인라인 + <br> 구분 — 가독성·간결성.
            //   - 그 외(중첩 표, 리스트, 헤딩 등) 는 전체 블록 렌더로 폴백 — 구조 보존.
            bool needsBlockRender = false;
            foreach (var b in cell.Blocks)
            {
                if (b is Paragraph p)
                {
                    if (p.Style.ListMarker is not null
                        || p.Style.Outline != OutlineLevel.Body
                        || p.Style.CodeLanguage is { Length: > 0 }
                        || p.Style.QuoteLevel > 0)
                    {
                        needsBlockRender = true; break;
                    }
                }
                else if (b is not null)
                {
                    // Table / ImageBlock / ShapeObject / TextBoxObject / etc.
                    needsBlockRender = true; break;
                }
            }

            if (needsBlockRender)
            {
                sb.Append('\n');
                WriteBlocks(sb, cell.Blocks, indent + "    ", notes);
                sb.Append(indent).Append("  ");
            }
            else
            {
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

        if (img.BorderThicknessPt > 0)
        {
            var bColor = img.BorderColor is { Length: > 0 } ? img.BorderColor : "#888888";
            parts.Add($"border:{img.BorderThicknessPt.ToString("0.##", CultureInfo.InvariantCulture)}pt solid {bColor}");
        }

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

    private static void WriteShape(StringBuilder sb, ShapeObject shape, string indent, int effectiveZ = 0)
    {
        var wPx = MmToPx(shape.WidthMm  > 0 ? shape.WidthMm  : 40);
        var hPx = MmToPx(shape.HeightMm > 0 ? shape.HeightMm : 30);

        // 정렬·여백·테두리 표시 — Inline 모드에서만 의미가 있음.
        var figStyleParts = new List<string>(6);
        if (shape.WrapMode == ImageWrapMode.Inline)
        {
            switch (shape.HAlign)
            {
                case ImageHAlign.Center:
                    figStyleParts.Add("display:block");
                    figStyleParts.Add("margin-left:auto");
                    figStyleParts.Add("margin-right:auto");
                    break;
                case ImageHAlign.Right:
                    figStyleParts.Add("display:block");
                    figStyleParts.Add("margin-left:auto");
                    break;
            }
        }
        if (shape.MarginTopMm    > 0) figStyleParts.Add($"margin-top:{FmtMm(shape.MarginTopMm)}");
        if (shape.MarginBottomMm > 0) figStyleParts.Add($"margin-bottom:{FmtMm(shape.MarginBottomMm)}");
        // z-index 가 0 이 아니면 CSS z-index 출력. position:static 에서는 z-index 가 동작하지 않으므로
        // position:relative 를 함께 부여한다(레이아웃에는 영향 없음 — left/top 미지정).
        if (effectiveZ != 0)
        {
            figStyleParts.Add("position:relative");
            figStyleParts.Add($"z-index:{effectiveZ}");
        }
        var alignStyle = figStyleParts.Count > 0 ? $" style=\"{string.Join(";", figStyleParts)}\"" : "";

        // 무손실 라운드트립을 위한 data-pd-* 속성 — geometry/SVG 만으로 복원 불가능한 필드를
        // 명시적으로 보존한다.
        var dataAttrs = new StringBuilder(128);
        dataAttrs.Append(" data-pd-kind=\"").Append(shape.Kind).Append('"');
        if (shape.StrokeDash != StrokeDash.Solid)
            dataAttrs.Append(" data-pd-stroke-dash=\"").Append(shape.StrokeDash).Append('"');
        if (shape.StartArrow != ShapeArrow.None)
            dataAttrs.Append(" data-pd-start-arrow=\"").Append(shape.StartArrow).Append('"');
        if (shape.EndArrow != ShapeArrow.None)
            dataAttrs.Append(" data-pd-end-arrow=\"").Append(shape.EndArrow).Append('"');
        if (shape.EndShapeSizeMm > 0)
            dataAttrs.Append(" data-pd-end-shape-size=\"").Append(FmtNum(shape.EndShapeSizeMm)).Append("mm\"");
        if (shape.Kind is ShapeKind.RegularPolygon or ShapeKind.Star)
            dataAttrs.Append(" data-pd-side-count=\"").Append(shape.SideCount).Append('"');
        if (shape.Kind == ShapeKind.Star)
            dataAttrs.Append(" data-pd-inner-radius-ratio=\"").Append(FmtNum(shape.InnerRadiusRatio)).Append('"');
        if (shape.Kind == ShapeKind.RoundedRect && shape.CornerRadiusMm > 0)
            dataAttrs.Append(" data-pd-corner-radius=\"").Append(FmtNum(shape.CornerRadiusMm)).Append("mm\"");
        if (shape.RotationAngleDeg != 0)
            dataAttrs.Append(" data-pd-rotation=\"").Append(FmtNum(shape.RotationAngleDeg)).Append("deg\"");
        if (Math.Abs(shape.FillOpacity - 1.0) > 0.001)
            dataAttrs.Append(" data-pd-fill-opacity=\"").Append(FmtNum(shape.FillOpacity)).Append('"');
        if (shape.ZOrder != 0)
            dataAttrs.Append(" data-pd-z-order=\"").Append(shape.ZOrder).Append('"');
        if (shape.WrapMode != ImageWrapMode.Inline)
            dataAttrs.Append(" data-pd-wrap-mode=\"").Append(shape.WrapMode).Append('"');
        if (shape.WrapMode is ImageWrapMode.InFrontOfText or ImageWrapMode.BehindText)
        {
            dataAttrs.Append(" data-pd-anchor-page=\"").Append(shape.AnchorPageIndex).Append('"');
            dataAttrs.Append(" data-pd-overlay-x=\"").Append(FmtNum(shape.OverlayXMm)).Append("mm\"");
            dataAttrs.Append(" data-pd-overlay-y=\"").Append(FmtNum(shape.OverlayYMm)).Append("mm\"");
        }

        var svgBody    = BuildShapeSvgBody(shape, wPx, hPx, xhtml: false);
        var defs       = BuildShapeSvgDefs(shape, xhtml: false);
        var rootXform  = shape.RotationAngleDeg != 0
            ? $" transform=\"rotate({FmtNum(shape.RotationAngleDeg)} {wPx / 2:0.##} {hPx / 2:0.##})\""
            : "";

        sb.Append(indent).Append("<figure class=\"pd-shape\"").Append(dataAttrs).Append(alignStyle).Append(">\n");
        sb.Append(indent).Append("  <svg xmlns=\"http://www.w3.org/2000/svg\" width=\"")
          .Append(wPx.ToString("0.#", CultureInfo.InvariantCulture))
          .Append("\" height=\"").Append(hPx.ToString("0.#", CultureInfo.InvariantCulture)).Append("\">");
        if (defs.Length > 0) sb.Append(defs);
        sb.Append("<g").Append(rootXform).Append('>');
        sb.Append(svgBody).Append("</g></svg>\n");
        if (!string.IsNullOrEmpty(shape.LabelText))
        {
            var labelStyle = BuildShapeLabelStyle(shape);
            var styleAttr  = labelStyle.Length > 0 ? $" style=\"{labelStyle}\"" : "";
            var capData = new StringBuilder();
            if (shape.LabelVAlign != ShapeLabelVAlign.Middle)
                capData.Append(" data-pd-valign=\"").Append(shape.LabelVAlign).Append('"');
            if (shape.LabelOffsetXMm != 0)
                capData.Append(" data-pd-offset-x=\"").Append(FmtNum(shape.LabelOffsetXMm)).Append("mm\"");
            if (shape.LabelOffsetYMm != 0)
                capData.Append(" data-pd-offset-y=\"").Append(FmtNum(shape.LabelOffsetYMm)).Append("mm\"");
            sb.Append(indent).Append("  <figcaption").Append(capData).Append(styleAttr).Append('>')
              .Append(EscapeHtml(shape.LabelText)).Append("</figcaption>\n");
        }
        sb.Append(indent).Append("</figure>\n");
    }

    private static string FmtNum(double v) => v.ToString("0.###", CultureInfo.InvariantCulture);

    /// <summary>도형의 stroke-dasharray, marker-start/end 정의 + arrow 마커 등 시각 보정을 위한 SVG &lt;defs&gt;.</summary>
    private static string BuildShapeSvgDefs(ShapeObject shape, bool xhtml)
    {
        bool needsStart = shape.StartArrow != ShapeArrow.None;
        bool needsEnd   = shape.EndArrow   != ShapeArrow.None;
        if (!needsStart && !needsEnd) return "";

        var sb = new StringBuilder("<defs>");
        if (needsStart) sb.Append(BuildArrowMarker("pd-arr-s", shape.StartArrow, shape.StrokeColor, reverse: true,  xhtml));
        if (needsEnd)   sb.Append(BuildArrowMarker("pd-arr-e", shape.EndArrow,   shape.StrokeColor, reverse: false, xhtml));
        sb.Append("</defs>");
        return sb.ToString();
    }

    private static string BuildArrowMarker(string id, ShapeArrow kind, string color, bool reverse, bool xhtml)
    {
        var clr  = EscapeAttr(string.IsNullOrEmpty(color) ? "#000000" : color);
        var refX = reverse ? "0" : "10";
        var orient = reverse ? "auto-start-reverse" : "auto";
        var sc = xhtml ? "/" : "";
        var inner = kind switch
        {
            ShapeArrow.Open    => $"<path d=\"M0,0 L10,5 L0,10\" fill=\"none\" stroke=\"{clr}\" stroke-width=\"1\"{sc}>" + (xhtml ? "" : "</path>"),
            ShapeArrow.Filled  => $"<path d=\"M0,0 L10,5 L0,10 Z\" fill=\"{clr}\" stroke=\"none\"{sc}>"                  + (xhtml ? "" : "</path>"),
            ShapeArrow.Diamond => $"<path d=\"M0,5 L5,0 L10,5 L5,10 Z\" fill=\"{clr}\" stroke=\"none\"{sc}>"             + (xhtml ? "" : "</path>"),
            ShapeArrow.Circle  => $"<circle cx=\"5\" cy=\"5\" r=\"4\" fill=\"{clr}\" stroke=\"none\"{sc}>"               + (xhtml ? "" : "</circle>"),
            _                  => "",
        };
        return $"<marker id=\"{id}\" viewBox=\"0 0 10 10\" refX=\"{refX}\" refY=\"5\" markerWidth=\"6\" markerHeight=\"6\" orient=\"{orient}\">"
             + inner + "</marker>";
    }

    private static string BuildShapeLabelStyle(ShapeObject shape)
    {
        var p = new List<string>(8);
        if (shape.LabelFontFamily is { Length: > 0 } ff) p.Add($"font-family:{EscapeAttr(ff)}");
        if (shape.LabelFontSizePt > 0 && Math.Abs(shape.LabelFontSizePt - 10) > 0.01)
            p.Add($"font-size:{FmtNum(shape.LabelFontSizePt)}pt");
        if (shape.LabelBold)   p.Add("font-weight:bold");
        if (shape.LabelItalic) p.Add("font-style:italic");
        if (shape.LabelColor is { Length: > 0 } lc) p.Add($"color:{lc}");
        if (shape.LabelBackgroundColor is { Length: > 0 } lbg) p.Add($"background-color:{lbg}");
        var ta = shape.LabelHAlign switch
        {
            ShapeLabelHAlign.Left   => "left",
            ShapeLabelHAlign.Right  => "right",
            ShapeLabelHAlign.Center => "center",
            _ => null,
        };
        if (ta is not null && shape.LabelHAlign != ShapeLabelHAlign.Center) p.Add($"text-align:{ta}");
        // VAlign / Offset 은 시각적으로 figcaption 으로 표현이 어려워 data-* 로만 보존.
        return string.Join(";", p);
    }

    private static string BuildShapeSvgBody(ShapeObject shape, double wPx, double hPx, bool xhtml)
    {
        var stroke = EscapeAttr(shape.StrokeColor);
        var fill   = shape.FillColor is { Length: > 0 } fc ? EscapeAttr(fc) : "none";
        var sw     = shape.StrokeThicknessPt.ToString("0.##", CultureInfo.InvariantCulture);

        var paint = new StringBuilder($"stroke=\"{stroke}\" stroke-width=\"{sw}\" fill=\"{fill}\"");
        if (shape.StrokeDash != StrokeDash.Solid)
        {
            var dash = shape.StrokeDash switch
            {
                StrokeDash.Dashed  => "6,3",
                StrokeDash.Dotted  => "1,2",
                StrokeDash.DashDot => "6,3,1,3",
                _                  => null,
            };
            if (dash is not null) paint.Append($" stroke-dasharray=\"{dash}\"");
        }
        if (Math.Abs(shape.FillOpacity - 1.0) > 0.001 && fill != "none")
            paint.Append($" fill-opacity=\"{FmtNum(shape.FillOpacity)}\"");

        // Line/Polyline/Spline 류만 끝모양 적용. Polygon/Closed 류는 닫힌 도형이라 의미 없음.
        bool linear = shape.Kind is ShapeKind.Line or ShapeKind.Polyline or ShapeKind.Spline;
        var markers = new StringBuilder();
        if (linear && shape.StartArrow != ShapeArrow.None) markers.Append(" marker-start=\"url(#pd-arr-s)\"");
        if (linear && shape.EndArrow   != ShapeArrow.None) markers.Append(" marker-end=\"url(#pd-arr-e)\"");

        var paintAttrs   = paint.ToString();
        var paintNoFill  = paintAttrs.Replace($"fill=\"{fill}\"", "fill=\"none\"", StringComparison.Ordinal);

        return shape.Kind switch
        {
            ShapeKind.Rectangle =>
                $"<rect x=\"0.5\" y=\"0.5\" width=\"{wPx - 1:0.#}\" height=\"{hPx - 1:0.#}\" {paintAttrs}{(xhtml ? "/" : "")}>"
                + (xhtml ? "" : "</rect>"),

            ShapeKind.RoundedRect =>
                $"<rect x=\"0.5\" y=\"0.5\" width=\"{wPx - 1:0.#}\" height=\"{hPx - 1:0.#}\" rx=\"{MmToPx(shape.CornerRadiusMm):0.#}\" {paintAttrs}{(xhtml ? "/" : "")}>"
                + (xhtml ? "" : "</rect>"),

            ShapeKind.Ellipse =>
                $"<ellipse cx=\"{wPx / 2:0.#}\" cy=\"{hPx / 2:0.#}\" rx=\"{(wPx - 1) / 2:0.#}\" ry=\"{(hPx - 1) / 2:0.#}\" {paintAttrs}{(xhtml ? "/" : "")}>"
                + (xhtml ? "" : "</ellipse>"),

            ShapeKind.Line when shape.Points.Count >= 2 =>
                $"<line x1=\"{MmToPx(shape.Points[0].X):0.#}\" y1=\"{MmToPx(shape.Points[0].Y):0.#}\" x2=\"{MmToPx(shape.Points[^1].X):0.#}\" y2=\"{MmToPx(shape.Points[^1].Y):0.#}\" {paintNoFill}{markers}{(xhtml ? "/" : "")}>"
                + (xhtml ? "" : "</line>"),

            ShapeKind.Polyline when shape.Points.Count >= 2 =>
                $"<polyline points=\"{PointsToSvg(shape.Points)}\" {paintNoFill}{markers}{(xhtml ? "/" : "")}>"
                + (xhtml ? "" : "</polyline>"),

            ShapeKind.Polygon or ShapeKind.Triangle when shape.Points.Count >= 3 =>
                $"<polygon points=\"{PointsToSvg(shape.Points)}\" {paintAttrs}{(xhtml ? "/" : "")}>"
                + (xhtml ? "" : "</polygon>"),

            ShapeKind.Spline or ShapeKind.ClosedSpline when shape.Points.Count >= 2 =>
                BuildSplinePath(shape, paintAttrs, markers.ToString(), xhtml),

            ShapeKind.RegularPolygon =>
                BuildRegularPolygonSvg(shape, wPx, hPx, paintAttrs, xhtml),

            ShapeKind.Star =>
                BuildStarSvg(shape, wPx, hPx, paintAttrs, xhtml),

            _ =>
                $"<rect x=\"0.5\" y=\"0.5\" width=\"{wPx - 1:0.#}\" height=\"{hPx - 1:0.#}\" {paintAttrs}{(xhtml ? "/" : "")}>"
                + (xhtml ? "" : "</rect>"),
        };
    }

    private static string PointsToSvg(IList<ShapePoint> pts)
        => string.Join(" ", pts.Select(p => $"{MmToPx(p.X):0.#},{MmToPx(p.Y):0.#}"));

    private static string BuildSplinePath(ShapeObject shape, string paintAttrs, string markers, bool xhtml)
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
        sb.Append($"\" {paintAttrs}{markers}{sc}>");
        if (!xhtml) sb.Append("</path>");
        return sb.ToString();
    }

    private static string BuildRegularPolygonSvg(ShapeObject shape, double wPx, double hPx, string paintAttrs, bool xhtml)
    {
        int sides = Math.Max(3, shape.SideCount);
        double cx = wPx / 2, cy = hPx / 2, rx = (wPx - 1) / 2, ry = (hPx - 1) / 2;
        var pts = string.Join(" ", Enumerable.Range(0, sides).Select(i =>
        {
            double a = 2 * Math.PI * i / sides - Math.PI / 2;
            return $"{(cx + rx * Math.Cos(a)):0.#},{(cy + ry * Math.Sin(a)):0.#}";
        }));
        var sc = xhtml ? "/" : "";
        return $"<polygon points=\"{pts}\" {paintAttrs}{sc}>"
               + (xhtml ? "" : "</polygon>");
    }

    private static string BuildStarSvg(ShapeObject shape, double wPx, double hPx, string paintAttrs, bool xhtml)
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
        return $"<polygon points=\"{pts}\" {paintAttrs}{sc}>"
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

using PolyDonky.Codecs.Html;
using PolyDonky.Core;

namespace PolyDonky.Codecs.Html.Tests;

public class HtmlTests
{
    // ── Reader ─────────────────────────────────────────────────────

    [Fact]
    public void Reader_HeadingsHaveOutlineLevel()
    {
        var doc = HtmlReader.FromHtml("<h1>A</h1><h2>B</h2><h6>F</h6>");
        var ps  = doc.EnumerateParagraphs().ToList();
        Assert.Equal(OutlineLevel.H1, ps[0].Style.Outline);
        Assert.Equal(OutlineLevel.H2, ps[1].Style.Outline);
        Assert.Equal(OutlineLevel.H6, ps[2].Style.Outline);
    }

    [Fact]
    public void Reader_BoldItalicStrikeAndSubSup()
    {
        var doc = HtmlReader.FromHtml(
            "<p><strong>b</strong> <em>i</em> <s>x</s> H<sub>2</sub>O <sup>2</sup></p>");
        var p   = doc.EnumerateParagraphs().Single();
        Assert.Contains(p.Runs, r => r.Style.Bold          && r.Text == "b");
        Assert.Contains(p.Runs, r => r.Style.Italic        && r.Text == "i");
        Assert.Contains(p.Runs, r => r.Style.Strikethrough && r.Text == "x");
        Assert.Contains(p.Runs, r => r.Style.Subscript     && r.Text == "2");
        Assert.Contains(p.Runs, r => r.Style.Superscript   && r.Text == "2");
    }

    [Fact]
    public void Reader_ListUlAndOl()
    {
        var doc = HtmlReader.FromHtml("<ul><li>a</li><li>b</li></ul><ol><li>1</li><li>2</li></ol>");
        var ps  = doc.EnumerateParagraphs().ToList();
        Assert.Equal(4, ps.Count);
        Assert.Equal(ListKind.Bullet,         ps[0].Style.ListMarker!.Kind);
        Assert.Equal(ListKind.Bullet,         ps[1].Style.ListMarker!.Kind);
        Assert.Equal(ListKind.OrderedDecimal, ps[2].Style.ListMarker!.Kind);
        Assert.Equal(2,                       ps[3].Style.ListMarker!.OrderedNumber);
    }

    [Fact]
    public void Reader_TaskListCheckbox()
    {
        var doc = HtmlReader.FromHtml(
            "<ul><li><input type=\"checkbox\" checked> done</li>" +
            "<li><input type=\"checkbox\"> todo</li></ul>");
        var ps  = doc.EnumerateParagraphs().ToList();
        Assert.True (ps[0].Style.ListMarker!.Checked);
        Assert.False(ps[1].Style.ListMarker!.Checked);
    }

    [Fact]
    public void Reader_BlockquoteNested()
    {
        var doc = HtmlReader.FromHtml("<blockquote><p>1</p><blockquote><p>2</p></blockquote></blockquote>");
        var ps  = doc.EnumerateParagraphs().ToList();
        Assert.Contains(ps, p => p.Style.QuoteLevel == 1 && p.GetPlainText() == "1");
        Assert.Contains(ps, p => p.Style.QuoteLevel == 2 && p.GetPlainText() == "2");
    }

    [Fact]
    public void Reader_HrIsThematicBreak()
    {
        var doc = HtmlReader.FromHtml("<p>before</p><hr><p>after</p>");
        Assert.Contains(doc.Sections[0].Blocks, b => b is ThematicBreakBlock);
    }

    [Fact]
    public void Reader_PreCodeWithLanguageClass()
    {
        var doc = HtmlReader.FromHtml(
            "<pre><code class=\"language-python\">print('hi')</code></pre>");
        var p = doc.EnumerateParagraphs().Single();
        Assert.Equal("python", p.Style.CodeLanguage);
        Assert.Contains("print", p.GetPlainText());
    }

    [Fact]
    public void Reader_LinkStoresUrlAndUnderline()
    {
        var doc = HtmlReader.FromHtml("<p>방문 <a href=\"https://example.com\">사이트</a> 하기</p>");
        var p   = doc.EnumerateParagraphs().Single();
        Assert.Contains(p.Runs, r => r.Url == "https://example.com" && r.Style.Underline);
    }

    [Fact]
    public void Reader_TableHeaderAndCellAlignment()
    {
        var doc = HtmlReader.FromHtml(
            "<table>" +
            "<thead><tr><th>이름</th><th>나이</th></tr></thead>" +
            "<tbody><tr><td>홍길동</td><td style=\"text-align:right\">30</td></tr></tbody>" +
            "</table>");
        var t = doc.Sections[0].Blocks.OfType<PolyDonky.Core.Table>().Single();
        Assert.Equal(2, t.Rows.Count);
        Assert.True (t.Rows[0].IsHeader);
        Assert.Equal(CellTextAlign.Right, t.Rows[1].Cells[1].TextAlign);
    }

    [Fact]
    public void Reader_ImgBecomesImageBlock()
    {
        var doc = HtmlReader.FromHtml("<img src=\"img/a.png\" alt=\"대체\" width=\"200\" height=\"100\">");
        var img = doc.Sections[0].Blocks.OfType<ImageBlock>().Single();
        Assert.Equal("img/a.png", img.ResourcePath);
        Assert.Equal("대체", img.Description);
        Assert.Equal("image/png", img.MediaType);
        Assert.True(img.WidthMm > 0);
    }

    [Fact]
    public void Reader_FigureWithCaptionUsesImageTitle()
    {
        var doc = HtmlReader.FromHtml(
            "<figure><img src=\"x.png\" alt=\"a\"><figcaption>제목</figcaption></figure>");
        var img = doc.Sections[0].Blocks.OfType<ImageBlock>().Single();
        Assert.True(img.ShowTitle);
        Assert.Equal("제목", img.Title);
    }

    [Fact]
    public void Reader_DataUriImageDecodesBytes()
    {
        // 1x1 transparent PNG, base64.
        var pngB64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=";
        var doc = HtmlReader.FromHtml($"<img src=\"data:image/png;base64,{pngB64}\" alt=\"x\">");
        var img = doc.Sections[0].Blocks.OfType<ImageBlock>().Single();
        Assert.Equal("image/png", img.MediaType);
        Assert.NotEmpty(img.Data);
    }

    [Fact]
    public void Reader_SpanInlineStyleParsesColors()
    {
        var doc = HtmlReader.FromHtml(
            "<p><span style=\"color:#FF0000;background-color:rgb(0,255,0);font-weight:bold\">x</span></p>");
        var run = doc.EnumerateParagraphs().Single().Runs.Single(r => r.Text == "x");
        Assert.True(run.Style.Bold);
        Assert.Equal(new Color(0xFF, 0, 0),   run.Style.Foreground);
        Assert.Equal(new Color(0,    0xFF, 0), run.Style.Background);
    }

    [Fact]
    public void Reader_BrInsideParagraphYieldsNewline()
    {
        var doc = HtmlReader.FromHtml("<p>a<br>b</p>");
        var p   = doc.EnumerateParagraphs().Single();
        Assert.Contains(p.Runs, r => r.Text == "\n");
    }

    [Fact]
    public void Reader_CodeInlineUsesMonospace()
    {
        var doc = HtmlReader.FromHtml("<p>call <code>foo()</code></p>");
        var run = doc.EnumerateParagraphs().Single().Runs.Single(r => r.Text == "foo()");
        Assert.Contains("monospace", run.Style.FontFamily, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Reader_IgnoresScriptStyleHead()
    {
        var doc = HtmlReader.FromHtml(
            "<head><title>T</title><style>p{color:red}</style></head>" +
            "<body><script>alert('x')</script><p>본문</p></body>");
        var p = doc.EnumerateParagraphs().Single();
        Assert.Equal("본문", p.GetPlainText());
    }

    // ── Writer ─────────────────────────────────────────────────────

    [Fact]
    public void Writer_FullDocumentEmitsDoctypeAndHead()
    {
        var pd  = new PolyDonkyument();
        var sec = new Section(); pd.Sections.Add(sec);
        sec.Blocks.Add(Paragraph.Of("hi"));
        var html = HtmlWriter.ToHtml(pd);
        Assert.StartsWith("<!DOCTYPE html>", html);
        Assert.Contains("<meta charset=\"utf-8\">", html);
        Assert.Contains("<title>", html);
    }

    [Fact]
    public void Writer_HeadingTags()
    {
        var pd  = new PolyDonkyument();
        var sec = new Section(); pd.Sections.Add(sec);
        var h2 = new Paragraph { Style = { Outline = OutlineLevel.H2 } };
        h2.AddText("section");
        sec.Blocks.Add(h2);

        var html = HtmlWriter.ToHtml(pd, fullDocument: false);
        Assert.Contains("<h2>section</h2>", html);
    }

    [Fact]
    public void Writer_BoldItalicStrikeSubSup()
    {
        var pd  = new PolyDonkyument();
        var sec = new Section(); pd.Sections.Add(sec);
        var p = new Paragraph();
        p.AddText("b", new RunStyle { Bold          = true });
        p.AddText("i", new RunStyle { Italic        = true });
        p.AddText("s", new RunStyle { Strikethrough = true });
        p.AddText("u", new RunStyle { Underline     = true });
        p.AddText("x", new RunStyle { Subscript     = true });
        p.AddText("y", new RunStyle { Superscript   = true });
        sec.Blocks.Add(p);

        var html = HtmlWriter.ToHtml(pd, fullDocument: false);
        Assert.Contains("<strong>b</strong>", html);
        Assert.Contains("<em>i</em>", html);
        Assert.Contains("<s>s</s>", html);
        Assert.Contains("<u>u</u>", html);
        Assert.Contains("<sub>x</sub>", html);
        Assert.Contains("<sup>y</sup>", html);
    }

    [Fact]
    public void Writer_LinksAndCode()
    {
        var pd  = new PolyDonkyument();
        var sec = new Section(); pd.Sections.Add(sec);
        var p = new Paragraph();
        p.Runs.Add(new Run { Text = "site", Style = new RunStyle(), Url = "https://x" });
        p.AddText("foo()", new RunStyle { FontFamily = "Consolas, monospace" });
        sec.Blocks.Add(p);

        var html = HtmlWriter.ToHtml(pd, fullDocument: false);
        Assert.Contains("<a href=\"https://x\">site</a>", html);
        Assert.Contains("<code>foo()</code>", html);
    }

    [Fact]
    public void Writer_BulletAndOrderedList()
    {
        var pd  = new PolyDonkyument();
        var sec = new Section(); pd.Sections.Add(sec);
        var p1 = new Paragraph();
        p1.Style.ListMarker = new ListMarker { Kind = ListKind.Bullet };
        p1.AddText("a");
        var p2 = new Paragraph();
        p2.Style.ListMarker = new ListMarker { Kind = ListKind.OrderedDecimal, OrderedNumber = 1 };
        p2.AddText("1st");
        sec.Blocks.Add(p1); sec.Blocks.Add(p2);

        var html = HtmlWriter.ToHtml(pd, fullDocument: false);
        Assert.Contains("<ul>",  html);
        Assert.Contains("<li>a</li>", html);
        Assert.Contains("<ol>",  html);
        Assert.Contains("<li>1st</li>", html);
    }

    [Fact]
    public void Writer_TaskListEmitsCheckbox()
    {
        var pd  = new PolyDonkyument();
        var sec = new Section(); pd.Sections.Add(sec);
        var p = new Paragraph();
        p.Style.ListMarker = new ListMarker { Kind = ListKind.Bullet, Checked = true };
        p.AddText("done");
        sec.Blocks.Add(p);

        var html = HtmlWriter.ToHtml(pd, fullDocument: false);
        Assert.Contains("<input type=\"checkbox\" disabled checked>", html);
    }

    [Fact]
    public void Writer_BlockquoteNested()
    {
        var pd  = new PolyDonkyument();
        var sec = new Section(); pd.Sections.Add(sec);
        var p1 = new Paragraph { Style = { QuoteLevel = 1 } }; p1.AddText("외");
        var p2 = new Paragraph { Style = { QuoteLevel = 2 } }; p2.AddText("내");
        sec.Blocks.Add(p1); sec.Blocks.Add(p2);

        var html = HtmlWriter.ToHtml(pd, fullDocument: false);
        Assert.Contains("<blockquote>", html);
        // 2단 인용은 두 개의 <blockquote> 가 중첩되어야 함.
        Assert.Equal(2, System.Text.RegularExpressions.Regex.Matches(html, "<blockquote>").Count);
    }

    [Fact]
    public void Writer_PreCodeWithLanguageClass()
    {
        var pd  = new PolyDonkyument();
        var sec = new Section(); pd.Sections.Add(sec);
        var p = new Paragraph { Style = { CodeLanguage = "cs" } };
        p.AddText("var x = 1;");
        sec.Blocks.Add(p);

        var html = HtmlWriter.ToHtml(pd, fullDocument: false);
        Assert.Contains("<pre><code class=\"language-cs\">var x = 1;</code></pre>", html);
    }

    [Fact]
    public void Writer_TableHeaderBodyAlignment()
    {
        var pd  = new PolyDonkyument();
        var sec = new Section(); pd.Sections.Add(sec);
        var t = new PolyDonky.Core.Table();
        t.Columns.Add(new TableColumn());
        t.Columns.Add(new TableColumn());
        var hdr = new TableRow { IsHeader = true };
        var c1 = new TableCell(); c1.Blocks.Add(Paragraph.Of("H1"));
        var c2 = new TableCell { TextAlign = CellTextAlign.Right }; c2.Blocks.Add(Paragraph.Of("H2"));
        hdr.Cells.Add(c1); hdr.Cells.Add(c2);
        var body = new TableRow();
        var b1 = new TableCell(); b1.Blocks.Add(Paragraph.Of("a"));
        var b2 = new TableCell { TextAlign = CellTextAlign.Right }; b2.Blocks.Add(Paragraph.Of("b"));
        body.Cells.Add(b1); body.Cells.Add(b2);
        t.Rows.Add(hdr); t.Rows.Add(body);
        sec.Blocks.Add(t);

        var html = HtmlWriter.ToHtml(pd, fullDocument: false);
        Assert.Contains("<thead>", html);
        Assert.Contains("<tbody>", html);
        Assert.Contains("<th style=\"text-align:right\">H2</th>", html);
        Assert.Contains("<td style=\"text-align:right\">b</td>",  html);
    }

    [Fact]
    public void Writer_ImageEmitsImg()
    {
        var pd  = new PolyDonkyument();
        var sec = new Section(); pd.Sections.Add(sec);
        sec.Blocks.Add(new ImageBlock
        {
            Description  = "alt",
            ResourcePath = "img/foo.png",
            MediaType    = "image/png",
        });

        var html = HtmlWriter.ToHtml(pd, fullDocument: false);
        Assert.Contains("<img src=\"img/foo.png\" alt=\"alt\"", html);
    }

    [Fact]
    public void Writer_FigureWithCaption()
    {
        var pd  = new PolyDonkyument();
        var sec = new Section(); pd.Sections.Add(sec);
        sec.Blocks.Add(new ImageBlock
        {
            ResourcePath = "x.png",
            ShowTitle    = true,
            Title        = "캡션",
        });

        var html = HtmlWriter.ToHtml(pd, fullDocument: false);
        Assert.Contains("<figure>", html);
        Assert.Contains("<figcaption>캡션</figcaption>", html);
    }

    // ── 안전 한도 (대용량 HTML) ─────────────────────────────────────

    [Fact]
    public void Reader_BlockLimit_TruncatesAndAppendsWarning()
    {
        // 50,000 개의 단락이 있는 거대한 HTML — 기본 한도(10,000) 초과해야 함.
        var sb = new System.Text.StringBuilder("<html><body>");
        for (int i = 0; i < 50_000; i++) sb.Append("<p>x</p>");
        sb.Append("</body></html>");

        var doc = HtmlReader.FromHtml(sb.ToString());
        var ps  = doc.EnumerateParagraphs().ToList();

        // 한도 + 마지막 경고 단락 = 약 10,001.
        Assert.True(ps.Count <= 10_001, $"잘림 후 단락 수 {ps.Count} 가 한도(10,001)를 초과");
        Assert.Contains(ps, p => p.GetPlainText().Contains("잘림") || p.GetPlainText().Contains("초과"));
    }

    [Fact]
    public void Reader_CustomMaxBlocks_RespectsCap()
    {
        var sb = new System.Text.StringBuilder("<html><body>");
        for (int i = 0; i < 200; i++) sb.Append("<p>x</p>");
        sb.Append("</body></html>");

        var doc = HtmlReader.FromHtml(sb.ToString(), maxBlocks: 50);
        var ps  = doc.EnumerateParagraphs().ToList();
        Assert.True(ps.Count <= 51, $"단락 수 {ps.Count} 가 50+1 을 초과");
    }

    [Fact]
    public void Writer_EscapesAngleBracketsAndAmpersand()
    {
        var pd  = new PolyDonkyument();
        var sec = new Section(); pd.Sections.Add(sec);
        var p = new Paragraph();
        p.AddText("a < b & c > d");
        sec.Blocks.Add(p);

        var html = HtmlWriter.ToHtml(pd, fullDocument: false);
        Assert.Contains("a &lt; b &amp; c &gt; d", html);
    }

    // ── Round-trip ──────────────────────────────────────────────────

    [Fact]
    public void RoundTrip_PreservesStructure()
    {
        const string source =
            "<!DOCTYPE html><html><body>" +
            "<h1>제목</h1>" +
            "<p>본문 <strong>굵게</strong> <em>기울임</em> <a href=\"https://x\">링크</a></p>" +
            "<ul><li>항목 1</li><li>항목 2</li></ul>" +
            "<blockquote><p>인용</p></blockquote>" +
            "<pre><code class=\"language-py\">print(1)</code></pre>" +
            "<hr>" +
            "<table><thead><tr><th>A</th><th>B</th></tr></thead>" +
            "<tbody><tr><td>1</td><td>2</td></tr></tbody></table>" +
            "</body></html>";

        var doc      = HtmlReader.FromHtml(source);
        var rendered = HtmlWriter.ToHtml(doc, fullDocument: false);
        var reread   = HtmlReader.FromHtml(rendered);

        Assert.Equal(OutlineLevel.H1, reread.EnumerateParagraphs().First().Style.Outline);
        Assert.Contains(reread.EnumerateParagraphs(), p => p.Style.QuoteLevel >= 1);
        Assert.Contains(reread.Sections[0].Blocks, b => b is ThematicBreakBlock);
        Assert.Contains(reread.EnumerateParagraphs(), p => p.Style.CodeLanguage == "py");
        Assert.Contains(reread.EnumerateParagraphs(), p => p.Style.ListMarker?.Kind == ListKind.Bullet);
        Assert.Single(reread.Sections[0].Blocks.OfType<PolyDonky.Core.Table>());
        Assert.Contains(reread.EnumerateParagraphs().SelectMany(p => p.Runs),
            r => r.Url == "https://x");
    }

    // ── 추가 Writer 테스트 ─────────────────────────────────────────

    [Fact]
    public void Writer_FragmentMode_NoDoctype()
    {
        var doc = new PolyDonkyument();
        doc.Sections.Add(new Section());
        doc.Sections[0].Blocks.Add(Paragraph.Of("안녕"));

        var html = HtmlWriter.ToHtml(doc, fullDocument: false);
        Assert.DoesNotContain("<!DOCTYPE", html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("<html", html);
        Assert.Contains("<p>안녕</p>", html);
    }

    [Fact]
    public void Writer_CustomDocumentTitle()
    {
        var doc = new PolyDonkyument();
        doc.Sections.Add(new Section());
        doc.Sections[0].Blocks.Add(Paragraph.Of("본문"));

        var html = HtmlWriter.ToHtml(doc, fullDocument: true, title: "내 문서");
        Assert.Contains("<title>내 문서</title>", html);
    }

    [Fact]
    public void Writer_OverlineEmitsCss()
    {
        var doc = new PolyDonkyument();
        var section = new Section();
        doc.Sections.Add(section);
        var p = new Paragraph();
        p.AddText("윗줄", new RunStyle { Overline = true });
        section.Blocks.Add(p);

        var html = HtmlWriter.ToHtml(doc, fullDocument: false);
        Assert.Contains("text-decoration:overline", html);
    }

    [Fact]
    public void RoundTrip_PreservesOverline()
    {
        var src = "<p><span style=\"text-decoration:overline\">윗줄</span></p>";
        var doc = HtmlReader.FromHtml(src);
        var run = doc.EnumerateParagraphs().Single().Runs.First(r => r.Text.Trim() == "윗줄");
        Assert.True(run.Style.Overline);

        var rendered = HtmlWriter.ToHtml(doc, fullDocument: false);
        var doc2 = HtmlReader.FromHtml(rendered);
        var run2 = doc2.EnumerateParagraphs().Single().Runs.First(r => r.Text.Trim() == "윗줄");
        Assert.True(run2.Style.Overline);
    }

    [Fact]
    public void Writer_ParagraphSpacingAndIndent()
    {
        var doc = new PolyDonkyument();
        var section = new Section();
        doc.Sections.Add(section);
        var p = new Paragraph();
        p.Style.SpaceBeforePt     = 12;
        p.Style.SpaceAfterPt      = 6;
        p.Style.IndentFirstLineMm = 10;
        p.Style.IndentLeftMm      = 5;
        p.AddText("들여쓰기 단락");
        section.Blocks.Add(p);

        var html = HtmlWriter.ToHtml(doc, fullDocument: false);
        Assert.Contains("margin-top:12pt", html);
        Assert.Contains("margin-bottom:6pt", html);
        Assert.Contains("text-indent:10mm", html);
        Assert.Contains("padding-left:5mm", html);
    }

    [Fact]
    public void RoundTrip_PreservesParagraphSpacing()
    {
        var src = "<p style=\"margin-top:12pt;margin-bottom:6pt;text-indent:10mm;padding-left:5mm\">텍스트</p>";
        var doc = HtmlReader.FromHtml(src);
        var p   = doc.EnumerateParagraphs().Single();
        Assert.Equal(12, p.Style.SpaceBeforePt,     precision: 1);
        Assert.Equal(6,  p.Style.SpaceAfterPt,      precision: 1);
        Assert.Equal(10, p.Style.IndentFirstLineMm,  precision: 1);
        Assert.Equal(5,  p.Style.IndentLeftMm,       precision: 1);

        var rendered = HtmlWriter.ToHtml(doc, fullDocument: false);
        var doc2 = HtmlReader.FromHtml(rendered);
        var p2   = doc2.EnumerateParagraphs().Single();
        Assert.Equal(p.Style.SpaceBeforePt,    p2.Style.SpaceBeforePt,    precision: 1);
        Assert.Equal(p.Style.SpaceAfterPt,     p2.Style.SpaceAfterPt,     precision: 1);
        Assert.Equal(p.Style.IndentFirstLineMm, p2.Style.IndentFirstLineMm, precision: 1);
        Assert.Equal(p.Style.IndentLeftMm,     p2.Style.IndentLeftMm,     precision: 1);
    }

    [Fact]
    public void RoundTrip_PreservesLineHeight()
    {
        var src = "<p style=\"line-height:1.8\">줄간격</p>";
        var doc = HtmlReader.FromHtml(src);
        var p   = doc.EnumerateParagraphs().Single();
        Assert.Equal(1.8, p.Style.LineHeightFactor, precision: 2);

        var rendered = HtmlWriter.ToHtml(doc, fullDocument: false);
        var doc2 = HtmlReader.FromHtml(rendered);
        Assert.Equal(1.8, doc2.EnumerateParagraphs().Single().Style.LineHeightFactor, precision: 2);
    }

    [Fact]
    public void Writer_TableColumnWidthsAndCellBackground()
    {
        var table = new PolyDonky.Core.Table();
        table.Columns.Add(new TableColumn { WidthMm = 40 });
        table.Columns.Add(new TableColumn { WidthMm = 60 });
        var row = new TableRow();
        row.Cells.Add(new TableCell
        {
            Blocks          = { Paragraph.Of("셀1") },
            BackgroundColor = "#FFCC00",
        });
        row.Cells.Add(new TableCell { Blocks = { Paragraph.Of("셀2") } });
        table.Rows.Add(row);

        var doc = new PolyDonkyument();
        doc.Sections.Add(new Section());
        doc.Sections[0].Blocks.Add(table);

        var html = HtmlWriter.ToHtml(doc, fullDocument: false);
        Assert.Contains("<colgroup>", html);
        Assert.Contains("width:40mm", html);
        Assert.Contains("width:60mm", html);
        Assert.Contains("background-color:#FFCC00", html);
    }

    [Fact]
    public void RoundTrip_PreservesTableColumnWidthsAndCellBackground()
    {
        var src = "<table><colgroup><col style=\"width:40mm\"><col style=\"width:60mm\"></colgroup>" +
                  "<tbody><tr><td style=\"background-color:#FFCC00\">A</td><td>B</td></tr></tbody></table>";
        var doc = HtmlReader.FromHtml(src);
        var t   = doc.Sections[0].Blocks.OfType<PolyDonky.Core.Table>().Single();

        Assert.Equal(40, t.Columns[0].WidthMm, precision: 0);
        Assert.Equal(60, t.Columns[1].WidthMm, precision: 0);
        Assert.Equal("#FFCC00", t.Rows[0].Cells[0].BackgroundColor, StringComparer.OrdinalIgnoreCase);
    }

    [Fact]
    public void Writer_TableCellBorderAndPadding()
    {
        var table = new PolyDonky.Core.Table();
        table.Columns.Add(new TableColumn { WidthMm = 50 });
        var row   = new TableRow();
        row.Cells.Add(new TableCell
        {
            Blocks            = { Paragraph.Of("X") },
            BorderThicknessPt = 2.0,
            BorderColor       = "#0000FF",
            PaddingTopMm      = 3,
            PaddingLeftMm     = 5,
        });
        table.Rows.Add(row);

        var doc = new PolyDonkyument();
        doc.Sections.Add(new Section());
        doc.Sections[0].Blocks.Add(table);

        var html = HtmlWriter.ToHtml(doc, fullDocument: false);
        Assert.Contains("border-top:2pt solid #0000FF", html);
        Assert.Contains("padding:", html);
    }

    [Fact]
    public void Writer_ImageWrapModeAndHAlign()
    {
        var doc = new PolyDonkyument();
        var section = new Section();
        doc.Sections.Add(section);

        section.Blocks.Add(new ImageBlock
        {
            ResourcePath = "img/a.png",
            WidthMm  = 50,
            HeightMm = 30,
            WrapMode = ImageWrapMode.WrapRight,
        });
        section.Blocks.Add(new ImageBlock
        {
            ResourcePath = "img/b.png",
            WidthMm  = 50,
            HeightMm = 30,
            HAlign   = ImageHAlign.Center,
        });

        var html = HtmlWriter.ToHtml(doc, fullDocument: false);
        Assert.Contains("float:left", html);
        Assert.Contains("display:block;margin-left:auto;margin-right:auto", html);
    }

    [Fact]
    public void RoundTrip_PreservesImageAlignment()
    {
        var src = "<img src=\"img/a.png\" style=\"float:left\">" +
                  "<img src=\"img/b.png\" style=\"display:block;margin-left:auto;margin-right:auto\">";
        var doc    = HtmlReader.FromHtml(src);
        var images = doc.Sections[0].Blocks.OfType<ImageBlock>().ToList();

        Assert.Equal(ImageWrapMode.WrapRight, images[0].WrapMode);
        Assert.Equal(ImageHAlign.Center,      images[1].HAlign);
    }

    [Fact]
    public void Reader_NamedColors()
    {
        var doc  = HtmlReader.FromHtml(
            "<p><span style=\"color:orange\">A</span> " +
            "<span style=\"color:navy\">B</span> " +
            "<span style=\"color:teal\">C</span></p>");
        var runs = doc.EnumerateParagraphs().Single().Runs.Where(r => r.Style.Foreground.HasValue).ToList();
        Assert.Contains(runs, r => r.Style.Foreground!.Value.R == 255 && r.Style.Foreground!.Value.G == 165); // orange
        Assert.Contains(runs, r => r.Style.Foreground!.Value.B == 128 && r.Style.Foreground!.Value.R == 0);   // navy
        Assert.Contains(runs, r => r.Style.Foreground!.Value.R == 0   && r.Style.Foreground!.Value.G == 128 && r.Style.Foreground!.Value.B == 128); // teal
    }

    [Fact]
    public void Reader_TableCellPadding()
    {
        var src = "<table><tr><td style=\"padding:5mm\">X</td></tr></table>";
        var doc = HtmlReader.FromHtml(src);
        var cell = doc.Sections[0].Blocks.OfType<PolyDonky.Core.Table>().Single().Rows[0].Cells[0];
        Assert.Equal(5, cell.PaddingTopMm,    precision: 1);
        Assert.Equal(5, cell.PaddingLeftMm,   precision: 1);
    }

    // ── Table.Caption (<caption> 요소) ─────────────────────────────────

    [Fact]
    public void Reader_TableCaption_Extracted()
    {
        const string html = """
            <table>
              <caption>표 1: 테스트 표</caption>
              <tr><td>A</td></tr>
            </table>
            """;
        var doc = HtmlReader.FromHtml(html);
        var table = doc.Sections[0].Blocks.OfType<PolyDonky.Core.Table>().Single();
        Assert.Equal("표 1: 테스트 표", table.Caption);
    }

    [Fact]
    public void Reader_TableWithoutCaption_CaptionIsNull()
    {
        const string html = "<table><tr><td>A</td></tr></table>";
        var doc = HtmlReader.FromHtml(html);
        var table = doc.Sections[0].Blocks.OfType<PolyDonky.Core.Table>().Single();
        Assert.Null(table.Caption);
    }

    [Fact]
    public void Writer_TableCaption_EmitsElement()
    {
        var pdDoc = new PolyDonkyument();
        var sec = new Section(); pdDoc.Sections.Add(sec);
        var t = new PolyDonky.Core.Table
        {
            Caption = "표 2: 샘플",
        };
        t.Columns.Add(new TableColumn());
        var row = new TableRow(); t.Rows.Add(row);
        row.Cells.Add(new TableCell());
        row.Cells[0].Blocks.Add(Paragraph.Of("X"));
        sec.Blocks.Add(t);

        var html = HtmlWriter.ToHtml(pdDoc, fullDocument: false);
        Assert.Contains("<caption>표 2: 샘플</caption>", html);
    }

    [Fact]
    public void Reader_TableCaption_RoundTrip()
    {
        // HTML → Model → HTML round-trip 시 caption 이 보존됨.
        const string html = "<table><caption>제목</caption><tr><td>A</td></tr></table>";
        var pdDoc = HtmlReader.FromHtml(html);
        var html2  = HtmlWriter.ToHtml(pdDoc, fullDocument: false);
        Assert.Contains("<caption>제목</caption>", html2);
    }

    [Fact]
    public void Writer_ForcePageBreakBefore_EmitsCss()
    {
        var doc = new PolyDonkyument();
        var sec = new Section(); doc.Sections.Add(sec);
        sec.Blocks.Add(Paragraph.Of("A"));
        var p2 = Paragraph.Of("B");
        p2.Style.ForcePageBreakBefore = true;
        sec.Blocks.Add(p2);

        var html = HtmlWriter.ToHtml(doc, fullDocument: false);
        Assert.Contains("page-break-before:always", html);
    }

    [Fact]
    public void Reader_ForcePageBreakBefore_ParsesCss()
    {
        var src = "<p>A</p><p style=\"page-break-before:always\">B</p>";
        var doc = HtmlReader.FromHtml(src);
        var paragraphs = doc.EnumerateParagraphs().ToList();

        Assert.False(paragraphs[0].Style.ForcePageBreakBefore);
        Assert.True(paragraphs[1].Style.ForcePageBreakBefore);
    }

    [Fact]
    public void Reader_ForcePageBreakBefore_ParsesModernCss()
    {
        var src = "<p style=\"break-before:page\">C</p>";
        var doc = HtmlReader.FromHtml(src);
        Assert.True(doc.EnumerateParagraphs().Single().Style.ForcePageBreakBefore);
    }

    [Fact]
    public void RoundTrip_ForcePageBreakBefore_Preserved()
    {
        var doc = new PolyDonkyument();
        var sec = new Section(); doc.Sections.Add(sec);
        sec.Blocks.Add(Paragraph.Of("첫 번째 단락"));
        var p2 = Paragraph.Of("두 번째 단락");
        p2.Style.ForcePageBreakBefore = true;
        sec.Blocks.Add(p2);
        sec.Blocks.Add(Paragraph.Of("세 번째 단락"));

        var html = HtmlWriter.ToHtml(doc, fullDocument: false);
        var rt   = HtmlReader.FromHtml(html);
        var paras = rt.EnumerateParagraphs().ToList();

        Assert.False(paras[0].Style.ForcePageBreakBefore);
        Assert.True(paras[1].Style.ForcePageBreakBefore);
        Assert.False(paras[2].Style.ForcePageBreakBefore);
    }

    [Fact]
    public void Writer_Footnote_EmitsPandocStyleSuperscript()
    {
        var doc = new PolyDonkyument();
        var sec = new Section(); doc.Sections.Add(sec);

        var fn = new FootnoteEntry { Id = "f1" };
        fn.Blocks.Add(Paragraph.Of("각주 내용"));
        doc.Footnotes.Add(fn);

        var p = new Paragraph();
        p.AddText("본문");
        p.Runs.Add(new Run { FootnoteId = "f1" });
        sec.Blocks.Add(p);

        var html = HtmlWriter.ToHtml(doc);
        Assert.Contains("fnref-1", html);
        Assert.Contains("fn-1", html);
        Assert.Contains("각주 내용", html);
        Assert.Contains("section class=\"footnotes\"", html);
    }

    [Fact]
    public void RoundTrip_FootnotesPreserved()
    {
        var doc = new PolyDonkyument();
        var sec = new Section(); doc.Sections.Add(sec);

        var fn = new FootnoteEntry { Id = "f1" };
        fn.Blocks.Add(Paragraph.Of("각주 내용"));
        doc.Footnotes.Add(fn);

        var en = new FootnoteEntry { Id = "e1" };
        en.Blocks.Add(Paragraph.Of("미주 내용"));
        doc.Endnotes.Add(en);

        var p = new Paragraph();
        p.AddText("본문");
        p.Runs.Add(new Run { FootnoteId = "f1" });
        p.AddText(" 텍스트");
        p.Runs.Add(new Run { EndnoteId = "e1" });
        sec.Blocks.Add(p);

        var html = HtmlWriter.ToHtml(doc);
        var rt   = HtmlReader.FromHtml(html);

        Assert.Single(rt.Footnotes);
        Assert.Contains("각주 내용", rt.Footnotes[0].Blocks.OfType<Paragraph>().First().GetPlainText());

        Assert.Single(rt.Endnotes);
        Assert.Contains("미주 내용", rt.Endnotes[0].Blocks.OfType<Paragraph>().First().GetPlainText());

        var runs = rt.Sections[0].Blocks.OfType<Paragraph>().First().Runs;
        Assert.Contains(runs, r => r.FootnoteId is not null);
        Assert.Contains(runs, r => r.EndnoteId is not null);
    }

    // ── SVG → ShapeObject 파서 ─────────────────────────────────────

    [Fact]
    public void RoundTrip_ShapeRectangle()
    {
        var doc = new PolyDonkyument();
        var sec = new Section(); doc.Sections.Add(sec);
        sec.Blocks.Add(new ShapeObject
        {
            Kind = ShapeKind.Rectangle, WidthMm = 40, HeightMm = 30,
            StrokeColor = "#FF0000", StrokeThicknessPt = 2, FillColor = "#AABBCC",
        });
        var html = HtmlWriter.ToHtml(doc);
        var rt   = HtmlReader.FromHtml(html);
        var shape = rt.Sections[0].Blocks.OfType<ShapeObject>().First();
        Assert.Equal(ShapeKind.Rectangle, shape.Kind);
        Assert.InRange(shape.WidthMm,  39.5, 40.5);
        Assert.InRange(shape.HeightMm, 29.5, 30.5);
        Assert.Equal("#FF0000", shape.StrokeColor);
        Assert.Equal("#AABBCC", shape.FillColor);
    }

    [Fact]
    public void RoundTrip_ShapeRoundedRect()
    {
        var doc = new PolyDonkyument();
        var sec = new Section(); doc.Sections.Add(sec);
        sec.Blocks.Add(new ShapeObject { Kind = ShapeKind.RoundedRect, WidthMm = 50, HeightMm = 20, CornerRadiusMm = 5 });
        var html = HtmlWriter.ToHtml(doc);
        var rt   = HtmlReader.FromHtml(html);
        var shape = rt.Sections[0].Blocks.OfType<ShapeObject>().First();
        Assert.Equal(ShapeKind.RoundedRect, shape.Kind);
        Assert.InRange(shape.CornerRadiusMm, 4.5, 5.5);
    }

    [Fact]
    public void RoundTrip_ShapeEllipse()
    {
        var doc = new PolyDonkyument();
        var sec = new Section(); doc.Sections.Add(sec);
        sec.Blocks.Add(new ShapeObject { Kind = ShapeKind.Ellipse, WidthMm = 60, HeightMm = 40 });
        var html = HtmlWriter.ToHtml(doc);
        var rt   = HtmlReader.FromHtml(html);
        var shape = rt.Sections[0].Blocks.OfType<ShapeObject>().First();
        Assert.Equal(ShapeKind.Ellipse, shape.Kind);
    }

    [Fact]
    public void RoundTrip_ShapeLine()
    {
        var doc = new PolyDonkyument();
        var sec = new Section(); doc.Sections.Add(sec);
        var ln = new ShapeObject { Kind = ShapeKind.Line, WidthMm = 50, HeightMm = 30 };
        ln.Points.Add(new ShapePoint { X = 0, Y = 0 });
        ln.Points.Add(new ShapePoint { X = 50, Y = 30 });
        sec.Blocks.Add(ln);
        var html = HtmlWriter.ToHtml(doc);
        var rt   = HtmlReader.FromHtml(html);
        var shape = rt.Sections[0].Blocks.OfType<ShapeObject>().First();
        Assert.Equal(ShapeKind.Line, shape.Kind);
        Assert.Equal(2, shape.Points.Count);
    }

    [Fact]
    public void RoundTrip_ShapePolyline()
    {
        var doc = new PolyDonkyument();
        var sec = new Section(); doc.Sections.Add(sec);
        var poly = new ShapeObject { Kind = ShapeKind.Polyline, WidthMm = 60, HeightMm = 40 };
        poly.Points.Add(new ShapePoint { X = 0,  Y = 40 });
        poly.Points.Add(new ShapePoint { X = 30, Y = 0  });
        poly.Points.Add(new ShapePoint { X = 60, Y = 40 });
        sec.Blocks.Add(poly);
        var html = HtmlWriter.ToHtml(doc);
        var rt   = HtmlReader.FromHtml(html);
        var shape = rt.Sections[0].Blocks.OfType<ShapeObject>().First();
        Assert.Equal(ShapeKind.Polyline, shape.Kind);
        Assert.Equal(3, shape.Points.Count);
    }

    [Fact]
    public void RoundTrip_ShapePolygon()
    {
        var doc = new PolyDonkyument();
        var sec = new Section(); doc.Sections.Add(sec);
        var poly = new ShapeObject { Kind = ShapeKind.Polygon, WidthMm = 50, HeightMm = 50 };
        for (int i = 0; i < 5; i++)
        {
            double a = 2 * Math.PI * i / 5 - Math.PI / 2;
            poly.Points.Add(new ShapePoint { X = 25 + 25 * Math.Cos(a), Y = 25 + 25 * Math.Sin(a) });
        }
        sec.Blocks.Add(poly);
        var html = HtmlWriter.ToHtml(doc);
        var rt   = HtmlReader.FromHtml(html);
        var shape = rt.Sections[0].Blocks.OfType<ShapeObject>().First();
        Assert.Equal(ShapeKind.Polygon, shape.Kind);
        Assert.Equal(5, shape.Points.Count);
    }

    [Fact]
    public void RoundTrip_ShapeTriangle()
    {
        var doc = new PolyDonkyument();
        var sec = new Section(); doc.Sections.Add(sec);
        var tri = new ShapeObject { Kind = ShapeKind.Triangle, WidthMm = 40, HeightMm = 35 };
        tri.Points.Add(new ShapePoint { X = 20, Y = 0  });
        tri.Points.Add(new ShapePoint { X = 40, Y = 35 });
        tri.Points.Add(new ShapePoint { X = 0,  Y = 35 });
        sec.Blocks.Add(tri);
        var html = HtmlWriter.ToHtml(doc);
        var rt   = HtmlReader.FromHtml(html);
        var shape = rt.Sections[0].Blocks.OfType<ShapeObject>().First();
        Assert.Equal(ShapeKind.Triangle, shape.Kind);
        Assert.Equal(3, shape.Points.Count);
    }

    [Fact]
    public void RoundTrip_ShapeSpline()
    {
        var doc = new PolyDonkyument();
        var sec = new Section(); doc.Sections.Add(sec);
        var sp = new ShapeObject { Kind = ShapeKind.Spline, WidthMm = 60, HeightMm = 40 };
        sp.Points.Add(new ShapePoint { X = 0,  Y = 20 });
        sp.Points.Add(new ShapePoint { X = 30, Y = 0  });
        sp.Points.Add(new ShapePoint { X = 60, Y = 20 });
        sec.Blocks.Add(sp);
        var html = HtmlWriter.ToHtml(doc);
        var rt   = HtmlReader.FromHtml(html);
        var shape = rt.Sections[0].Blocks.OfType<ShapeObject>().First();
        Assert.Equal(ShapeKind.Spline, shape.Kind);
        Assert.Equal(3, shape.Points.Count);
        // 제어점이 복원돼야 함.
        Assert.True(shape.Points[0].OutCtrlX.HasValue || shape.Points[1].InCtrlX.HasValue);
    }

    [Fact]
    public void RoundTrip_ShapeClosedSpline()
    {
        var doc = new PolyDonkyument();
        var sec = new Section(); doc.Sections.Add(sec);
        var sp = new ShapeObject { Kind = ShapeKind.ClosedSpline, WidthMm = 50, HeightMm = 50 };
        sp.Points.Add(new ShapePoint { X = 25, Y = 0  });
        sp.Points.Add(new ShapePoint { X = 50, Y = 50 });
        sp.Points.Add(new ShapePoint { X = 0,  Y = 50 });
        sec.Blocks.Add(sp);
        var html = HtmlWriter.ToHtml(doc);
        var rt   = HtmlReader.FromHtml(html);
        var shape = rt.Sections[0].Blocks.OfType<ShapeObject>().First();
        Assert.Equal(ShapeKind.ClosedSpline, shape.Kind);
    }

    [Fact]
    public void RoundTrip_ShapeWithLabel()
    {
        var doc = new PolyDonkyument();
        var sec = new Section(); doc.Sections.Add(sec);
        sec.Blocks.Add(new ShapeObject
        {
            Kind = ShapeKind.Ellipse, WidthMm = 40, HeightMm = 30,
            LabelText = "타원 레이블",
        });
        var html = HtmlWriter.ToHtml(doc);
        var rt   = HtmlReader.FromHtml(html);
        var shape = rt.Sections[0].Blocks.OfType<ShapeObject>().First();
        Assert.Equal(ShapeKind.Ellipse, shape.Kind);
        Assert.Equal("타원 레이블", shape.LabelText);
    }

    // ── 편집용지 설정 ──────────────────────────────────────────────────

    [Fact]
    public void Writer_EmitsPageMetaAndAtPage_DefaultA4()
    {
        var doc = new PolyDonkyument();
        doc.Sections.Add(new Section()); // 기본값: A4 세로
        var html = HtmlWriter.ToHtml(doc);
        Assert.Contains("pd-page-size\" content=\"A4\"", html);
        Assert.Contains("pd-page-orientation\" content=\"portrait\"", html);
        Assert.Contains("@page", html);
        Assert.Contains("210", html);  // A4 너비 210mm
        Assert.Contains("margin:", html);
    }

    [Fact]
    public void Writer_EmitsLandscapeOrientation()
    {
        var doc = new PolyDonkyument();
        var sec = new Section();
        sec.Page.ApplySizeKind(PaperSizeKind.A4);
        sec.Page.Orientation = PageOrientation.Landscape;
        doc.Sections.Add(sec);
        var html = HtmlWriter.ToHtml(doc);
        Assert.Contains("pd-page-orientation\" content=\"landscape\"", html);
        // 가로 방향이면 SVG 크기 297mm × 210mm 순서로 출력.
        Assert.Contains("297", html);
    }

    [Fact]
    public void Writer_EmitsCustomPageSize()
    {
        var doc = new PolyDonkyument();
        var sec = new Section();
        sec.Page.SizeKind = PaperSizeKind.Custom;
        sec.Page.WidthMm  = 180;
        sec.Page.HeightMm = 240;
        doc.Sections.Add(sec);
        var html = HtmlWriter.ToHtml(doc);
        Assert.Contains("pd-page-size\" content=\"Custom\"", html);
        Assert.Contains("pd-page-width\" content=\"180mm\"", html);
        Assert.Contains("pd-page-height\" content=\"240mm\"", html);
    }

    [Fact]
    public void RoundTrip_PageSettings_A4Portrait()
    {
        var doc = new PolyDonkyument();
        var sec = new Section();
        sec.Page.ApplySizeKind(PaperSizeKind.A4);
        sec.Page.Orientation    = PageOrientation.Portrait;
        sec.Page.MarginTopMm    = 30;
        sec.Page.MarginBottomMm = 25;
        sec.Page.MarginLeftMm   = 35;
        sec.Page.MarginRightMm  = 20;
        doc.Sections.Add(sec);

        var html = HtmlWriter.ToHtml(doc);
        var rt   = HtmlReader.FromHtml(html);

        var page = rt.Sections[0].Page;
        Assert.Equal(PaperSizeKind.A4,            page.SizeKind);
        Assert.Equal(PageOrientation.Portrait,     page.Orientation);
        Assert.InRange(page.MarginTopMm,    29.5, 30.5);
        Assert.InRange(page.MarginBottomMm, 24.5, 25.5);
        Assert.InRange(page.MarginLeftMm,   34.5, 35.5);
        Assert.InRange(page.MarginRightMm,  19.5, 20.5);
    }

    [Fact]
    public void RoundTrip_PageSettings_A4Landscape()
    {
        var doc = new PolyDonkyument();
        var sec = new Section();
        sec.Page.ApplySizeKind(PaperSizeKind.A4);
        sec.Page.Orientation = PageOrientation.Landscape;
        doc.Sections.Add(sec);
        var html = HtmlWriter.ToHtml(doc);
        var rt   = HtmlReader.FromHtml(html);
        Assert.Equal(PaperSizeKind.A4,          rt.Sections[0].Page.SizeKind);
        Assert.Equal(PageOrientation.Landscape,  rt.Sections[0].Page.Orientation);
    }

    [Fact]
    public void RoundTrip_PageSettings_CustomSize()
    {
        var doc = new PolyDonkyument();
        var sec = new Section();
        sec.Page.SizeKind = PaperSizeKind.Custom;
        sec.Page.WidthMm  = 170;
        sec.Page.HeightMm = 235;
        doc.Sections.Add(sec);
        var html = HtmlWriter.ToHtml(doc);
        var rt   = HtmlReader.FromHtml(html);
        var page = rt.Sections[0].Page;
        Assert.Equal(PaperSizeKind.Custom, page.SizeKind);
        Assert.InRange(page.WidthMm,  169, 171);
        Assert.InRange(page.HeightMm, 234, 236);
    }

    [Fact]
    public void Reader_NoPageMeta_DefaultsToA4Portrait()
    {
        // 페이지 정보가 없는 단순 HTML → A4 세로 기본 여백이어야 한다.
        const string html = "<p>테스트</p>";
        var rt   = HtmlReader.FromHtml(html);
        var page = rt.Sections[0].Page;
        Assert.Equal(PaperSizeKind.A4,        page.SizeKind);
        Assert.Equal(PageOrientation.Portrait, page.Orientation);
        Assert.InRange(page.WidthMm,  209, 211);
        Assert.InRange(page.HeightMm, 296, 298);
    }

    [Fact]
    public void Reader_ExternalAtPage_ParsedCorrectly()
    {
        // 외부 HTML 의 @page CSS 를 읽어 페이지 설정을 복원해야 한다.
        const string html = """
            <!DOCTYPE html>
            <html><head>
              <style>
                @page {
                  size: 215.9mm 279.4mm;
                  margin: 25mm 20mm 25mm 30mm;
                }
              </style>
            </head><body><p>내용</p></body></html>
            """;
        var rt   = HtmlReader.FromHtml(html);
        var page = rt.Sections[0].Page;
        Assert.Equal(PaperSizeKind.Letter, page.SizeKind);
        Assert.InRange(page.MarginTopMm,   24.5, 25.5);
        Assert.InRange(page.MarginRightMm, 19.5, 20.5);
        Assert.InRange(page.MarginLeftMm,  29.5, 30.5);
    }

    [Fact]
    public void Reader_StandaloneSvgParsedAsShapeObject()
    {
        // <svg> 가 <figure> 없이 직접 나타나도 ShapeObject 로 파싱돼야 함.
        const string html = "<svg width=\"100\" height=\"75\"><rect x=\"0.5\" y=\"0.5\" width=\"99\" height=\"74\" stroke=\"#000\" stroke-width=\"1\" fill=\"none\"></rect></svg>";
        var rt = HtmlReader.FromHtml(html);
        Assert.Contains(rt.Sections[0].Blocks, b => b is ShapeObject { Kind: ShapeKind.Rectangle });
    }

    // ── 복합 SVG → ImageBlock ──────────────────────────────────────────

    [Fact]
    public void Reader_MultiShapeSvg_BecomesImageBlock()
    {
        // 다중 도형 SVG 는 ShapeObject 가 아닌 ImageBlock (image/svg+xml) 으로 보존돼야 한다.
        const string html = @"<svg width=""600"" height=""160"">
            <rect x=""20"" y=""30"" width=""100"" height=""80"" fill=""#4A90E2""></rect>
            <circle cx=""310"" cy=""70"" r=""40"" fill=""#F5A623""></circle>
        </svg>";
        var rt = HtmlReader.FromHtml(html);
        var block = Assert.Single(rt.Sections[0].Blocks);
        var img = Assert.IsType<ImageBlock>(block);
        Assert.Equal("image/svg+xml", img.MediaType);
        Assert.True(img.Data.Length > 0);
        Assert.InRange(img.WidthMm,  155, 160);
        Assert.InRange(img.HeightMm,  41,  43);
    }

    [Fact]
    public void Reader_SvgWithTextLabel_BecomesImageBlock()
    {
        // <text> 레이블이 있는 SVG 는 ImageBlock 으로 보존돼야 한다.
        const string html = @"<svg width=""200"" height=""100"">
            <rect x=""10"" y=""10"" width=""80"" height=""60"" fill=""blue""></rect>
            <text x=""50"" y=""50"">레이블</text>
        </svg>";
        var rt    = HtmlReader.FromHtml(html);
        var block = Assert.Single(rt.Sections[0].Blocks);
        Assert.IsType<ImageBlock>(block);
    }

    [Fact]
    public void Reader_FigureWithSvg_BecomesImageBlockWithCaption()
    {
        // <figure><svg>...</svg><figcaption>캡션</figcaption></figure> → ImageBlock(Title=캡션).
        const string html = @"<figure>
            <svg width=""600"" height=""320"">
                <ellipse cx=""300"" cy=""35"" rx=""60"" ry=""20"" fill=""#A8E6CF""></ellipse>
                <text x=""300"" y=""40"">시작</text>
                <rect x=""230"" y=""150"" width=""140"" height=""40"" fill=""#FFAAA5""></rect>
            </svg>
            <figcaption>그림 5: 플로우차트</figcaption>
        </figure>";
        var rt    = HtmlReader.FromHtml(html);
        var block = Assert.Single(rt.Sections[0].Blocks);
        var img   = Assert.IsType<ImageBlock>(block);
        Assert.Equal("image/svg+xml", img.MediaType);
        Assert.Equal("그림 5: 플로우차트", img.Title);
        Assert.True(img.ShowTitle);
    }

    [Fact]
    public void RoundTrip_SvgImageBlock_PreservesSvgContent()
    {
        // ImageBlock(image/svg+xml) → HtmlWriter → HtmlReader → 동일한 ImageBlock.
        var doc   = new PolyDonkyument();
        var sec   = new Section();
        doc.Sections.Add(sec);
        var svgData = "<svg width=\"200\" height=\"100\"><rect x=\"10\" y=\"10\" width=\"80\" height=\"60\" fill=\"blue\"/><circle cx=\"150\" cy=\"50\" r=\"30\" fill=\"red\"/></svg>";
        sec.Blocks.Add(new ImageBlock
        {
            MediaType = "image/svg+xml",
            Data      = System.Text.Encoding.UTF8.GetBytes(svgData),
            WidthMm   = 200 * 25.4 / 96.0,
            HeightMm  = 100 * 25.4 / 96.0,
        });

        var html = HtmlWriter.ToHtml(doc);
        Assert.Contains("<svg", html);

        var rt    = HtmlReader.FromHtml(html);
        var block = Assert.Single(rt.Sections[0].Blocks);
        var img   = Assert.IsType<ImageBlock>(block);
        Assert.Equal("image/svg+xml", img.MediaType);
        Assert.True(img.Data.Length > 0);
    }

    // ── CSS 도형 → ShapeObject ─────────────────────────────────────────

    [Fact]
    public void Reader_CssRectangle_BecomesShapeObject()
    {
        // 텍스트 없는 색상 div → Rectangle ShapeObject.
        const string html = "<div style=\"width: 80px; height: 80px; background: #4A90E2\"></div>";
        var rt    = HtmlReader.FromHtml(html);
        var block = rt.Sections[0].Blocks.OfType<ShapeObject>().FirstOrDefault();
        Assert.NotNull(block);
        Assert.Equal(ShapeKind.Rectangle, block.Kind);
        Assert.Equal("#4A90E2", block.FillColor);
        Assert.InRange(block.WidthMm,  20, 22);
        Assert.InRange(block.HeightMm, 20, 22);
    }

    [Fact]
    public void Reader_CssEllipse_BecomesShapeObject()
    {
        // border-radius:50% → Ellipse ShapeObject.
        const string html = "<div style=\"width: 80px; height: 80px; background: #7ED321; border-radius: 50%\"></div>";
        var rt    = HtmlReader.FromHtml(html);
        var block = rt.Sections[0].Blocks.OfType<ShapeObject>().FirstOrDefault();
        Assert.NotNull(block);
        Assert.Equal(ShapeKind.Ellipse, block.Kind);
        Assert.Equal("#7ED321", block.FillColor);
    }

    [Fact]
    public void Reader_CssRoundedRect_BecomesShapeObject()
    {
        // border-radius:10px → RoundedRect ShapeObject.
        const string html = "<div style=\"width: 100px; height: 50px; background: #D0021B; border-radius: 10px\"></div>";
        var rt    = HtmlReader.FromHtml(html);
        var block = rt.Sections[0].Blocks.OfType<ShapeObject>().FirstOrDefault();
        Assert.NotNull(block);
        Assert.Equal(ShapeKind.RoundedRect, block.Kind);
        Assert.InRange(block.CornerRadiusMm, 2.5, 2.8);
    }

    [Fact]
    public void Reader_CssBorderTrickTriangle_BecomesShapeObject()
    {
        // CSS border-trick 삼각형 → Triangle ShapeObject.
        const string html = "<div style=\"width: 0; height: 0; border-left: 40px solid transparent; border-right: 40px solid transparent; border-bottom: 70px solid #F5A623\"></div>";
        var rt    = HtmlReader.FromHtml(html);
        var block = rt.Sections[0].Blocks.OfType<ShapeObject>().FirstOrDefault();
        Assert.NotNull(block);
        Assert.Equal(ShapeKind.Triangle, block.Kind);
        Assert.Equal("#F5A623", block.FillColor);
        Assert.InRange(block.WidthMm,  20, 22);   // 80px = 21.2mm
        Assert.InRange(block.HeightMm, 18, 20);   // 70px = 18.5mm
    }

    [Fact]
    public void Reader_CssRotatedRect_BecomesShapeObject()
    {
        // transform:rotate(45deg) → Rectangle with RotationAngleDeg=45.
        const string html = "<div style=\"width: 80px; height: 80px; background: #BD10E0; transform: rotate(45deg)\"></div>";
        var rt    = HtmlReader.FromHtml(html);
        var block = rt.Sections[0].Blocks.OfType<ShapeObject>().FirstOrDefault();
        Assert.NotNull(block);
        Assert.Equal(ShapeKind.Rectangle, block.Kind);
        Assert.Equal(45, block.RotationAngleDeg);
        Assert.Equal("#BD10E0", block.FillColor);
    }

    [Fact]
    public void Reader_DivWithChildrenNotTreatedAsCssShape()
    {
        // 자식이 있는 div 는 CSS 도형으로 감지하지 않아야 함.
        const string html = "<div style=\"width: 80px; height: 80px; background: #ccc\"><span>텍스트</span></div>";
        var rt = HtmlReader.FromHtml(html);
        Assert.DoesNotContain(rt.Sections[0].Blocks, b => b is ShapeObject);
    }

    // ── <style> 블록 CSS 클래스 규칙 머지 ──────────────────────────────

    [Fact]
    public void Reader_CssClassRule_AppliedToParagraph()
    {
        // <style> 의 .center 클래스 → text-align:center 가 매칭 단락에 적용돼야 한다.
        const string html = """
            <html><head><style>
              .center { text-align: center; }
            </style></head><body>
              <p class="center">중앙</p>
            </body></html>
            """;
        var rt = HtmlReader.FromHtml(html);
        var p  = rt.EnumerateParagraphs().Single(x => x.Runs.Any(r => r.Text == "중앙"));
        Assert.Equal(Alignment.Center, p.Style.Alignment);
    }

    [Fact]
    public void Reader_CssTagRule_AppliedToHeading()
    {
        // h1 { text-align: center; } → 모든 h1 이 중앙 정렬.
        const string html = """
            <html><head><style>
              h1 { text-align: center; }
            </style></head><body>
              <h1>제목</h1>
            </body></html>
            """;
        var rt = HtmlReader.FromHtml(html);
        var p  = rt.EnumerateParagraphs().Single();
        Assert.Equal(Alignment.Center, p.Style.Alignment);
    }

    [Fact]
    public void Reader_CssIdRule_Applied()
    {
        const string html = """
            <html><head><style>
              #title { text-align: right; }
            </style></head><body>
              <p id="title">우측</p>
            </body></html>
            """;
        var rt = HtmlReader.FromHtml(html);
        var p  = rt.EnumerateParagraphs().Single();
        Assert.Equal(Alignment.Right, p.Style.Alignment);
    }

    [Fact]
    public void Reader_InlineStyle_OverridesClassRule()
    {
        // 인라인 style 이 클래스 규칙보다 우선해야 한다.
        const string html = """
            <html><head><style>
              .x { text-align: left; }
            </style></head><body>
              <p class="x" style="text-align: right;">우측</p>
            </body></html>
            """;
        var rt = HtmlReader.FromHtml(html);
        var p  = rt.EnumerateParagraphs().Single();
        Assert.Equal(Alignment.Right, p.Style.Alignment);
    }

    [Fact]
    public void Reader_CssDescendantSelector_OnlyAppliesToActualDescendants()
    {
        // ".container p" 자손 셀렉터는 .container 안의 <p> 에만 적용된다 (CSS 표준).
        // 이전 구현은 우측 단순 셀렉터만 보고 모든 <p> 에 잘못 전파했었다.
        const string html = """
            <html><head><style>
              .container p { text-align: right; }
            </style></head><body>
              <p>outside</p>
              <div class="container"><p>inside</p></div>
            </body></html>
            """;
        var rt = HtmlReader.FromHtml(html);
        var ps = rt.EnumerateParagraphs().ToList();
        var outside = ps.First(p => p.GetPlainText() == "outside");
        var inside  = ps.First(p => p.GetPlainText() == "inside");
        Assert.NotEqual(Alignment.Right, outside.Style.Alignment);
        Assert.Equal   (Alignment.Right, inside.Style.Alignment);
    }

    // ── list-style-type 파싱 ────────────────────────────────────────────

    [Fact]
    public void Reader_OlTypeA_BecomesOrderedAlpha()
    {
        // <ol type="A"> → OrderedAlpha
        const string html = "<ol type=\"A\"><li>a</li><li>b</li></ol>";
        var rt = HtmlReader.FromHtml(html);
        var markers = rt.EnumerateParagraphs()
            .Select(p => p.Style.ListMarker).Where(m => m is not null).ToList();
        Assert.All(markers, m => Assert.Equal(ListKind.OrderedAlpha, m!.Kind));
    }

    [Fact]
    public void Reader_OlStyleUpperRoman_BecomesOrderedRoman()
    {
        // <ol style="list-style-type: upper-roman"> → OrderedRoman
        const string html = "<ol style=\"list-style-type: upper-roman\"><li>i</li><li>ii</li></ol>";
        var rt = HtmlReader.FromHtml(html);
        var markers = rt.EnumerateParagraphs()
            .Select(p => p.Style.ListMarker).Where(m => m is not null).ToList();
        Assert.All(markers, m => Assert.Equal(ListKind.OrderedRoman, m!.Kind));
    }

    [Fact]
    public void Reader_OlStyleLowerAlpha_BecomesOrderedAlpha()
    {
        // <ol style="list-style-type: lower-alpha"> → OrderedAlpha
        const string html = "<ol style=\"list-style-type: lower-alpha\"><li>a</li></ol>";
        var rt = HtmlReader.FromHtml(html);
        var marker = rt.EnumerateParagraphs().Single().Style.ListMarker;
        Assert.NotNull(marker);
        Assert.Equal(ListKind.OrderedAlpha, marker.Kind);
    }

    [Fact]
    public void Reader_OlDefault_BecomesOrderedDecimal()
    {
        // <ol> 기본값 → OrderedDecimal
        const string html = "<ol><li>1</li></ol>";
        var rt = HtmlReader.FromHtml(html);
        var marker = rt.EnumerateParagraphs().Single().Style.ListMarker;
        Assert.NotNull(marker);
        Assert.Equal(ListKind.OrderedDecimal, marker.Kind);
    }

    [Fact]
    public void Reader_LiTypeOverridesOlType()
    {
        // <li type="i"> 개별 항목 지정이 부모 <ol> 을 오버라이드함.
        const string html = "<ol type=\"A\"><li type=\"i\">first</li><li>second</li></ol>";
        var rt = HtmlReader.FromHtml(html);
        var markers = rt.EnumerateParagraphs()
            .Select(p => p.Style.ListMarker).Where(m => m is not null).ToList();
        Assert.Equal(2, markers.Count);
        Assert.Equal(ListKind.OrderedRoman,   markers[0]!.Kind);
        Assert.Equal(ListKind.OrderedAlpha,   markers[1]!.Kind);
    }

    // ── ListMarker.UpperCase — type 속성으로 대소문자 보존 ────────────

    [Fact]
    public void Reader_OlTypeUpperA_PreservesUpperCase()
    {
        // <ol type="A"> 가 중첩 레벨에 와도 대문자 알파벳 유지.
        const string html = "<ol><li>top<ol type=\"A\"><li>nested</li></ol></li></ol>";
        var rt = HtmlReader.FromHtml(html);
        var nested = rt.EnumerateParagraphs()
            .Select(p => p.Style.ListMarker)
            .FirstOrDefault(m => m?.Kind == ListKind.OrderedAlpha);
        Assert.NotNull(nested);
        Assert.True(nested.UpperCase);
    }

    [Fact]
    public void Reader_OlTypeLowerA_PreservesLowerCase()
    {
        const string html = "<ol type=\"a\"><li>x</li></ol>";
        var rt = HtmlReader.FromHtml(html);
        var marker = rt.EnumerateParagraphs().Single().Style.ListMarker;
        Assert.NotNull(marker);
        Assert.Equal(ListKind.OrderedAlpha, marker.Kind);
        Assert.False(marker.UpperCase);
    }

    [Fact]
    public void Reader_OlTypeLowerI_PreservesLowerCase()
    {
        // <ol type="i"> 가 중첩 레벨에 와도 소문자 로마자로 명시 보존.
        const string html = "<ol><li>top<ol type=\"i\"><li>nested</li></ol></li></ol>";
        var rt = HtmlReader.FromHtml(html);
        var nested = rt.EnumerateParagraphs()
            .Select(p => p.Style.ListMarker)
            .FirstOrDefault(m => m?.Kind == ListKind.OrderedRoman);
        Assert.NotNull(nested);
        Assert.False(nested.UpperCase);
    }

    [Fact]
    public void Reader_OlStyleUpperLatin_PreservesUpperCase()
    {
        const string html = "<ol style=\"list-style-type: upper-latin\"><li>x</li></ol>";
        var rt = HtmlReader.FromHtml(html);
        var marker = rt.EnumerateParagraphs().Single().Style.ListMarker;
        Assert.NotNull(marker);
        Assert.True(marker.UpperCase);
    }

    [Fact]
    public void Reader_OlDefault_UpperCaseIsNull()
    {
        // 일반 <ol> (decimal) 은 대소문자 정보가 없음.
        const string html = "<ol><li>x</li></ol>";
        var rt = HtmlReader.FromHtml(html);
        var marker = rt.EnumerateParagraphs().Single().Style.ListMarker;
        Assert.NotNull(marker);
        Assert.Null(marker.UpperCase);
    }

    // ── CSS-only 체크리스트 (class="checklist", li class="checked") ────

    [Fact]
    public void Reader_ChecklistClass_DetectsCheckedState()
    {
        // <ul class="checklist"><li class="checked">…</li><li>…</li></ul>
        // 첫째 항목 = Checked=true, 둘째 = Checked=false.
        const string html = """
            <ul class="checklist">
              <li class="checked">완료 항목</li>
              <li>대기 항목</li>
            </ul>
            """;
        var rt = HtmlReader.FromHtml(html);
        var markers = rt.EnumerateParagraphs()
            .Select(p => p.Style.ListMarker).Where(m => m is not null).ToList();
        Assert.Equal(2, markers.Count);
        Assert.True(markers[0]!.Checked);
        Assert.False(markers[1]!.Checked);
    }

    [Fact]
    public void Reader_ChecklistClass_AlsoDetectsTaskListClass()
    {
        // GitHub 의 task-list / contains-task-list 클래스도 동일하게 인식.
        const string html = """
            <ul class="task-list">
              <li class="checked">a</li>
              <li>b</li>
            </ul>
            """;
        var rt = HtmlReader.FromHtml(html);
        var markers = rt.EnumerateParagraphs()
            .Select(p => p.Style.ListMarker).Where(m => m is not null).ToList();
        Assert.Equal(2, markers.Count);
        Assert.True(markers[0]!.Checked);
        Assert.False(markers[1]!.Checked);
    }

    [Fact]
    public void Reader_NormalUl_DoesNotSetCheckedState()
    {
        // 평범한 <ul> 항목은 Checked 가 null 이어야 한다.
        const string html = "<ul><li>x</li><li>y</li></ul>";
        var rt = HtmlReader.FromHtml(html);
        var markers = rt.EnumerateParagraphs()
            .Select(p => p.Style.ListMarker).Where(m => m is not null).ToList();
        Assert.Equal(2, markers.Count);
        Assert.Null(markers[0]!.Checked);
        Assert.Null(markers[1]!.Checked);
    }

    [Fact]
    public void ListMarker_Clone_PreservesUpperCaseAndChecked()
    {
        // ListMarker.Clone() 이 새 필드를 모두 복사하는지.
        var lm = new ListMarker
        {
            Kind          = ListKind.OrderedAlpha,
            Level         = 2,
            OrderedNumber = 5,
            Checked       = true,
            UpperCase     = true,
        };
        var clone = lm.Clone();
        Assert.NotSame(lm, clone);
        Assert.Equal(lm.Kind,          clone.Kind);
        Assert.Equal(lm.Level,         clone.Level);
        Assert.Equal(lm.OrderedNumber, clone.OrderedNumber);
        Assert.Equal(lm.Checked,       clone.Checked);
        Assert.Equal(lm.UpperCase,     clone.UpperCase);
    }

    // ── CSS Grid/Flex → Table 변환 ──────────────────────────────────

    [Fact]
    public void Reader_CssGrid2Col_BecomesTable()
    {
        // display:grid; grid-template-columns: 1fr 1fr → 2-column Table
        const string html = """
            <div style="display:grid;grid-template-columns:1fr 1fr">
              <div><p>Left</p></div>
              <div><p>Right</p></div>
            </div>
            """;
        var rt    = HtmlReader.FromHtml(html);
        var table = rt.Sections[0].Blocks.OfType<Table>().FirstOrDefault();
        Assert.NotNull(table);
        Assert.Equal(2, table.Columns.Count);
        var row = Assert.Single(table.Rows);
        Assert.Equal(2, row.Cells.Count);
        Assert.Contains(row.Cells[0].Blocks.OfType<Paragraph>(), p => p.GetPlainText() == "Left");
        Assert.Contains(row.Cells[1].Blocks.OfType<Paragraph>(), p => p.GetPlainText() == "Right");
    }

    [Fact]
    public void Reader_CssGrid3Col_BecomesTable()
    {
        // grid-template-columns: repeat(3, 1fr) → 3-column Table
        const string html = """
            <div style="display:grid;grid-template-columns:repeat(3,1fr)">
              <div><p>A</p></div><div><p>B</p></div><div><p>C</p></div>
            </div>
            """;
        var rt    = HtmlReader.FromHtml(html);
        var table = rt.Sections[0].Blocks.OfType<Table>().FirstOrDefault();
        Assert.NotNull(table);
        Assert.Equal(3, table.Columns.Count);
    }

    [Fact]
    public void Reader_CssFlex2Col_BecomesTable()
    {
        // display:flex with 2 child divs → 2-column Table
        const string html = """
            <div style="display:flex">
              <div><p>Col A</p></div>
              <div><p>Col B</p></div>
            </div>
            """;
        var rt    = HtmlReader.FromHtml(html);
        var table = rt.Sections[0].Blocks.OfType<Table>().FirstOrDefault();
        Assert.NotNull(table);
        Assert.Equal(2, table.Columns.Count);
    }

    [Fact]
    public void Reader_CssFlexColumnDirection_NotConvertedToTable()
    {
        // flex-direction:column → 세로 배치이므로 Table 로 변환하지 않아야 함.
        const string html = """
            <div style="display:flex;flex-direction:column">
              <div><p>A</p></div>
              <div><p>B</p></div>
            </div>
            """;
        var rt    = HtmlReader.FromHtml(html);
        var table = rt.Sections[0].Blocks.OfType<Table>().FirstOrDefault();
        Assert.Null(table);
    }

    // ── CSS 상속 (text-align inheritance) ────────────────────────────

    [Fact]
    public void Reader_TextAlignInheritedFromParent()
    {
        // 부모의 text-align 이 자체 값이 없는 자식에게 전파.
        const string html = """
            <html><head><style>
              .center { text-align: center; }
            </style></head><body>
              <div class="center">
                <h1>Title</h1>
                <p>Subtitle</p>
              </div>
            </body></html>
            """;
        var rt = HtmlReader.FromHtml(html);
        var ps = rt.EnumerateParagraphs().ToList();
        Assert.Equal(2, ps.Count);
        Assert.Equal(Alignment.Center, ps[0].Style.Alignment); // h1
        Assert.Equal(Alignment.Center, ps[1].Style.Alignment); // p
    }

    [Fact]
    public void Reader_TextAlignChildOverridesInherited()
    {
        // 자식의 자체 text-align 이 부모로부터 상속받는 값보다 우선.
        const string html = """
            <html><head><style>
              .center { text-align: center; }
              .right  { text-align: right; }
            </style></head><body>
              <div class="center">
                <h1>Centered</h1>
                <div class="right"><p>Right aligned</p></div>
              </div>
            </body></html>
            """;
        var rt = HtmlReader.FromHtml(html);
        var ps = rt.EnumerateParagraphs().ToList();
        Assert.Equal(Alignment.Center, ps[0].Style.Alignment);
        Assert.Equal(Alignment.Right,  ps[1].Style.Alignment);
    }

    [Fact]
    public void Reader_DivWithTextAlignBecomesParagraph()
    {
        // 블록 자식 없는 <div>text</div> 가 단락으로 변환되어 div 의 text-align 적용.
        const string html = "<div style=\"text-align: right\">우측 정렬 텍스트</div>";
        var rt = HtmlReader.FromHtml(html);
        var p  = rt.EnumerateParagraphs().Single();
        Assert.Equal(Alignment.Right, p.Style.Alignment);
        Assert.Equal("우측 정렬 텍스트", p.GetPlainText());
    }

    // ── list-style-type: none → 마커 비표시 ─────────────────────────

    [Fact]
    public void Reader_UlListStyleNone_HideBullet()
    {
        // list-style-type:none 은 ListMarker 자체를 제거하지 않고 HideBullet=true 로 보존 —
        // 목차/링크 목록의 구조(들여쓰기·중첩)를 유지하기 위함.
        const string html = "<ul style=\"list-style-type: none\"><li>A</li><li>B</li></ul>";
        var rt = HtmlReader.FromHtml(html);
        var ps = rt.EnumerateParagraphs().ToList();
        Assert.Equal(2, ps.Count);
        Assert.NotNull(ps[0].Style.ListMarker);
        Assert.True  (ps[0].Style.ListMarker!.HideBullet);
        Assert.NotNull(ps[1].Style.ListMarker);
        Assert.True  (ps[1].Style.ListMarker!.HideBullet);
    }

    [Fact]
    public void Reader_UlListStyleNoneViaCssClass_HideBullet()
    {
        // <style> 블록의 .toc ul { list-style-type: none } 도 동일하게 HideBullet=true.
        const string html = """
            <html><head><style>
              .toc ul { list-style-type: none; }
            </style></head><body>
              <div class="toc"><ul><li><a href="#x">link</a></li></ul></div>
            </body></html>
            """;
        var rt = HtmlReader.FromHtml(html);
        var p  = rt.EnumerateParagraphs().Single();
        Assert.NotNull(p.Style.ListMarker);
        Assert.True  (p.Style.ListMarker!.HideBullet);
    }

    [Fact]
    public void RoundTrip_UlHideBullet_PreservesListStyleNone()
    {
        // HideBullet=true 인 목록은 라운드트립 시 <ul style="list-style-type:none"> 로 직렬화.
        var doc = new PolyDonkyument();
        var sec = new Section();
        for (int i = 0; i < 3; i++)
        {
            var p = new Paragraph();
            p.Style.ListMarker = new ListMarker { Kind = ListKind.Bullet, HideBullet = true };
            p.AddText($"item {i}");
            sec.Blocks.Add(p);
        }
        doc.Sections.Add(sec);
        var html = HtmlWriter.ToHtml(doc);
        Assert.Contains("list-style-type:none", html);

        var rt = HtmlReader.FromHtml(html);
        var ps = rt.Sections[0].Blocks.OfType<Paragraph>().ToList();
        Assert.Equal(3, ps.Count);
        Assert.All(ps, p => Assert.True(p.Style.ListMarker!.HideBullet));
    }

    [Fact]
    public void RoundTrip_TaskListStillShowsCheckboxEvenWithHideBullet()
    {
        // 체크리스트는 <ul class="checklist" style="list-style:none"> 패턴으로 입력되더라도
        // checkbox 가 마커를 대신하므로 HideBullet 은 false 로 (writer 가 list-style-type:none 출력 안 함).
        const string html = """
            <ul style="list-style-type: none">
              <li><input type="checkbox" checked> done</li>
              <li><input type="checkbox"> not done</li>
            </ul>
            """;
        var rt = HtmlReader.FromHtml(html);
        var ps = rt.EnumerateParagraphs().ToList();
        Assert.Equal(2, ps.Count);
        Assert.Equal(true,  ps[0].Style.ListMarker!.Checked);
        Assert.Equal(false, ps[1].Style.ListMarker!.Checked);
        Assert.False(ps[0].Style.ListMarker!.HideBullet);
        Assert.False(ps[1].Style.ListMarker!.HideBullet);
    }

    // ── <a> text-decoration: none → 밑줄 제거 ───────────────────────

    [Fact]
    public void Reader_AnchorTextDecorationNone_NoUnderline()
    {
        // 인라인 컨텍스트(`<p>` 안) 의 `<a>` 가 `text-decoration: none` 을 따라 밑줄 제거.
        const string html = "<p><a href=\"#x\" style=\"text-decoration: none\">link</a></p>";
        var rt  = HtmlReader.FromHtml(html);
        var run = rt.EnumerateParagraphs().Single().Runs.Single();
        Assert.False(run.Style.Underline);
        Assert.Equal("#x", run.Url);
    }

    [Fact]
    public void Reader_AnchorDefault_HasUnderline()
    {
        // 인라인 컨텍스트(`<p>` 안) 의 `<a>` 는 기본으로 밑줄 적용.
        const string html = "<p><a href=\"#x\">link</a></p>";
        var rt  = HtmlReader.FromHtml(html);
        var run = rt.EnumerateParagraphs().Single().Runs.Single();
        Assert.True(run.Style.Underline);
    }

    [Fact]
    public void Reader_AnchorInListItem_TextDecorationNoneViaCssClass()
    {
        // .toc a { text-decoration: none } 시나리오 — `<li>` 안 `<a>` 도 동일 동작.
        const string html = """
            <html><head><style>
              .toc a { text-decoration: none; }
            </style></head><body>
              <div class="toc"><ul><li><a href="#x">link</a></li></ul></div>
            </body></html>
            """;
        var rt  = HtmlReader.FromHtml(html);
        var run = rt.EnumerateParagraphs().Single().Runs.Single();
        Assert.False(run.Style.Underline);
    }

    [Fact]
    public void Reader_CssGridOddCells_LastRowPadded()
    {
        // 3개 셀을 2열 그리드에 배치 → 2행(2+1), 마지막 행은 빈 셀 패딩
        const string html = """
            <div style="display:grid;grid-template-columns:1fr 1fr">
              <div><p>A</p></div><div><p>B</p></div><div><p>C</p></div>
            </div>
            """;
        var rt    = HtmlReader.FromHtml(html);
        var table = rt.Sections[0].Blocks.OfType<Table>().FirstOrDefault();
        Assert.NotNull(table);
        Assert.Equal(2, table.Rows.Count);
        Assert.Equal(2, table.Rows[1].Cells.Count);
    }

    // ── border-bottom / em margins / CSS inheritance ──────────────────────

    [Fact]
    public void Reader_H1BorderBottom_AppliedFromCss()
    {
        // CSS h1 { border-bottom: 1px solid #cccccc } → ParagraphStyle.BorderBottomPt > 0
        const string html = """
            <html><head><style>
              h1 { border-bottom: 1px solid #cccccc }
            </style></head><body>
              <h1>Title</h1>
            </body></html>
            """;
        var doc = HtmlReader.FromHtml(html);
        var p   = doc.Sections[0].Blocks.OfType<Paragraph>().First();
        Assert.Equal(OutlineLevel.H1, p.Style.Outline);
        Assert.True(p.Style.BorderBottomPt > 0);
        Assert.NotNull(p.Style.BorderBottomColor);
    }

    [Fact]
    public void Reader_H1BorderBottom_TwoPixels()
    {
        // border-bottom: 2px solid #000000 → BorderBottomPt = 2 * 72/96 ≈ 1.5pt
        const string html = """
            <html><head><style>
              h1 { border-bottom: 2px solid #000000 }
            </style></head><body>
              <h1>H</h1>
            </body></html>
            """;
        var doc = HtmlReader.FromHtml(html);
        var p   = doc.Sections[0].Blocks.OfType<Paragraph>().First();
        Assert.Equal(OutlineLevel.H1, p.Style.Outline);
        Assert.True(p.Style.BorderBottomPt > 1.4 && p.Style.BorderBottomPt < 1.6);
    }

    [Fact]
    public void Reader_HeadingMarginTopPt_Resolved()
    {
        // h2 { margin-top: 30pt } → SpaceBeforePt ≈ 30pt > 20
        const string html = """
            <html><head><style>
              h2 { margin-top: 30pt }
            </style></head><body>
              <h2>Heading</h2>
            </body></html>
            """;
        var doc = HtmlReader.FromHtml(html);
        var p   = doc.Sections[0].Blocks.OfType<Paragraph>().First();
        Assert.Equal(OutlineLevel.H2, p.Style.Outline);
        Assert.True(p.Style.SpaceBeforePt > 20.0);
    }

    [Fact]
    public void Reader_HeadingMarginTopPx_InlineStyle()
    {
        // inline style margin-top:40px → SpaceBeforePt = 40*72/96 = 30pt > 20
        const string html = """
            <html><head></head><body>
              <h2 style="margin-top:40px">Heading</h2>
            </body></html>
            """;
        var doc = HtmlReader.FromHtml(html);
        var p   = doc.Sections[0].Blocks.OfType<Paragraph>().First();
        Assert.Equal(OutlineLevel.H2, p.Style.Outline);
        Assert.True(p.Style.SpaceBeforePt > 20.0);
    }

    [Fact]
    public void Reader_HeadingMarginTopEm_Resolved()
    {
        // h2 inline style margin-top:1.5em — bypasses InlineCssClassRules to isolate em parsing.
        // h2 기본 font-size = 20pt → 1.5em = 30pt; OR AngleSharp normalizes to px (1.5*16=24px=18pt)
        // Either way, SpaceBeforePt should be > 0
        const string html = """
            <html><head></head><body>
              <h2 style="margin-top:1.5em">Heading</h2>
            </body></html>
            """;
        var doc = HtmlReader.FromHtml(html);
        var p   = doc.Sections[0].Blocks.OfType<Paragraph>().First();
        Assert.Equal(OutlineLevel.H2, p.Style.Outline);
        Assert.True(p.Style.SpaceBeforePt > 0);
    }

    [Fact]
    public void Reader_HeadingMarginBottomEm_Resolved()
    {
        // h3 기본 font-size = 17pt. margin-bottom: 0.5em → 8.5pt > 5
        const string html = """
            <html><head><style>
              h3 { margin-bottom: 0.5em }
            </style></head><body>
              <h3>H3</h3>
            </body></html>
            """;
        var doc = HtmlReader.FromHtml(html);
        var p   = doc.Sections[0].Blocks.OfType<Paragraph>().First();
        Assert.Equal(OutlineLevel.H3, p.Style.Outline);
        Assert.True(p.Style.SpaceAfterPt > 5.0);
    }

    [Fact]
    public void Reader_BodyLineHeight_PropagatedToParagraph()
    {
        // body { line-height: 1.8 } → p 단락의 LineHeightFactor ≈ 1.8
        const string html = """
            <html><head><style>
              body { line-height: 1.8 }
            </style></head><body>
              <p>Text</p>
            </body></html>
            """;
        var doc = HtmlReader.FromHtml(html);
        var p   = doc.Sections[0].Blocks.OfType<Paragraph>().First();
        Assert.True(Math.Abs(p.Style.LineHeightFactor - 1.8) < 0.01);
    }

    [Fact]
    public void Reader_BodyColor_PropagatedToRuns()
    {
        // body { color: #336699 } → p 내 run 의 Foreground 가 해당 색으로 설정됨
        const string html = """
            <html><head><style>
              body { color: #336699 }
            </style></head><body>
              <p>Text</p>
            </body></html>
            """;
        var doc = HtmlReader.FromHtml(html);
        var p   = doc.Sections[0].Blocks.OfType<Paragraph>().First();
        Assert.NotEmpty(p.Runs);
        var fg = p.Runs[0].Style.Foreground;
        Assert.NotNull(fg);
        Assert.Equal(0x33, fg!.Value.R);
        Assert.Equal(0x66, fg!.Value.G);
        Assert.Equal(0x99, fg!.Value.B);
    }

    [Fact]
    public void Reader_ChildColorOverridesBodyColor()
    {
        // body { color: #333333 } but h1 { color: #000000 } → h1 text 는 black
        const string html = """
            <html><head><style>
              body { color: #333333 }
              h1   { color: #000000 }
            </style></head><body>
              <h1>H</h1>
            </body></html>
            """;
        var doc = HtmlReader.FromHtml(html);
        var p   = doc.Sections[0].Blocks.OfType<Paragraph>().First();
        Assert.NotEmpty(p.Runs);
        var fg = p.Runs[0].Style.Foreground;
        Assert.NotNull(fg);
        // h1 style 에 명시된 #000000 이 body #333333 보다 우선
        Assert.Equal(0x00, fg!.Value.R);
        Assert.Equal(0x00, fg!.Value.G);
        Assert.Equal(0x00, fg!.Value.B);
    }

    // ── 도형 확장 라운드트립 (data-pd-* 속성) ─────────────────────────────────

    [Fact]
    public void RoundTrip_Shape_RegularPolygonPreservesKindAndSideCount()
    {
        var doc = new PolyDonkyument();
        var sec = new Section();
        sec.Blocks.Add(new ShapeObject
        {
            Kind = ShapeKind.RegularPolygon, WidthMm = 40, HeightMm = 40, SideCount = 7,
        });
        doc.Sections.Add(sec);
        var rt = HtmlReader.FromHtml(HtmlWriter.ToHtml(doc));
        var s  = rt.Sections[0].Blocks.OfType<ShapeObject>().Single();
        Assert.Equal(ShapeKind.RegularPolygon, s.Kind);
        Assert.Equal(7, s.SideCount);
    }

    [Fact]
    public void RoundTrip_Shape_StarPreservesKindSideCountInnerRatio()
    {
        var doc = new PolyDonkyument();
        var sec = new Section();
        sec.Blocks.Add(new ShapeObject
        {
            Kind = ShapeKind.Star, WidthMm = 40, HeightMm = 40, SideCount = 6, InnerRadiusRatio = 0.4,
        });
        doc.Sections.Add(sec);
        var rt = HtmlReader.FromHtml(HtmlWriter.ToHtml(doc));
        var s  = rt.Sections[0].Blocks.OfType<ShapeObject>().Single();
        Assert.Equal(ShapeKind.Star,  s.Kind);
        Assert.Equal(6,               s.SideCount);
        Assert.InRange(s.InnerRadiusRatio, 0.39, 0.41);
    }

    [Fact]
    public void RoundTrip_Shape_RotationAngle()
    {
        var doc = new PolyDonkyument();
        var sec = new Section();
        sec.Blocks.Add(new ShapeObject
        {
            Kind = ShapeKind.Rectangle, WidthMm = 50, HeightMm = 30, RotationAngleDeg = 30,
        });
        doc.Sections.Add(sec);
        var rt = HtmlReader.FromHtml(HtmlWriter.ToHtml(doc));
        var s  = rt.Sections[0].Blocks.OfType<ShapeObject>().Single();
        Assert.InRange(s.RotationAngleDeg, 29.5, 30.5);
    }

    [Fact]
    public void RoundTrip_Shape_ArrowsPreserved()
    {
        var doc = new PolyDonkyument();
        var sec = new Section();
        var ln  = new ShapeObject
        {
            Kind = ShapeKind.Line, WidthMm = 80, HeightMm = 5,
            StartArrow = ShapeArrow.Open, EndArrow = ShapeArrow.Filled,
        };
        ln.Points.Add(new ShapePoint { X = 0,  Y = 2 });
        ln.Points.Add(new ShapePoint { X = 80, Y = 2 });
        sec.Blocks.Add(ln);
        doc.Sections.Add(sec);
        var rt = HtmlReader.FromHtml(HtmlWriter.ToHtml(doc));
        var s  = rt.Sections[0].Blocks.OfType<ShapeObject>().Single();
        Assert.Equal(ShapeArrow.Open,   s.StartArrow);
        Assert.Equal(ShapeArrow.Filled, s.EndArrow);
    }

    [Fact]
    public void RoundTrip_Shape_StrokeDashAllVariants()
    {
        foreach (var dash in new[] { StrokeDash.Dashed, StrokeDash.Dotted, StrokeDash.DashDot })
        {
            var doc = new PolyDonkyument();
            var sec = new Section();
            sec.Blocks.Add(new ShapeObject
            {
                Kind = ShapeKind.Rectangle, WidthMm = 30, HeightMm = 20, StrokeDash = dash,
            });
            doc.Sections.Add(sec);
            var rt = HtmlReader.FromHtml(HtmlWriter.ToHtml(doc));
            var s  = rt.Sections[0].Blocks.OfType<ShapeObject>().Single();
            Assert.Equal(dash, s.StrokeDash);
        }
    }

    [Fact]
    public void RoundTrip_Shape_LabelStylingPreserved()
    {
        var doc = new PolyDonkyument();
        var sec = new Section();
        sec.Blocks.Add(new ShapeObject
        {
            Kind = ShapeKind.Rectangle, WidthMm = 60, HeightMm = 30,
            LabelText = "라벨", LabelFontSizePt = 14, LabelBold = true, LabelItalic = true,
            LabelColor = "#FF0000", LabelBackgroundColor = "#FFFF00",
            LabelHAlign = ShapeLabelHAlign.Right, LabelVAlign = ShapeLabelVAlign.Bottom,
            LabelOffsetXMm = 2, LabelOffsetYMm = -1,
        });
        doc.Sections.Add(sec);
        var rt = HtmlReader.FromHtml(HtmlWriter.ToHtml(doc));
        var s  = rt.Sections[0].Blocks.OfType<ShapeObject>().Single();
        Assert.Equal("라벨",                  s.LabelText);
        Assert.InRange(s.LabelFontSizePt,    13.5, 14.5);
        Assert.True(s.LabelBold);
        Assert.True(s.LabelItalic);
        Assert.Equal("#FF0000",                       s.LabelColor, ignoreCase: true);
        Assert.Equal("#FFFF00",                       s.LabelBackgroundColor, ignoreCase: true);
        Assert.Equal(ShapeLabelHAlign.Right,           s.LabelHAlign);
        Assert.Equal(ShapeLabelVAlign.Bottom,          s.LabelVAlign);
        Assert.InRange(s.LabelOffsetXMm,  1.9, 2.1);
        Assert.InRange(s.LabelOffsetYMm, -1.1, -0.9);
    }

    [Fact]
    public void RoundTrip_Shape_FillOpacityPreserved()
    {
        var doc = new PolyDonkyument();
        var sec = new Section();
        sec.Blocks.Add(new ShapeObject
        {
            Kind = ShapeKind.Ellipse, WidthMm = 40, HeightMm = 40,
            FillColor = "#3366CC", FillOpacity = 0.4,
        });
        doc.Sections.Add(sec);
        var rt = HtmlReader.FromHtml(HtmlWriter.ToHtml(doc));
        var s  = rt.Sections[0].Blocks.OfType<ShapeObject>().Single();
        Assert.InRange(s.FillOpacity, 0.39, 0.41);
    }

    [Fact]
    public void RoundTrip_Shape_OverlayPositionPreserved()
    {
        var doc = new PolyDonkyument();
        var sec = new Section();
        sec.Blocks.Add(new ShapeObject
        {
            Kind = ShapeKind.Rectangle, WidthMm = 30, HeightMm = 20,
            WrapMode = ImageWrapMode.InFrontOfText,
            AnchorPageIndex = 2, OverlayXMm = 12.5, OverlayYMm = 7.25,
        });
        doc.Sections.Add(sec);
        var rt = HtmlReader.FromHtml(HtmlWriter.ToHtml(doc));
        var s  = rt.Sections[0].Blocks.OfType<ShapeObject>().Single();
        Assert.Equal(ImageWrapMode.InFrontOfText, s.WrapMode);
        Assert.Equal(2,                            s.AnchorPageIndex);
        Assert.InRange(s.OverlayXMm, 12.4, 12.6);
        Assert.InRange(s.OverlayYMm,  7.2,  7.3);
    }

    // ── 페이지 설정 확장 라운드트립 ─────────────────────────────────────────────

    [Fact]
    public void RoundTrip_HeaderFooter_LeftCenterRight()
    {
        var doc = new PolyDonkyument();
        var sec = new Section();
        sec.Page.Header.Left   = HeaderFooterSlot.FromText("문서 제목");
        sec.Page.Header.Center = HeaderFooterSlot.FromText("회사명");
        sec.Page.Footer.Right  = HeaderFooterSlot.FromText("페이지");
        doc.Sections.Add(sec);

        var html = HtmlWriter.ToHtml(doc);
        Assert.Contains("<header class=\"pd-header\">", html);
        Assert.Contains("<footer class=\"pd-footer\">", html);

        var rt = HtmlReader.FromHtml(html);
        var p  = rt.Sections[0].Page;
        Assert.Equal("문서 제목", p.Header.Left.GetPlainText());
        Assert.Equal("회사명",   p.Header.Center.GetPlainText());
        Assert.Equal("페이지",   p.Footer.Right.GetPlainText());
        Assert.True(p.Header.Right.IsEmpty);
        Assert.True(p.Footer.Left.IsEmpty);
    }

    [Fact]
    public void RoundTrip_PageSettings_MultiColumnLayout()
    {
        var doc = new PolyDonkyument();
        var sec = new Section();
        sec.Page.ColumnCount             = 3;
        sec.Page.ColumnGapMm             = 10;
        sec.Page.ColumnDividerVisible    = true;
        sec.Page.ColumnDividerStyle      = ColumnDividerStyle.Solid;
        sec.Page.ColumnDividerColor      = "#0000FF";
        sec.Page.ColumnDividerThicknessPt = 1.2;
        doc.Sections.Add(sec);

        var rt = HtmlReader.FromHtml(HtmlWriter.ToHtml(doc));
        var p  = rt.Sections[0].Page;
        Assert.Equal(3,                       p.ColumnCount);
        Assert.InRange(p.ColumnGapMm, 9.5, 10.5);
        Assert.True(p.ColumnDividerVisible);
        Assert.Equal(ColumnDividerStyle.Solid, p.ColumnDividerStyle);
        Assert.Equal("#0000FF",                p.ColumnDividerColor, ignoreCase: true);
        Assert.InRange(p.ColumnDividerThicknessPt, 1.1, 1.3);
    }

    [Fact]
    public void HtmlWriter_EmitsCssZIndexForExplicitZOrder()
    {
        var doc = new PolyDonkyument();
        var sec = new Section();
        sec.Blocks.Add(new ShapeObject
        {
            Kind = ShapeKind.Rectangle, WidthMm = 50, HeightMm = 30, ZOrder = -5,
        });
        sec.Blocks.Add(new ShapeObject
        {
            Kind = ShapeKind.Ellipse, WidthMm = 40, HeightMm = 40, ZOrder = 12,
        });
        doc.Sections.Add(sec);
        var html = HtmlWriter.ToHtml(doc);
        // position:relative 가 z-index 와 함께 출력되어야 함 (브라우저에서 z-index 가 동작하도록).
        Assert.Contains("position:relative", html);
        Assert.Contains("z-index:-5", html);
        Assert.Contains("z-index:12", html);
    }

    [Fact]
    public void HtmlWriter_EmitsAutoCssZIndexForContainedShapes()
    {
        // ZOrder=0 자동 그룹: 외곽 도형이 안쪽 도형을 포함 → 안쪽 도형에 z-index >= 1 부여.
        var doc = new PolyDonkyument();
        var sec = new Section();
        sec.Blocks.Add(new ShapeObject
        {
            Kind = ShapeKind.Rectangle, WidthMm = 100, HeightMm = 100,
            WrapMode = ImageWrapMode.InFrontOfText,
            OverlayXMm = 0, OverlayYMm = 0,
        });
        sec.Blocks.Add(new ShapeObject
        {
            Kind = ShapeKind.Ellipse, WidthMm = 30, HeightMm = 30,
            WrapMode = ImageWrapMode.InFrontOfText,
            OverlayXMm = 20, OverlayYMm = 20,
        });
        doc.Sections.Add(sec);
        var html = HtmlWriter.ToHtml(doc);
        // 안쪽 도형(타원)에 z-index:1 이 출력돼야 함. 외곽 도형(사각형)은 depth 0 → z-index 출력 안 됨.
        Assert.Contains("z-index:1", html);
    }

    [Fact]
    public void HtmlWriter_OmitsCssZIndexForAutoZeroDepth()
    {
        // 두 도형이 겹치지 않으면 컨테인먼트 깊이 0 → z-index 출력하지 않음.
        var doc = new PolyDonkyument();
        var sec = new Section();
        sec.Blocks.Add(new ShapeObject
        {
            Kind = ShapeKind.Rectangle, WidthMm = 30, HeightMm = 30,
            WrapMode = ImageWrapMode.InFrontOfText,
            OverlayXMm = 0, OverlayYMm = 0,
        });
        sec.Blocks.Add(new ShapeObject
        {
            Kind = ShapeKind.Ellipse, WidthMm = 30, HeightMm = 30,
            WrapMode = ImageWrapMode.InFrontOfText,
            OverlayXMm = 100, OverlayYMm = 100,
        });
        doc.Sections.Add(sec);
        var html = HtmlWriter.ToHtml(doc);
        Assert.DoesNotContain("z-index:", html);
    }

    [Fact]
    public void RoundTrip_Shape_AutoCssZIndexDoesNotPolluteZOrder()
    {
        // 자동 컨테인먼트로 emit 된 CSS z-index 가 reader 에서 명시적 ZOrder 로 해석되면 안 됨.
        // (reader 는 data-pd-z-order 만 보고 ZOrder 를 채운다.)
        var doc = new PolyDonkyument();
        var sec = new Section();
        sec.Blocks.Add(new ShapeObject
        {
            Kind = ShapeKind.Rectangle, WidthMm = 100, HeightMm = 100,
            WrapMode = ImageWrapMode.InFrontOfText, OverlayXMm = 0, OverlayYMm = 0,
        });
        sec.Blocks.Add(new ShapeObject
        {
            Kind = ShapeKind.Ellipse, WidthMm = 30, HeightMm = 30,
            WrapMode = ImageWrapMode.InFrontOfText, OverlayXMm = 20, OverlayYMm = 20,
        });
        doc.Sections.Add(sec);
        var html = HtmlWriter.ToHtml(doc);
        Assert.Contains("z-index:1", html);  // CSS 출력은 됐고

        var rt = HtmlReader.FromHtml(html);
        var shapes = rt.Sections[0].Blocks.OfType<ShapeObject>().ToList();
        Assert.All(shapes, s => Assert.Equal(0, s.ZOrder));  // ZOrder 는 0 (자동) 유지
    }

    // ── <hr> 두께·선스타일 라운드트립 ──────────────────────────────────

    [Fact]
    public void RoundTrip_Hr_ThickSolidPreservesThicknessAndColor()
    {
        const string html = """
            <html><head><style>
              hr.thick { border: 0; border-top: 3px solid #000000; margin: 20px 0; }
            </style></head><body>
              <hr class="thick">
            </body></html>
            """;
        var doc = HtmlReader.FromHtml(html);
        var thb = doc.Sections[0].Blocks.OfType<ThematicBreakBlock>().Single();
        Assert.Equal(ThematicLineStyle.Solid, thb.LineStyle);
        Assert.InRange(thb.ThicknessPt, 2.0, 2.5);   // 3px ≈ 2.25pt
        Assert.Equal("#000000", thb.LineColor, ignoreCase: true);

        var rt = HtmlReader.FromHtml(HtmlWriter.ToHtml(doc));
        var thb2 = rt.Sections[0].Blocks.OfType<ThematicBreakBlock>().Single();
        Assert.Equal(thb.LineStyle,                          thb2.LineStyle);
        Assert.InRange(Math.Abs(thb2.ThicknessPt - thb.ThicknessPt), 0, 0.05);
        Assert.Equal(thb.LineColor, thb2.LineColor, ignoreCase: true);
    }

    [Fact]
    public void RoundTrip_Hr_DashedDottedDoublePreserveStyle()
    {
        const string html = """
            <html><head><style>
              hr.dashed { border: 0; border-top: 1px dashed #666666; }
              hr.dotted { border: 0; border-top: 1px dotted #999999; }
            </style></head><body>
              <hr class="dashed">
              <hr class="dotted">
              <hr style="border:0;border-top:3px double #333333">
            </body></html>
            """;
        var doc = HtmlReader.FromHtml(html);
        var thbs = doc.Sections[0].Blocks.OfType<ThematicBreakBlock>().ToList();
        Assert.Equal(3, thbs.Count);
        Assert.Equal(ThematicLineStyle.Dashed, thbs[0].LineStyle);
        Assert.Equal(ThematicLineStyle.Dotted, thbs[1].LineStyle);
        Assert.Equal(ThematicLineStyle.Double, thbs[2].LineStyle);

        var html2 = HtmlWriter.ToHtml(doc);
        Assert.Contains("border-top:", html2);
        Assert.Contains("dashed", html2);
        Assert.Contains("dotted", html2);
        Assert.Contains("double", html2);

        var rt = HtmlReader.FromHtml(html2);
        var thbs2 = rt.Sections[0].Blocks.OfType<ThematicBreakBlock>().ToList();
        Assert.Equal(3, thbs2.Count);
        Assert.Equal(ThematicLineStyle.Dashed, thbs2[0].LineStyle);
        Assert.Equal(ThematicLineStyle.Dotted, thbs2[1].LineStyle);
        Assert.Equal(ThematicLineStyle.Double, thbs2[2].LineStyle);
    }

    [Fact]
    public void RoundTrip_Hr_InlineShorthandMarginOverridesClassLonghand()
    {
        // <hr style="margin: 2px 0"> 는 그 자체로 작은 분수 막대 의도. CSS 클래스에서 expand 된
        // longhand(margin-top:20px 등) 가 inline shorthand 를 가리지 않아야 한다.
        const string html = """
            <html><head><style>
              hr { margin: 20px 0; border-top: 1px solid #000; }
            </style></head><body>
              <hr style="margin: 2px 0; border-top: 1px solid #000">
            </body></html>
            """;
        var doc = HtmlReader.FromHtml(html);
        var thb = doc.Sections[0].Blocks.OfType<ThematicBreakBlock>().Single();
        // 1.5pt 부근(2px) — 15pt(20px) 가 아니어야 한다.
        Assert.InRange(thb.MarginPt, 1.0, 2.0);
    }

    [Fact]
    public void HtmlWriter_DefaultHr_OmitsBorderStyle()
    {
        // 기본 ThematicBreakBlock(두께·스타일·색상 모두 기본) 은 border-top 없이 단순 <hr> 출력.
        var doc = new PolyDonkyument();
        var sec = new Section();
        sec.Blocks.Add(new ThematicBreakBlock());
        doc.Sections.Add(sec);
        var html = HtmlWriter.ToHtml(doc);
        Assert.Contains("<hr>", html);
        Assert.DoesNotContain("border-top", html);
    }

    [Fact]
    public void RoundTrip_Shape_ZOrderPreserved()
    {
        var doc = new PolyDonkyument();
        var sec = new Section();
        sec.Blocks.Add(new ShapeObject
        {
            Kind = ShapeKind.Rectangle, WidthMm = 50, HeightMm = 30, ZOrder = -3,
        });
        sec.Blocks.Add(new ShapeObject
        {
            Kind = ShapeKind.Ellipse,   WidthMm = 40, HeightMm = 40, ZOrder = 0,
        });
        sec.Blocks.Add(new ShapeObject
        {
            Kind = ShapeKind.Star,      WidthMm = 30, HeightMm = 30, ZOrder = 7, SideCount = 5,
        });
        doc.Sections.Add(sec);

        var html = HtmlWriter.ToHtml(doc);
        Assert.Contains("data-pd-z-order=\"-3\"", html);
        Assert.Contains("data-pd-z-order=\"7\"",  html);

        var rt = HtmlReader.FromHtml(html);
        var shapes = rt.Sections[0].Blocks.OfType<ShapeObject>().ToList();
        Assert.Equal(3, shapes.Count);
        Assert.Equal(-3, shapes[0].ZOrder);
        Assert.Equal( 0, shapes[1].ZOrder);
        Assert.Equal( 7, shapes[2].ZOrder);
    }

    [Fact]
    public void RoundTrip_PageSettings_PageNumberStartAndPaperColor()
    {
        var doc = new PolyDonkyument();
        var sec = new Section();
        sec.Page.PageNumberStart = 5;
        sec.Page.PaperColor      = "#FFFEEE";
        sec.Page.MarginHeaderMm  = 15;
        sec.Page.MarginFooterMm  = 12;
        doc.Sections.Add(sec);

        var rt = HtmlReader.FromHtml(HtmlWriter.ToHtml(doc));
        var p  = rt.Sections[0].Page;
        Assert.Equal(5,           p.PageNumberStart);
        Assert.Equal("#FFFEEE",    p.PaperColor, ignoreCase: true);
        Assert.InRange(p.MarginHeaderMm, 14.5, 15.5);
        Assert.InRange(p.MarginFooterMm, 11.5, 12.5);
    }
}

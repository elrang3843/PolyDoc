using PolyDonky.Codecs.Markdown;
using PolyDonky.Core;

namespace PolyDonky.Codecs.Markdown.Tests;

public class MarkdownTests
{
    [Fact]
    public void Reader_ExtractsAtxHeaderLevels()
    {
        const string source = "# H1\n\n## H2\n\n### H3\n";
        var doc = MarkdownReader.FromMarkdown(source);

        var paragraphs = doc.EnumerateParagraphs().ToList();
        Assert.Equal(3, paragraphs.Count);
        Assert.Equal(OutlineLevel.H1, paragraphs[0].Style.Outline);
        Assert.Equal(OutlineLevel.H2, paragraphs[1].Style.Outline);
        Assert.Equal(OutlineLevel.H3, paragraphs[2].Style.Outline);
    }

    [Fact]
    public void Reader_BuildsBoldAndItalicRuns()
    {
        var doc = MarkdownReader.FromMarkdown("plain **bold** and *italic* end");
        var p = doc.EnumerateParagraphs().Single();

        Assert.Contains(p.Runs, r => r.Style.Bold && r.Text == "bold");
        Assert.Contains(p.Runs, r => r.Style.Italic && r.Text == "italic");
    }

    [Fact]
    public void Reader_DetectsBulletAndOrderedLists()
    {
        const string source = "- item one\n- item two\n\n1. first\n2. second\n";
        var doc = MarkdownReader.FromMarkdown(source);
        var paragraphs = doc.EnumerateParagraphs().ToList();

        Assert.Equal(4, paragraphs.Count);
        Assert.Equal(ListKind.Bullet, paragraphs[0].Style.ListMarker!.Kind);
        Assert.Equal(ListKind.Bullet, paragraphs[1].Style.ListMarker!.Kind);
        Assert.Equal(ListKind.OrderedDecimal, paragraphs[2].Style.ListMarker!.Kind);
        Assert.Equal(2, paragraphs[3].Style.ListMarker!.OrderedNumber);
    }

    [Fact]
    public void Writer_RendersHeaderHashesByOutlineLevel()
    {
        var doc = new PolyDonkyument();
        var section = new Section();
        doc.Sections.Add(section);
        var h2 = new Paragraph { Style = { Outline = OutlineLevel.H2 } };
        h2.AddText("section");
        section.Blocks.Add(h2);

        var rendered = MarkdownWriter.ToMarkdown(doc);

        Assert.StartsWith("## section", rendered);
    }

    [Fact]
    public void Writer_EmitsBulletAndOrderedMarkers()
    {
        var doc = new PolyDonkyument();
        var section = new Section();
        doc.Sections.Add(section);
        var bullet = new Paragraph();
        bullet.Style.ListMarker = new ListMarker { Kind = ListKind.Bullet };
        bullet.AddText("a");
        section.Blocks.Add(bullet);
        var ordered = new Paragraph();
        ordered.Style.ListMarker = new ListMarker { Kind = ListKind.OrderedDecimal, OrderedNumber = 3 };
        ordered.AddText("b");
        section.Blocks.Add(ordered);

        var rendered = MarkdownWriter.ToMarkdown(doc);

        Assert.Contains("- a", rendered);
        Assert.Contains("3. b", rendered);
    }

    [Fact]
    public void RoundTrip_HeadersAndListsArePreserved()
    {
        const string source =
            "# 제목\n\n본문 단락\n\n- 항목 A\n- 항목 B\n\n1. 첫째\n2. 둘째\n";
        var doc = MarkdownReader.FromMarkdown(source);
        var rendered = MarkdownWriter.ToMarkdown(doc);
        var reparsed = MarkdownReader.FromMarkdown(rendered);

        var original = doc.EnumerateParagraphs().Select(p => (p.Style.Outline, p.Style.ListMarker?.Kind)).ToList();
        var roundTripped = reparsed.EnumerateParagraphs().Select(p => (p.Style.Outline, p.Style.ListMarker?.Kind)).ToList();
        Assert.Equal(original, roundTripped);
    }

    [Fact]
    public void Reader_FencedCodeBlock_BecomesMonospaceParagraph()
    {
        const string source = "```\nint main() {\n  return 0;\n}\n```\n";
        var doc = MarkdownReader.FromMarkdown(source);

        var p = doc.EnumerateParagraphs().Single();
        Assert.Single(p.Runs);
        Assert.Contains("monospace", p.Runs[0].Style.FontFamily, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("int main()", p.Runs[0].Text);
    }

    [Fact]
    public void Reader_InlineCode_HintsMonospaceFont()
    {
        var doc = MarkdownReader.FromMarkdown("call `printf()` to print");

        var p = doc.EnumerateParagraphs().Single();
        var codeRun = p.Runs.Single(r => r.Text == "printf()");
        Assert.NotNull(codeRun.Style.FontFamily);
        Assert.Contains("monospace", codeRun.Style.FontFamily!, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Reader_Quote_FlattensToParagraphInPhaseA()
    {
        // Phase A 의 Core 모델은 인용 블록을 별도 노드로 갖고 있지 않다.
        // Markdig 의 QuoteBlock 은 일반 단락으로 격하되어야 한다.
        var doc = MarkdownReader.FromMarkdown("> 인용된 한 줄\n");

        var p = doc.EnumerateParagraphs().Single();
        Assert.Equal("인용된 한 줄", p.GetPlainText());
    }

    [Fact]
    public void Reader_Link_UnderlinesText()
    {
        var doc = MarkdownReader.FromMarkdown("[홈](https://example.com)");

        var p = doc.EnumerateParagraphs().Single();
        Assert.All(p.Runs, r => Assert.True(r.Style.Underline));
        Assert.Equal("홈", p.GetPlainText());
    }

    [Fact]
    public void Reader_NestedBoldItalic_MergesStyles()
    {
        var doc = MarkdownReader.FromMarkdown("***둘 다***");

        var p = doc.EnumerateParagraphs().Single();
        Assert.Contains(p.Runs, r => r.Style.Bold && r.Style.Italic && r.Text == "둘 다");
    }

    // ── GFM 확장 ────────────────────────────────────────────────────

    [Fact]
    public void Reader_Strikethrough_SetsRunStyle()
    {
        var doc = MarkdownReader.FromMarkdown("~~지움~~");
        var p   = doc.EnumerateParagraphs().Single();
        Assert.Contains(p.Runs, r => r.Style.Strikethrough && r.Text == "지움");
    }

    [Fact]
    public void Reader_SubAndSuperscript_SetsRunFlags()
    {
        var doc = MarkdownReader.FromMarkdown("H~2~O 의 X^2^");
        var p   = doc.EnumerateParagraphs().Single();
        Assert.Contains(p.Runs, r => r.Style.Subscript   && r.Text == "2");
        Assert.Contains(p.Runs, r => r.Style.Superscript && r.Text == "2");
    }

    [Fact]
    public void Reader_TaskList_PopulatesCheckedFlag()
    {
        const string source = "- [x] 한 일\n- [ ] 할 일\n";
        var doc = MarkdownReader.FromMarkdown(source);
        var ps  = doc.EnumerateParagraphs().ToList();

        Assert.Equal(2, ps.Count);
        Assert.True (ps[0].Style.ListMarker!.Checked);
        Assert.False(ps[1].Style.ListMarker!.Checked);
    }

    [Fact]
    public void Reader_BlockQuote_RecordsLevel()
    {
        var doc = MarkdownReader.FromMarkdown("> 1단\n>\n> > 2단\n");
        var ps  = doc.EnumerateParagraphs().ToList();

        Assert.Contains(ps, p => p.Style.QuoteLevel == 1);
        Assert.Contains(ps, p => p.Style.QuoteLevel == 2);
    }

    [Fact]
    public void Reader_ThematicBreak_CreatesFlaggedParagraph()
    {
        var doc = MarkdownReader.FromMarkdown("앞\n\n---\n\n뒤\n");
        var ps  = doc.EnumerateParagraphs().ToList();
        Assert.Contains(ps, p => p.Style.IsThematicBreak);
    }

    [Fact]
    public void Reader_FencedCode_CapturesLanguage()
    {
        const string source = "```python\nprint('hi')\n```\n";
        var doc = MarkdownReader.FromMarkdown(source);
        var p   = doc.EnumerateParagraphs().Single();
        Assert.Equal("python", p.Style.CodeLanguage);
        Assert.Contains("print", p.GetPlainText());
    }

    [Fact]
    public void Reader_Link_StoresUrl()
    {
        var doc = MarkdownReader.FromMarkdown("[홈](https://example.com)");
        var p   = doc.EnumerateParagraphs().Single();
        Assert.All(p.Runs, r => Assert.Equal("https://example.com", r.Url));
    }

    [Fact]
    public void Reader_Autolink_CreatesUrlRun()
    {
        var doc = MarkdownReader.FromMarkdown("연결: <https://example.com>");
        var p   = doc.EnumerateParagraphs().Single();
        Assert.Contains(p.Runs, r => r.Url == "https://example.com");
    }

    [Fact]
    public void Reader_PipeTable_CreatesTableBlock()
    {
        const string source =
            "| 이름 | 나이 |\n" +
            "| :--- | ---: |\n" +
            "| 홍길동 | 30 |\n" +
            "| 김영희 | 25 |\n";

        var doc   = MarkdownReader.FromMarkdown(source);
        var table = doc.Sections[0].Blocks.OfType<PolyDonky.Core.Table>().Single();

        Assert.Equal(3, table.Rows.Count);
        Assert.True(table.Rows[0].IsHeader);
        Assert.Equal(2, table.Rows[0].Cells.Count);
        Assert.Equal(CellTextAlign.Left,  table.Rows[1].Cells[0].TextAlign);
        Assert.Equal(CellTextAlign.Right, table.Rows[1].Cells[1].TextAlign);
    }

    [Fact]
    public void Reader_SoloImageParagraph_CreatesImageBlock()
    {
        var doc = MarkdownReader.FromMarkdown("![대체](path/to/img.png)");
        var img = doc.Sections[0].Blocks.OfType<ImageBlock>().Single();
        Assert.Equal("대체", img.Description);
        Assert.Equal("path/to/img.png", img.ResourcePath);
        Assert.Equal("image/png", img.MediaType);
    }

    [Fact]
    public void Reader_Math_InlineAndBlock_PreserveLatex()
    {
        // 별행 수식은 펜스 형식 ($$ 줄바꿈) 으로 표기해야 MathBlock 으로 파싱된다.
        const string source = "인라인 $a^2 + b^2$\n\n$$\nE=mc^2\n$$\n";
        var doc = MarkdownReader.FromMarkdown(source);
        var ps  = doc.EnumerateParagraphs().ToList();

        Assert.Contains(ps[0].Runs, r => r.LatexSource == "a^2 + b^2" && !r.IsDisplayEquation);
        Assert.Contains(ps[1].Runs, r => r.LatexSource!.Contains("E=mc^2") && r.IsDisplayEquation);
    }

    // ── 작성기 (Writer) ─────────────────────────────────────────────

    [Fact]
    public void Writer_Strikethrough_EmitsTilde()
    {
        var doc = new PolyDonkyument();
        var sec = new Section(); doc.Sections.Add(sec);
        var p = new Paragraph();
        p.AddText("지움", new RunStyle { Strikethrough = true });
        sec.Blocks.Add(p);

        var md = MarkdownWriter.ToMarkdown(doc);
        Assert.Contains("~~지움~~", md);
    }

    [Fact]
    public void Writer_TaskList_EmitsCheckboxes()
    {
        var doc = new PolyDonkyument();
        var sec = new Section(); doc.Sections.Add(sec);
        var p1 = new Paragraph();
        p1.Style.ListMarker = new ListMarker { Kind = ListKind.Bullet, Checked = true };
        p1.AddText("done");
        var p2 = new Paragraph();
        p2.Style.ListMarker = new ListMarker { Kind = ListKind.Bullet, Checked = false };
        p2.AddText("todo");
        sec.Blocks.Add(p1); sec.Blocks.Add(p2);

        var md = MarkdownWriter.ToMarkdown(doc);
        Assert.Contains("- [x] done", md);
        Assert.Contains("- [ ] todo", md);
    }

    [Fact]
    public void Writer_Link_EmitsBracketUrl()
    {
        var doc = new PolyDonkyument();
        var sec = new Section(); doc.Sections.Add(sec);
        var p = new Paragraph();
        p.Runs.Add(new Run { Text = "홈", Style = new RunStyle(), Url = "https://example.com" });
        sec.Blocks.Add(p);

        Assert.Contains("[홈](https://example.com)", MarkdownWriter.ToMarkdown(doc));
    }

    [Fact]
    public void Writer_FencedCode_WrapsWithLanguage()
    {
        var doc = new PolyDonkyument();
        var sec = new Section(); doc.Sections.Add(sec);
        var p = new Paragraph { Style = { CodeLanguage = "cs" } };
        p.AddText("var x = 1;");
        sec.Blocks.Add(p);

        var md = MarkdownWriter.ToMarkdown(doc);
        Assert.Contains("```cs", md);
        Assert.Contains("var x = 1;", md);
        Assert.Contains("```\n", md);
    }

    [Fact]
    public void Writer_BlockQuote_PrefixesWithGreaterThan()
    {
        var doc = new PolyDonkyument();
        var sec = new Section(); doc.Sections.Add(sec);
        var p = new Paragraph { Style = { QuoteLevel = 2 } };
        p.AddText("이중 인용");
        sec.Blocks.Add(p);

        Assert.Contains("> > 이중 인용", MarkdownWriter.ToMarkdown(doc));
    }

    [Fact]
    public void Writer_ThematicBreak_EmitsThreeDashes()
    {
        var doc = new PolyDonkyument();
        var sec = new Section(); doc.Sections.Add(sec);
        sec.Blocks.Add(new Paragraph { Style = { IsThematicBreak = true } });

        Assert.Contains("---", MarkdownWriter.ToMarkdown(doc));
    }

    [Fact]
    public void Writer_PipeTable_EmitsHeaderAndAlignment()
    {
        var doc = new PolyDonkyument();
        var sec = new Section(); doc.Sections.Add(sec);
        var t = new PolyDonky.Core.Table();
        t.Columns.Add(new TableColumn());
        t.Columns.Add(new TableColumn());
        var hdr = new TableRow { IsHeader = true };
        var c1 = new TableCell { TextAlign = CellTextAlign.Left  };
        c1.Blocks.Add(Paragraph.Of("이름"));
        var c2 = new TableCell { TextAlign = CellTextAlign.Right };
        c2.Blocks.Add(Paragraph.Of("나이"));
        hdr.Cells.Add(c1); hdr.Cells.Add(c2);
        t.Rows.Add(hdr);

        var body = new TableRow();
        var b1 = new TableCell(); b1.Blocks.Add(Paragraph.Of("홍길동"));
        var b2 = new TableCell(); b2.Blocks.Add(Paragraph.Of("30"));
        body.Cells.Add(b1); body.Cells.Add(b2);
        t.Rows.Add(body);
        sec.Blocks.Add(t);

        var md = MarkdownWriter.ToMarkdown(doc);
        Assert.Contains("| 이름 | 나이 |", md);
        Assert.Contains("| --- | ---: |", md);
        Assert.Contains("| 홍길동 | 30 |", md);
    }

    [Fact]
    public void Writer_Image_EmitsAltAndPath()
    {
        var doc = new PolyDonkyument();
        var sec = new Section(); doc.Sections.Add(sec);
        sec.Blocks.Add(new ImageBlock
        {
            Description  = "대체",
            ResourcePath = "img/foo.png",
            MediaType    = "image/png",
        });

        Assert.Contains("![대체](img/foo.png)", MarkdownWriter.ToMarkdown(doc));
    }

    [Fact]
    public void Writer_InlineCode_WrapsBackticks()
    {
        var doc = new PolyDonkyument();
        var sec = new Section(); doc.Sections.Add(sec);
        var p = new Paragraph();
        p.AddText("plain ");
        p.AddText("printf()", new RunStyle { FontFamily = "Consolas, monospace" });
        sec.Blocks.Add(p);

        Assert.Contains("`printf()`", MarkdownWriter.ToMarkdown(doc));
    }

    // ── Writer: 이스케이프 보완 ─────────────────────────────────────────

    [Fact]
    public void Writer_EscapesCaretToPreventSuperscript()
    {
        // EmphasisExtras 가 ^text^ 를 superscript 로 파싱하므로 일반 ^ 는 이스케이프 필요.
        var doc = new PolyDonkyument();
        var sec = new Section(); doc.Sections.Add(sec);
        var p = new Paragraph();
        p.AddText("2^10 = 1024");
        sec.Blocks.Add(p);

        var md = MarkdownWriter.ToMarkdown(doc);
        Assert.Contains(@"\^", md);
        // 이스케이프된 출력이 다시 파싱돼도 원문 텍스트가 보존돼야 한다.
        var reread = MarkdownReader.FromMarkdown(md);
        Assert.Contains("2^10 = 1024", reread.EnumerateParagraphs().Single().GetPlainText());
    }

    [Fact]
    public void Writer_CodeFence_LongerWhenCodeContainsBackticks()
    {
        // 코드 내에 ``` 이 있으면 기본 3-backtick 펜스가 충돌 — 4-backtick 이상으로 자동 확장.
        var doc = new PolyDonkyument();
        var sec = new Section(); doc.Sections.Add(sec);
        var p = new Paragraph { Style = { CodeLanguage = "md" } };
        p.AddText("use ```code``` here");  // 3-backtick 이 코드 내에 있음.
        sec.Blocks.Add(p);

        var md = MarkdownWriter.ToMarkdown(doc);
        // 펜스는 4개 이상 백틱이어야 한다.
        Assert.Contains("````", md);
        // 파싱해도 코드 내용이 보존돼야 한다.
        var reread = MarkdownReader.FromMarkdown(md);
        var rp     = reread.EnumerateParagraphs().Single();
        Assert.Equal("md", rp.Style.CodeLanguage);
        Assert.Contains("```code```", rp.GetPlainText());
    }

    // ── Writer: Sub/Superscript ─────────────────────────────────────

    [Fact]
    public void Writer_SubAndSuperscript_Emitted()
    {
        var doc = new PolyDonkyument();
        var sec = new Section(); doc.Sections.Add(sec);
        var p = new Paragraph();
        p.AddText("H", new RunStyle());
        p.Runs.Add(new Run { Text = "2", Style = new RunStyle { Subscript = true } });
        p.AddText("O of X", new RunStyle());
        p.Runs.Add(new Run { Text = "2", Style = new RunStyle { Superscript = true } });
        sec.Blocks.Add(p);

        var md = MarkdownWriter.ToMarkdown(doc);
        Assert.Contains("~2~", md);  // subscript
        Assert.Contains("^2^", md);  // superscript
    }

    [Fact]
    public void RoundTrip_SubAndSuperscript_Preserved()
    {
        var doc = new PolyDonkyument();
        var sec = new Section(); doc.Sections.Add(sec);
        var p = new Paragraph();
        p.Runs.Add(new Run { Text = "2", Style = new RunStyle { Subscript   = true } });
        p.Runs.Add(new Run { Text = "n", Style = new RunStyle { Superscript = true } });
        sec.Blocks.Add(p);

        var md      = MarkdownWriter.ToMarkdown(doc);
        var reread  = MarkdownReader.FromMarkdown(md);
        var runs    = reread.EnumerateParagraphs().Single().Runs;
        Assert.Contains(runs, r => r.Style.Subscript   && r.Text == "2");
        Assert.Contains(runs, r => r.Style.Superscript && r.Text == "n");
    }

    // ── Writer: 이미지 임베디드 data URI ────────────────────────────

    [Fact]
    public void Writer_Image_EmitsDataUri_WhenNoResourcePath()
    {
        var doc = new PolyDonkyument();
        var sec = new Section(); doc.Sections.Add(sec);
        var img = new ImageBlock
        {
            Description = "test",
            MediaType   = "image/png",
            Data        = new byte[] { 0x89, 0x50, 0x4E, 0x47 },  // PNG 매직 바이트
        };
        sec.Blocks.Add(img);

        var md = MarkdownWriter.ToMarkdown(doc);
        Assert.Contains("![test](data:image/png;base64,", md);
        Assert.Contains("iVBORw==", md);  // 위 4바이트의 base64
    }

    // ── Reader/Writer: 수식 round-trip ───────────────────────────────

    [Fact]
    public void RoundTrip_InlineAndDisplayMath_Preserved()
    {
        const string source = "인라인 $a^2$ 및\n\n$$\nE=mc^2\n$$\n";

        var doc      = MarkdownReader.FromMarkdown(source);
        var rendered = MarkdownWriter.ToMarkdown(doc);
        var reread   = MarkdownReader.FromMarkdown(rendered);

        var ps = reread.EnumerateParagraphs().ToList();
        Assert.Contains(ps[0].Runs, r => r.LatexSource == "a^2" && !r.IsDisplayEquation);
        Assert.Contains(ps[1].Runs, r => r.LatexSource!.Contains("E=mc^2") && r.IsDisplayEquation);
    }

    // ── Writer: 중첩 리스트 들여쓰기 ───────────────────────────────

    [Fact]
    public void RoundTrip_NestedBulletList_Preserved()
    {
        const string source =
            "- 상위 A\n" +
            "  - 하위 A1\n" +
            "  - 하위 A2\n" +
            "- 상위 B\n";

        var doc     = MarkdownReader.FromMarkdown(source);
        var md      = MarkdownWriter.ToMarkdown(doc);
        var reread  = MarkdownReader.FromMarkdown(md);
        var ps      = reread.EnumerateParagraphs().ToList();

        Assert.Equal(4, ps.Count);
        Assert.Equal(0, ps[0].Style.ListMarker!.Level);  // 상위
        Assert.Equal(1, ps[1].Style.ListMarker!.Level);  // 하위
        Assert.Equal(1, ps[2].Style.ListMarker!.Level);  // 하위
        Assert.Equal(0, ps[3].Style.ListMarker!.Level);  // 상위
    }

    [Fact]
    public void RoundTrip_OrderedListStartNumber_Preserved()
    {
        const string source = "5. 다섯째\n6. 여섯째\n7. 일곱째\n";

        var doc    = MarkdownReader.FromMarkdown(source);
        var md     = MarkdownWriter.ToMarkdown(doc);
        var reread = MarkdownReader.FromMarkdown(md);
        var ps     = reread.EnumerateParagraphs().ToList();

        Assert.Equal(3, ps.Count);
        Assert.Equal(5, ps[0].Style.ListMarker!.OrderedNumber);
        Assert.Equal(6, ps[1].Style.ListMarker!.OrderedNumber);
        Assert.Equal(7, ps[2].Style.ListMarker!.OrderedNumber);
    }

    [Fact]
    public void RoundTrip_GfmFeatures_Preserved()
    {
        const string source =
            "# 제목\n\n" +
            "본문 *기울임* **굵게** ~~취소~~ [링크](https://x)\n\n" +
            "- [x] 끝\n" +
            "- [ ] 미완\n\n" +
            "> 인용\n\n" +
            "---\n\n" +
            "```python\nprint('hi')\n```\n\n" +
            "| A | B |\n| :--- | ---: |\n| 1 | 2 |\n";

        var doc      = MarkdownReader.FromMarkdown(source);
        var rendered = MarkdownWriter.ToMarkdown(doc);
        var roundTripped = MarkdownReader.FromMarkdown(rendered);

        // 핵심 구조가 라운드트립에서 보존되는지 확인.
        Assert.Equal(OutlineLevel.H1, roundTripped.EnumerateParagraphs().First().Style.Outline);
        Assert.Contains(roundTripped.EnumerateParagraphs(), p => p.Style.IsThematicBreak);
        Assert.Contains(roundTripped.EnumerateParagraphs(), p => p.Style.CodeLanguage == "python");
        Assert.Contains(roundTripped.EnumerateParagraphs(), p => p.Style.ListMarker?.Checked == true);
        Assert.Contains(roundTripped.EnumerateParagraphs(), p => p.Style.ListMarker?.Checked == false);
        Assert.Contains(roundTripped.EnumerateParagraphs(), p => p.Style.QuoteLevel >= 1);
        Assert.Single(roundTripped.Sections[0].Blocks.OfType<PolyDonky.Core.Table>());
    }
}

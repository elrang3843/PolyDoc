using PolyDonky.Core;
using PdXmlReader = PolyDonky.Codecs.Xml.XmlReader;
using PdXmlWriter = PolyDonky.Codecs.Xml.XmlWriter;

namespace PolyDonky.Codecs.Xml.Tests;

public class XmlTests
{
    // ── Reader: XHTML ──────────────────────────────────────────────

    [Fact]
    public void Reader_XhtmlWithDeclarationParsesAsHtml()
    {
        const string source =
            "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
            "<!DOCTYPE html>" +
            "<html xmlns=\"http://www.w3.org/1999/xhtml\"><body>" +
            "<h1>제목</h1><p>본문 <strong>굵게</strong></p>" +
            "</body></html>";

        var doc = PdXmlReader.FromXml(source);
        var ps  = doc.EnumerateParagraphs().ToList();
        Assert.Equal(OutlineLevel.H1, ps[0].Style.Outline);
        Assert.Contains(ps[1].Runs, r => r.Style.Bold && r.Text == "굵게");
    }

    [Fact]
    public void Reader_FragmentXhtmlIsAccepted()
    {
        var doc = PdXmlReader.FromXml("<html><body><p>fragment</p></body></html>");
        var p = doc.EnumerateParagraphs().Single();
        Assert.Equal("fragment", p.GetPlainText());
    }

    // ── Reader: 일반 XML (DocBook 등 비-HTML) ──────────────────────

    [Fact]
    public void Reader_GenericXml_FlattensToParagraphs()
    {
        // DocBook 스타일 — HTML 의미 없는 임의 XML.
        const string source =
            "<?xml version=\"1.0\"?>" +
            "<article><title>주제</title>" +
            "<para>첫째 단락.</para>" +
            "<para>둘째 단락.</para>" +
            "</article>";

        var doc = PdXmlReader.FromXml(source);
        var ps  = doc.EnumerateParagraphs().ToList();
        Assert.Contains(ps, p => p.GetPlainText().Contains("첫째"));
        Assert.Contains(ps, p => p.GetPlainText().Contains("둘째"));
    }

    [Fact]
    public void Reader_RejectsExternalDtd()
    {
        // XXE 차단 — DTD 처리는 거부되어 예외 또는 빈 본문이어야 한다.
        const string evil =
            "<?xml version=\"1.0\"?>" +
            "<!DOCTYPE foo [<!ENTITY xxe SYSTEM \"file:///etc/passwd\">]>" +
            "<doc>&xxe;</doc>";

        // 보안: XXE 가 확장되지 않아야 한다 — 예외 또는 무해한 결과.
        var threw = false;
        try { _ = PdXmlReader.FromXml(evil); } catch { threw = true; }

        if (!threw)
        {
            // 예외가 안 났다면 외부 엔티티가 확장되지 않았는지만 확인.
            var doc = PdXmlReader.FromXml(evil);
            var text = string.Concat(doc.EnumerateParagraphs().Select(p => p.GetPlainText()));
            Assert.DoesNotContain("root:", text);
        }
    }

    // ── Writer (XHTML5 출력) ────────────────────────────────────────

    [Fact]
    public void Writer_FullDocumentEmitsXmlDeclarationAndDoctype()
    {
        var doc = new PolyDonkyument();
        var sec = new Section(); doc.Sections.Add(sec);
        sec.Blocks.Add(Paragraph.Of("hi"));

        var xml = PdXmlWriter.ToXml(doc);
        Assert.StartsWith("<?xml version=\"1.0\" encoding=\"utf-8\"?>", xml);
        Assert.Contains("<!DOCTYPE html>", xml);
        Assert.Contains("xmlns=\"http://www.w3.org/1999/xhtml\"", xml);
    }

    [Fact]
    public void Writer_VoidElementsAreSelfClosing()
    {
        var doc = new PolyDonkyument();
        var sec = new Section(); doc.Sections.Add(sec);
        sec.Blocks.Add(new Paragraph { Style = { IsThematicBreak = true } });
        sec.Blocks.Add(new ImageBlock { ResourcePath = "x.png", Description = "대체" });

        var xml = PdXmlWriter.ToXml(doc, fullDocument: false);
        Assert.Contains("<hr/>", xml);
        Assert.Contains("/>", xml); // <img ... />
        // 자체 닫지 않는 형태 (예: <hr> 또는 <img ...>) 가 없어야 한다.
        Assert.DoesNotContain("<hr>",  xml);
    }

    [Fact]
    public void Writer_MetaTagsSelfClose()
    {
        var doc = new PolyDonkyument();
        doc.Sections.Add(new Section());
        var xml = PdXmlWriter.ToXml(doc);
        Assert.Contains("<meta charset=\"utf-8\"/>", xml);
    }

    [Fact]
    public void Writer_TaskCheckboxIsSelfClosing()
    {
        var doc = new PolyDonkyument();
        var sec = new Section(); doc.Sections.Add(sec);
        var p = new Paragraph();
        p.Style.ListMarker = new ListMarker { Kind = ListKind.Bullet, Checked = true };
        p.AddText("done");
        sec.Blocks.Add(p);

        var xml = PdXmlWriter.ToXml(doc, fullDocument: false);
        Assert.Contains("<input type=\"checkbox\" disabled=\"disabled\" checked=\"checked\"/>", xml);
    }

    [Fact]
    public void Writer_BrAndHrInTableCellSelfClose()
    {
        var doc = new PolyDonkyument();
        var sec = new Section(); doc.Sections.Add(sec);
        var t = new Table();
        t.Columns.Add(new TableColumn());
        var row = new TableRow();
        var cell = new TableCell();
        cell.Blocks.Add(Paragraph.Of("first"));
        cell.Blocks.Add(Paragraph.Of("second"));
        row.Cells.Add(cell);
        t.Rows.Add(row);
        sec.Blocks.Add(t);

        var xml = PdXmlWriter.ToXml(doc, fullDocument: false);
        Assert.Contains("first<br/>second", xml);
    }

    [Fact]
    public void Writer_EscapesAttributeQuotesAndApostrophes()
    {
        var doc = new PolyDonkyument();
        var sec = new Section(); doc.Sections.Add(sec);
        sec.Blocks.Add(new ImageBlock { ResourcePath = "a\"'b.png", Description = "alt\"'" });

        var xml = PdXmlWriter.ToXml(doc, fullDocument: false);
        Assert.Contains("&quot;", xml);
        Assert.Contains("&apos;", xml);
    }

    [Fact]
    public void Writer_AmpersandLessGreaterThan_AreEscaped()
    {
        var doc = new PolyDonkyument();
        var sec = new Section(); doc.Sections.Add(sec);
        var p = new Paragraph();
        p.AddText("a < b & c > d");
        sec.Blocks.Add(p);

        var xml = PdXmlWriter.ToXml(doc, fullDocument: false);
        Assert.Contains("a &lt; b &amp; c &gt; d", xml);
    }

    [Fact]
    public void Writer_OutputParsesAsValidXml()
    {
        // 작성기 출력은 XML 파서로 다시 파싱 가능해야 (well-formed) 한다.
        var doc = new PolyDonkyument();
        var sec = new Section(); doc.Sections.Add(sec);
        var h = new Paragraph { Style = { Outline = OutlineLevel.H1 } }; h.AddText("<제목 & 부제목>");
        sec.Blocks.Add(h);
        sec.Blocks.Add(Paragraph.Of("본문"));
        sec.Blocks.Add(new Paragraph { Style = { IsThematicBreak = true } });
        sec.Blocks.Add(new ImageBlock { ResourcePath = "x.png", Description = "대체" });

        var xml  = PdXmlWriter.ToXml(doc);
        var settings = new System.Xml.XmlReaderSettings
        {
            DtdProcessing  = System.Xml.DtdProcessing.Parse,  // <!DOCTYPE html> 허용
            XmlResolver    = null,                            // 외부 참조 차단
            ValidationType = System.Xml.ValidationType.None,
        };

        using var sr = new StringReader(xml);
        using var xr = System.Xml.XmlReader.Create(sr, settings);
        // 단순히 끝까지 읽기 — 형식 오류면 예외 throw.
        while (xr.Read()) { }
    }

    // ── Round-trip ─────────────────────────────────────────────────

    [Fact]
    public void RoundTrip_PreservesStructure()
    {
        const string source =
            "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
            "<!DOCTYPE html>" +
            "<html xmlns=\"http://www.w3.org/1999/xhtml\"><body>" +
            "<h1>제목</h1>" +
            "<p>본문 <strong>굵게</strong> <a href=\"https://x\">링크</a></p>" +
            "<ul><li>항목</li></ul>" +
            "<blockquote><p>인용</p></blockquote>" +
            "<pre><code class=\"language-py\">print(1)</code></pre>" +
            "<hr/>" +
            "<table><thead><tr><th>A</th></tr></thead><tbody><tr><td>1</td></tr></tbody></table>" +
            "</body></html>";

        var doc      = PdXmlReader.FromXml(source);
        var rendered = PdXmlWriter.ToXml(doc, fullDocument: false);
        var reread   = PdXmlReader.FromXml(rendered);

        Assert.Equal(OutlineLevel.H1, reread.EnumerateParagraphs().First().Style.Outline);
        Assert.Contains(reread.EnumerateParagraphs(), p => p.Style.QuoteLevel >= 1);
        Assert.Contains(reread.EnumerateParagraphs(), p => p.Style.IsThematicBreak);
        Assert.Contains(reread.EnumerateParagraphs(), p => p.Style.CodeLanguage == "py");
        Assert.Contains(reread.EnumerateParagraphs(), p => p.Style.ListMarker?.Kind == ListKind.Bullet);
        Assert.Single(reread.Sections[0].Blocks.OfType<Table>());
        Assert.Contains(reread.EnumerateParagraphs().SelectMany(p => p.Runs),
            r => r.Url == "https://x");
    }

    [Fact]
    public void RoundTrip_IsXmlWellFormed()
    {
        var doc = new PolyDonkyument();
        var sec = new Section(); doc.Sections.Add(sec);
        var h = new Paragraph { Style = { Outline = OutlineLevel.H2 } }; h.AddText("hi");
        sec.Blocks.Add(h);
        sec.Blocks.Add(Paragraph.Of("body"));

        var xml = PdXmlWriter.ToXml(doc);

        // System.Xml 로 well-formed 검증.
        var settings = new System.Xml.XmlReaderSettings
        {
            DtdProcessing = System.Xml.DtdProcessing.Parse,
            XmlResolver   = null,
        };
        using var sr = new StringReader(xml);
        using var xr = System.Xml.XmlReader.Create(sr, settings);
        while (xr.Read()) { }  // 형식 오류 시 예외.
    }
}

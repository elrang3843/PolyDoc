using System.Text;
using PolyDoc.Codecs.Docx;
using PolyDoc.Codecs.Hwpx;
using PolyDoc.Codecs.Markdown;
using PolyDoc.Codecs.Text;
using PolyDoc.Core;
using PolyDoc.Iwpf;

// PolyDoc.SmokeTest — BCL 만으로 동작하는 자체 스모크 러너.
// xUnit/NuGet 차단 환경에서도 핵심 라운드트립을 검증하기 위한 임시 도구.
// 정식 테스트는 tests/PolyDoc.*.Tests 의 xUnit 프로젝트에서 수행한다.

var harness = new SmokeHarness();

harness.Run("Plain text round-trip", PlainTextRoundTrip);
harness.Run("Markdown round-trip (headers + emphasis + lists)", MarkdownRoundTrip);
harness.Run("IWPF round-trip (manifest + integrity)", IwpfRoundTrip);
harness.Run("IWPF tampering detection", IwpfTamperingDetection);
harness.Run("DOCX round-trip (headings + emphasis)", DocxRoundTrip);
harness.Run("HWPX round-trip (KS X 6101 self interop)", HwpxRoundTrip);

return harness.Finish();


static void PlainTextRoundTrip()
{
    const string sample = "첫 번째 줄\n두 번째 줄입니다.\n세 번째 — 한글 조판 테스트.";
    var doc = PlainTextReader.FromText(sample);
    SmokeHarness.Equal(3, doc.Sections[0].Blocks.Count, "block count");

    var roundTripped = PlainTextWriter.ToText(doc);
    SmokeHarness.Equal(sample, roundTripped, "plain text body");
}

static void MarkdownRoundTrip()
{
    const string source =
        "# 제목 1\n" +
        "\n" +
        "본문 단락은 **굵게** 와 *기울임* 을 함께 가진다.\n" +
        "\n" +
        "## 부제목\n" +
        "\n" +
        "- 첫 번째 항목\n" +
        "- 두 번째 항목\n" +
        "\n" +
        "1. 순서 항목 A\n" +
        "2. 순서 항목 B\n";

    var doc = MarkdownReader.FromMarkdown(source);

    var paragraphs = doc.EnumerateParagraphs().ToList();
    SmokeHarness.Equal(7, paragraphs.Count, "paragraph count");
    SmokeHarness.Equal(OutlineLevel.H1, paragraphs[0].Style.Outline, "first outline level");
    SmokeHarness.Equal(OutlineLevel.H2, paragraphs[2].Style.Outline, "third outline level");
    SmokeHarness.Equal(ListKind.Bullet, paragraphs[3].Style.ListMarker!.Kind, "bullet list kind");
    SmokeHarness.Equal(ListKind.OrderedDecimal, paragraphs[5].Style.ListMarker!.Kind, "ordered list kind");

    // bold/italic Run 분리 확인
    var bodyRuns = paragraphs[1].Runs;
    var hasBold = bodyRuns.Any(r => r.Style.Bold && r.Text == "굵게");
    var hasItalic = bodyRuns.Any(r => r.Style.Italic && r.Text == "기울임");
    SmokeHarness.True(hasBold, "bold run present");
    SmokeHarness.True(hasItalic, "italic run present");

    // 라운드트립 후 다시 파싱했을 때 헤더/리스트 구조가 보존되는지
    var rendered = MarkdownWriter.ToMarkdown(doc);
    var reparsed = MarkdownReader.FromMarkdown(rendered);
    var reparsedParagraphs = reparsed.EnumerateParagraphs().ToList();
    SmokeHarness.Equal(paragraphs.Count, reparsedParagraphs.Count, "reparsed paragraph count");
    SmokeHarness.Equal(OutlineLevel.H1, reparsedParagraphs[0].Style.Outline, "reparsed first outline level");
    SmokeHarness.Equal(ListKind.Bullet, reparsedParagraphs[3].Style.ListMarker!.Kind, "reparsed bullet kind");
}

static void IwpfRoundTrip()
{
    var doc = new PolyDocument();
    doc.Metadata.Title = "스모크 테스트 문서";
    doc.Metadata.Author = "Noh JinMoon";

    var section = new Section();
    doc.Sections.Add(section);

    var heading = new Paragraph();
    heading.Style.Outline = OutlineLevel.H1;
    heading.AddText("PolyDoc IWPF 라운드트립");
    section.Blocks.Add(heading);

    var body = new Paragraph();
    body.AddText("이것은 IWPF 패키지의 ", new RunStyle());
    body.AddText("핵심", new RunStyle { Bold = true });
    body.AddText(" 라운드트립을 검증합니다.", new RunStyle());
    section.Blocks.Add(body);

    using var ms = new MemoryStream();
    new IwpfWriter().Write(doc, ms);
    SmokeHarness.True(ms.Length > 100, $"package size > 100 bytes (got {ms.Length})");

    ms.Position = 0;
    var read = new IwpfReader().Read(ms);

    SmokeHarness.Equal("스모크 테스트 문서", read.Metadata.Title!, "metadata.title");
    SmokeHarness.Equal("Noh JinMoon", read.Metadata.Author!, "metadata.author");

    var roundParagraphs = read.EnumerateParagraphs().ToList();
    SmokeHarness.Equal(2, roundParagraphs.Count, "paragraph count after read");
    SmokeHarness.Equal(OutlineLevel.H1, roundParagraphs[0].Style.Outline, "heading outline preserved");
    SmokeHarness.Equal("PolyDoc IWPF 라운드트립", roundParagraphs[0].GetPlainText(), "heading text");
    SmokeHarness.True(roundParagraphs[1].Runs.Any(r => r.Style.Bold && r.Text == "핵심"), "bold run preserved");
}

static void IwpfTamperingDetection()
{
    var doc = new PolyDocument();
    doc.Sections.Add(new Section());
    doc.Sections[0].Blocks.Add(Paragraph.Of("위변조 검출 테스트"));

    using var ms = new MemoryStream();
    new IwpfWriter().Write(doc, ms);

    // ZIP 내부의 document.json 페이로드를 추출/변조해 다시 ZIP 으로 묶는다.
    var tampered = TamperDocumentJson(ms.ToArray());

    using var ts = new MemoryStream(tampered);
    var caught = false;
    try
    {
        new IwpfReader().Read(ts);
    }
    catch (InvalidDataException)
    {
        caught = true;
    }
    SmokeHarness.True(caught, "tampered package was rejected");
}

static void DocxRoundTrip()
{
    var doc = new PolyDocument();
    doc.Metadata.Title = "DOCX 스모크";
    doc.Metadata.Author = "Noh JinMoon";
    var section = new Section();
    doc.Sections.Add(section);

    var heading = new Paragraph { Style = { Outline = OutlineLevel.H1 } };
    heading.AddText("DOCX 1급 시민");
    section.Blocks.Add(heading);

    var body = new Paragraph();
    body.AddText("OpenXml 기반 ");
    body.AddText("DOCX", new RunStyle { Bold = true });
    body.AddText(" 라운드트립을 검증합니다.");
    section.Blocks.Add(body);

    using var ms = new MemoryStream();
    new DocxWriter().Write(doc, ms);
    SmokeHarness.True(ms.Length > 1000, $"DOCX size > 1 KB (got {ms.Length})");

    ms.Position = 0;
    var read = new DocxReader().Read(ms);

    SmokeHarness.Equal("DOCX 스모크", read.Metadata.Title!, "DOCX metadata.title");
    SmokeHarness.Equal("Noh JinMoon", read.Metadata.Author!, "DOCX metadata.author");

    var paragraphs = read.EnumerateParagraphs().ToList();
    SmokeHarness.Equal(2, paragraphs.Count, "DOCX paragraph count");
    SmokeHarness.Equal(OutlineLevel.H1, paragraphs[0].Style.Outline, "DOCX heading outline");
    SmokeHarness.True(paragraphs[1].Runs.Any(r => r.Style.Bold && r.Text == "DOCX"), "DOCX bold run preserved");
}

static void HwpxRoundTrip()
{
    var doc = new PolyDocument();
    doc.Metadata.Title = "HWPX 스모크";
    doc.Metadata.Author = "Noh JinMoon";
    var section = new Section();
    doc.Sections.Add(section);

    var heading = new Paragraph { Style = { Outline = OutlineLevel.H2 } };
    heading.AddText("KS X 6101 자체 라운드트립");
    section.Blocks.Add(heading);

    var body = new Paragraph { Style = { Alignment = Alignment.Center } };
    body.AddText("한글 ");
    body.AddText("굵게", new RunStyle { Bold = true });
    body.AddText(" 가운데 정렬.");
    section.Blocks.Add(body);

    using var ms = new MemoryStream();
    new HwpxWriter().Write(doc, ms);
    SmokeHarness.True(ms.Length > 500, $"HWPX size > 500 bytes (got {ms.Length})");

    ms.Position = 0;
    var read = new HwpxReader().Read(ms);

    SmokeHarness.Equal("HWPX 스모크", read.Metadata.Title!, "HWPX metadata.title");
    SmokeHarness.Equal("Noh JinMoon", read.Metadata.Author!, "HWPX metadata.author");

    var paragraphs = read.EnumerateParagraphs().ToList();
    SmokeHarness.Equal(2, paragraphs.Count, "HWPX paragraph count");
    SmokeHarness.Equal(OutlineLevel.H2, paragraphs[0].Style.Outline, "HWPX heading outline");
    SmokeHarness.Equal(Alignment.Center, paragraphs[1].Style.Alignment, "HWPX center alignment");
    SmokeHarness.True(paragraphs[1].Runs.Any(r => r.Style.Bold && r.Text == "굵게"), "HWPX bold run preserved");
}

static byte[] TamperDocumentJson(byte[] original)
{
    using var input = new MemoryStream(original);
    using var output = new MemoryStream();
    using (var inputZip = new System.IO.Compression.ZipArchive(input, System.IO.Compression.ZipArchiveMode.Read, leaveOpen: true))
    using (var outputZip = new System.IO.Compression.ZipArchive(output, System.IO.Compression.ZipArchiveMode.Create, leaveOpen: true))
    {
        foreach (var entry in inputZip.Entries)
        {
            byte[] payload;
            using (var es = entry.Open())
            using (var ms = new MemoryStream())
            {
                es.CopyTo(ms);
                payload = ms.ToArray();
            }

            if (entry.FullName == "content/document.json")
            {
                // 한 바이트 뒤집어 해시를 깬다.
                payload[^2] ^= 0x55;
            }

            var newEntry = outputZip.CreateEntry(entry.FullName, System.IO.Compression.CompressionLevel.Optimal);
            using var ws = newEntry.Open();
            ws.Write(payload, 0, payload.Length);
        }
    }
    return output.ToArray();
}


internal sealed class SmokeHarness
{
    private int _passed;
    private int _failed;
    private readonly List<string> _failures = new();

    public void Run(string name, Action body)
    {
        try
        {
            body();
            _passed++;
            Console.WriteLine($"  PASS  {name}");
        }
        catch (Exception ex)
        {
            _failed++;
            _failures.Add($"{name}: {ex.Message}");
            Console.WriteLine($"  FAIL  {name}");
            Console.WriteLine($"        {ex.GetType().Name}: {ex.Message}");
        }
    }

    public int Finish()
    {
        Console.WriteLine();
        Console.WriteLine(new string('-', 60));
        Console.WriteLine($"PolyDoc.SmokeTest: {_passed} passed, {_failed} failed");
        if (_failed > 0)
        {
            foreach (var f in _failures)
            {
                Console.WriteLine($"  - {f}");
            }
            return 1;
        }
        return 0;
    }

    public static void Equal<T>(T expected, T actual, string what)
    {
        if (!EqualityComparer<T>.Default.Equals(expected, actual))
        {
            throw new InvalidOperationException($"{what}: expected={expected}, actual={actual}");
        }
    }

    public static void True(bool condition, string what)
    {
        if (!condition)
        {
            throw new InvalidOperationException($"assertion failed: {what}");
        }
    }
}

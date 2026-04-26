using System.IO.Compression;
using PolyDoc.Core;
using PolyDoc.Iwpf;

namespace PolyDoc.Iwpf.Tests;

public class IwpfRoundTripTests
{
    [Fact]
    public void RoundTrip_PreservesMetadataAndBody()
    {
        var doc = new PolyDocument();
        doc.Metadata.Title = "라운드트립";
        doc.Metadata.Author = "Noh JinMoon";
        var section = new Section();
        doc.Sections.Add(section);
        var heading = new Paragraph { Style = { Outline = OutlineLevel.H1 } };
        heading.AddText("제목");
        section.Blocks.Add(heading);
        var body = new Paragraph();
        body.AddText("본문 ");
        body.AddText("강조", new RunStyle { Bold = true });
        body.AddText(" 끝.");
        section.Blocks.Add(body);

        var bytes = WriteToBytes(doc);
        var read = ReadFromBytes(bytes);

        Assert.Equal("라운드트립", read.Metadata.Title);
        Assert.Equal("Noh JinMoon", read.Metadata.Author);
        var paragraphs = read.EnumerateParagraphs().ToList();
        Assert.Equal(2, paragraphs.Count);
        Assert.Equal(OutlineLevel.H1, paragraphs[0].Style.Outline);
        Assert.Equal("제목", paragraphs[0].GetPlainText());
        Assert.Contains(paragraphs[1].Runs, r => r.Style.Bold && r.Text == "강조");
    }

    [Fact]
    public void Package_ContainsRequiredParts()
    {
        var doc = PolyDocument.Empty();
        var bytes = WriteToBytes(doc);

        using var ms = new MemoryStream(bytes);
        using var archive = new ZipArchive(ms, ZipArchiveMode.Read);

        Assert.NotNull(archive.GetEntry("manifest.json"));
        Assert.NotNull(archive.GetEntry("content/document.json"));
        Assert.NotNull(archive.GetEntry("content/styles.json"));
    }

    [Fact]
    public void Read_ThrowsOnTamperedDocumentJson()
    {
        var doc = PolyDocument.Empty();
        doc.Sections[0].Blocks.Add(Paragraph.Of("원본 데이터"));
        var bytes = WriteToBytes(doc);

        var tampered = TamperPart(bytes, "content/document.json");

        using var ms = new MemoryStream(tampered);
        Assert.Throws<InvalidDataException>(() => new IwpfReader().Read(ms));
    }

    [Fact]
    public void Read_RejectsUnknownPackageType()
    {
        // 매니페스트의 packageType 을 다른 값으로 바꿔 위장한다.
        var doc = PolyDocument.Empty();
        var bytes = WriteToBytes(doc);
        var faked = ReplaceInManifest(bytes, "polydoc.iwpf", "vendor.unknown");

        using var ms = new MemoryStream(faked);
        Assert.Throws<InvalidDataException>(() => new IwpfReader().Read(ms));
    }

    [Fact]
    public void RoundTrip_TableStructure()
    {
        var doc = new PolyDocument();
        var section = new Section();
        doc.Sections.Add(section);

        var table = new Table();
        table.Columns.Add(new TableColumn { WidthMm = 30 });
        table.Columns.Add(new TableColumn { WidthMm = 50 });

        var row = new TableRow();
        row.Cells.Add(new TableCell { Blocks = { Paragraph.Of("이름") } });
        row.Cells.Add(new TableCell { Blocks = { Paragraph.Of("값"), Paragraph.Of("부가 설명") } });
        table.Rows.Add(row);
        section.Blocks.Add(table);

        var bytes = WriteToBytes(doc);
        var read = ReadFromBytes(bytes);

        var t = read.Sections[0].Blocks.OfType<Table>().Single();
        Assert.Equal(2, t.Columns.Count);
        Assert.Single(t.Rows);
        Assert.Equal(2, t.Rows[0].Cells.Count);
        Assert.Equal(2, t.Rows[0].Cells[1].Blocks.Count);
        Assert.Equal("값", ((Paragraph)t.Rows[0].Cells[1].Blocks[0]).GetPlainText());
        Assert.Equal("부가 설명", ((Paragraph)t.Rows[0].Cells[1].Blocks[1]).GetPlainText());
    }

    [Fact]
    public void RoundTrip_ImageStoredInResourcesImagesDir()
    {
        var image = new ImageBlock
        {
            MediaType = "image/png",
            Data = Enumerable.Range(0, 256).Select(i => (byte)i).ToArray(),
            WidthMm = 30,
            HeightMm = 20,
            Description = "샘플",
        };

        var doc = new PolyDocument();
        var section = new Section();
        section.Blocks.Add(image);
        doc.Sections.Add(section);

        var bytes = WriteToBytes(doc);

        // ZIP 안에 resources/images/ 파트가 만들어져야 한다.
        using var ms = new MemoryStream(bytes);
        using var archive = new ZipArchive(ms, ZipArchiveMode.Read);
        var imageEntries = archive.Entries.Where(e => e.FullName.StartsWith("resources/images/", StringComparison.Ordinal)).ToList();
        Assert.Single(imageEntries);

        // 라운드트립 후 ImageBlock 의 바이트가 동일하고 ResourcePath 가 채워져 있어야 한다.
        var read = ReadFromBytes(bytes);
        var roundTripImage = read.Sections[0].Blocks.OfType<ImageBlock>().Single();
        Assert.Equal(image.Data, roundTripImage.Data);
        Assert.NotNull(roundTripImage.ResourcePath);
        Assert.StartsWith("resources/images/", roundTripImage.ResourcePath!);
        Assert.NotNull(roundTripImage.Sha256);
        Assert.Matches("^[0-9a-f]{64}$", roundTripImage.Sha256!);
    }

    [Fact]
    public void RoundTrip_DedupesIdenticalImages()
    {
        var bytes = new byte[] { 1, 2, 3, 4, 5 };
        var doc = new PolyDocument();
        var section = new Section();
        doc.Sections.Add(section);
        section.Blocks.Add(new ImageBlock { MediaType = "image/png", Data = bytes });
        section.Blocks.Add(new ImageBlock { MediaType = "image/png", Data = (byte[])bytes.Clone() });

        var packageBytes = WriteToBytes(doc);

        using var ms = new MemoryStream(packageBytes);
        using var archive = new ZipArchive(ms, ZipArchiveMode.Read);
        var imageEntries = archive.Entries.Count(e => e.FullName.StartsWith("resources/images/", StringComparison.Ordinal));
        Assert.Equal(1, imageEntries);
    }

    [Fact]
    public void RoundTrip_OpaqueBlockPreservesXml()
    {
        const string xml = "<vendor:custom xmlns:vendor=\"urn:x-vendor\">payload</vendor:custom>";
        var doc = new PolyDocument();
        var section = new Section();
        section.Blocks.Add(new OpaqueBlock
        {
            Format = "docx",
            Kind = "vendor",
            Xml = xml,
            DisplayLabel = "[보존된 vendor]",
        });
        doc.Sections.Add(section);

        var read = ReadFromBytes(WriteToBytes(doc));
        var preserved = read.Sections[0].Blocks.OfType<OpaqueBlock>().Single();
        Assert.Equal("docx", preserved.Format);
        Assert.Equal("vendor", preserved.Kind);
        Assert.Equal(xml, preserved.Xml);
        Assert.Equal(NodeStatus.Opaque, preserved.Status);
    }

    [Fact]
    public void Manifest_RecordsSha256ForEveryPart()
    {
        var doc = PolyDocument.Empty();
        var bytes = WriteToBytes(doc);
        using var ms = new MemoryStream(bytes);
        using var archive = new ZipArchive(ms, ZipArchiveMode.Read);

        var manifestEntry = archive.GetEntry("manifest.json");
        Assert.NotNull(manifestEntry);
        using var stream = manifestEntry!.Open();
        using var memory = new MemoryStream();
        stream.CopyTo(memory);
        var manifest = System.Text.Json.JsonSerializer.Deserialize<IwpfManifest>(memory.ToArray(), JsonDefaults.Options);
        Assert.NotNull(manifest);
        Assert.NotEmpty(manifest!.Parts);
        Assert.All(manifest.Parts.Values, e => Assert.Matches("^[0-9a-f]{64}$", e.Sha256));
    }

    [Fact]
    public void Read_LegacyKindDiscriminator_StillReadable()
    {
        // 옛 빌드(29c09bd)는 discriminator 가 "kind" 였다.
        // 사용자가 그 시기에 저장한 iwpf 가 신 빌드에서도 열려야 한다.
        var doc = PolyDocument.Empty();
        doc.Sections[0].Blocks.Add(Paragraph.Of("legacy"));
        var bytes = WriteToBytes(doc);
        var rewritten = RewriteDocumentJson(bytes,
            json => json.Replace("\"$type\": \"paragraph\"", "\"kind\": \"paragraph\"", StringComparison.Ordinal));

        var read = new IwpfReader { VerifyHashes = false }.Read(new MemoryStream(rewritten));
        var paragraphs = read.EnumerateParagraphs().ToList();
        Assert.Single(paragraphs);
        Assert.Equal("legacy", paragraphs[0].GetPlainText());
    }

    [Fact]
    public void Read_MissingDiscriminator_FallsBackToParagraph()
    {
        // 매우 옛 / 손상된 파일 시뮬레이션 — discriminator 가 아예 없으면 Paragraph 로 해석.
        var doc = PolyDocument.Empty();
        doc.Sections[0].Blocks.Add(Paragraph.Of("no-disc"));
        var bytes = WriteToBytes(doc);
        var rewritten = RewriteDocumentJson(bytes,
            json => json.Replace("\"$type\": \"paragraph\",\n", "", StringComparison.Ordinal)
                       .Replace("\"$type\": \"paragraph\",\r\n", "", StringComparison.Ordinal));

        var read = new IwpfReader { VerifyHashes = false }.Read(new MemoryStream(rewritten));
        var paragraphs = read.EnumerateParagraphs().ToList();
        Assert.Single(paragraphs);
        Assert.Equal("no-disc", paragraphs[0].GetPlainText());
    }

    private static byte[] RewriteDocumentJson(byte[] original, Func<string, string> transform)
    {
        using var input = new MemoryStream(original);
        using var output = new MemoryStream();
        using (var inputZip = new ZipArchive(input, ZipArchiveMode.Read, leaveOpen: true))
        using (var outputZip = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
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
                    var text = System.Text.Encoding.UTF8.GetString(payload);
                    text = transform(text);
                    payload = System.Text.Encoding.UTF8.GetBytes(text);
                }
                var newEntry = outputZip.CreateEntry(entry.FullName, CompressionLevel.Optimal);
                using var ws = newEntry.Open();
                ws.Write(payload, 0, payload.Length);
            }
        }
        return output.ToArray();
    }

    private static byte[] WriteToBytes(PolyDocument doc)
    {
        using var ms = new MemoryStream();
        new IwpfWriter().Write(doc, ms);
        return ms.ToArray();
    }

    private static PolyDocument ReadFromBytes(byte[] bytes)
    {
        using var ms = new MemoryStream(bytes);
        return new IwpfReader().Read(ms);
    }

    private static byte[] TamperPart(byte[] original, string path)
    {
        using var input = new MemoryStream(original);
        using var output = new MemoryStream();
        using (var inputZip = new ZipArchive(input, ZipArchiveMode.Read, leaveOpen: true))
        using (var outputZip = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
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
                if (entry.FullName == path && payload.Length > 1)
                {
                    payload[^2] ^= 0x55;
                }
                var newEntry = outputZip.CreateEntry(entry.FullName, CompressionLevel.Optimal);
                using var ws = newEntry.Open();
                ws.Write(payload, 0, payload.Length);
            }
        }
        return output.ToArray();
    }

    private static byte[] ReplaceInManifest(byte[] original, string find, string replace)
    {
        using var input = new MemoryStream(original);
        using var output = new MemoryStream();
        using (var inputZip = new ZipArchive(input, ZipArchiveMode.Read, leaveOpen: true))
        using (var outputZip = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
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
                if (entry.FullName == "manifest.json")
                {
                    var text = System.Text.Encoding.UTF8.GetString(payload);
                    text = text.Replace(find, replace, StringComparison.Ordinal);
                    payload = System.Text.Encoding.UTF8.GetBytes(text);
                }
                var newEntry = outputZip.CreateEntry(entry.FullName, CompressionLevel.Optimal);
                using var ws = newEntry.Open();
                ws.Write(payload, 0, payload.Length);
            }
        }
        return output.ToArray();
    }
}

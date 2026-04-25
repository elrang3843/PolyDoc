using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using PolyDoc.Core;

namespace PolyDoc.Codecs.Hwpx;

/// <summary>
/// HWPX (KS X 6101) → PolyDocument 리더.
///
/// HwpxWriter 와 같은 charPr/paraPr/style ID 약속을 따른다 (writer 의 doc-comment 참조).
/// 다른 도구(한컴 오피스 등)가 만든 HWPX 의 임의 ID 매핑은 다음 사이클에서 header.xml 의
/// charPr/paraPr 정의를 풀어서 정확하게 회수한다.
/// </summary>
public sealed class HwpxReader : IDocumentReader
{
    public string FormatId => "hwpx";

    private static readonly XNamespace OpfContainer = HwpxNamespaces.OpfContainer;
    private static readonly XNamespace Opf = HwpxNamespaces.OpfPackage;
    private static readonly XNamespace Dc = HwpxNamespaces.DcMetadata;
    private static readonly XNamespace Hh = HwpxNamespaces.Head;
    private static readonly XNamespace Hp = HwpxNamespaces.Paragraph;
    private static readonly XNamespace Hs = HwpxNamespaces.Section;

    public PolyDocument Read(Stream input)
    {
        ArgumentNullException.ThrowIfNull(input);

        Stream zipStream = input;
        MemoryStream? buffered = null;
        if (!input.CanSeek)
        {
            buffered = new MemoryStream();
            input.CopyTo(buffered);
            buffered.Position = 0;
            zipStream = buffered;
        }

        try
        {
            using var archive = new ZipArchive(zipStream, ZipArchiveMode.Read, leaveOpen: true);

            ValidateMimetype(archive);

            var rootHpfPath = ResolveContentHpf(archive);
            var (sectionPaths, metadata) = ReadOpfManifest(archive, rootHpfPath);

            // OPF 가 변종 namespace 또는 빠진 spine 으로 비어 있을 수 있어 ZIP 직접 스캔으로 fallback.
            if (sectionPaths.Count == 0)
            {
                sectionPaths.AddRange(FallbackSectionPaths(archive));
            }

            var document = new PolyDocument { Metadata = metadata };
            int totalParagraphs = 0;
            int totalTextRuns = 0;
            foreach (var path in sectionPaths)
            {
                var section = ReadSection(archive, path);
                document.Sections.Add(section);
                foreach (var block in section.Blocks)
                {
                    if (block is Paragraph p)
                    {
                        totalParagraphs++;
                        foreach (var run in p.Runs)
                        {
                            if (!string.IsNullOrEmpty(run.Text))
                            {
                                totalTextRuns++;
                            }
                        }
                    }
                }
            }
            if (document.Sections.Count == 0)
            {
                document.Sections.Add(new Section());
            }

            // 진단 정보 — MainViewModel 이 "본문 0건" 같은 경고 메시지를 띄울 수 있게 metadata 에 박는다.
            document.Metadata.Custom["hwpx.sectionFilesFound"] = sectionPaths.Count.ToString(System.Globalization.CultureInfo.InvariantCulture);
            document.Metadata.Custom["hwpx.paragraphCount"] = totalParagraphs.ToString(System.Globalization.CultureInfo.InvariantCulture);
            document.Metadata.Custom["hwpx.nonEmptyRunCount"] = totalTextRuns.ToString(System.Globalization.CultureInfo.InvariantCulture);
            if (sectionPaths.Count > 0)
            {
                document.Metadata.Custom["hwpx.firstSectionPath"] = sectionPaths[0];
            }

            return document;
        }
        finally
        {
            buffered?.Dispose();
        }
    }

    /// <summary>
    /// content.hpf 가 못 풀렸거나 spine 이 비었을 때, ZIP 안의 section 파일을 직접 스캔.
    /// 한컴 변종에 대비해 폴더 위치 무관하게 파일명에 "section" 이 들어간 모든 .xml 을 잡는다
    /// (대소문자 무시).
    /// </summary>
    private static IEnumerable<string> FallbackSectionPaths(ZipArchive archive)
    {
        return archive.Entries
            .Select(e => e.FullName)
            .Where(p => p.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                     && System.IO.Path.GetFileName(p).Contains("section", StringComparison.OrdinalIgnoreCase))
            .OrderBy(p => p, StringComparer.OrdinalIgnoreCase);
    }

    private static void ValidateMimetype(ZipArchive archive)
    {
        var entry = archive.GetEntry(HwpxPaths.Mimetype)
            ?? throw new InvalidDataException("HWPX package is missing 'mimetype' entry.");
        using var stream = entry.Open();
        using var reader = new StreamReader(stream, Encoding.UTF8);
        var content = reader.ReadToEnd().Trim();
        if (content != HwpxPaths.MimetypeContent)
        {
            throw new InvalidDataException(
                $"Unexpected HWPX mimetype: '{content}'. Expected '{HwpxPaths.MimetypeContent}'.");
        }
    }

    private static string ResolveContentHpf(ZipArchive archive)
    {
        var container = LoadXml(archive, HwpxPaths.ContainerXml);
        if (container?.Root is null)
        {
            return HwpxPaths.ContentHpf;
        }

        // 한컴 변종/EPUB 호환 모두를 받아내기 위해 LocalName 매칭으로 rootfile 검색.
        var rootfile = container.Root.Descendants()
            .FirstOrDefault(e => e.Name.LocalName == "rootfile");
        var fullPath = rootfile?.Attribute("full-path")?.Value;
        return string.IsNullOrEmpty(fullPath) ? HwpxPaths.ContentHpf : fullPath!;
    }

    private static (List<string> sectionPaths, DocumentMetadata metadata) ReadOpfManifest(ZipArchive archive, string rootHpfPath)
    {
        var doc = LoadXml(archive, rootHpfPath);
        var metadata = new DocumentMetadata();
        var sectionPaths = new List<string>();
        if (doc?.Root is null)
        {
            return (sectionPaths, metadata);
        }

        var packageElem = doc.Root;

        // metadata 는 dc namespace 가 다르거나 default namespace 일 수도 있어 LocalName 매칭.
        var metaContainer = packageElem.Descendants().FirstOrDefault(e => e.Name.LocalName == "metadata");
        if (metaContainer is not null)
        {
            metadata.Title = metaContainer.Descendants().FirstOrDefault(e => e.Name.LocalName == "title")?.Value;
            metadata.Author = metaContainer.Descendants().FirstOrDefault(e => e.Name.LocalName == "creator")?.Value;
            var lang = metaContainer.Descendants().FirstOrDefault(e => e.Name.LocalName == "language")?.Value;
            if (!string.IsNullOrEmpty(lang)) metadata.Language = lang;
        }

        var manifestContainer = packageElem.Descendants().FirstOrDefault(e => e.Name.LocalName == "manifest");
        var manifestItems = manifestContainer?.Descendants()
            .Where(e => e.Name.LocalName == "item")
            .ToDictionary(
                e => e.Attribute("id")?.Value ?? string.Empty,
                e => e.Attribute("href")?.Value ?? string.Empty,
                StringComparer.Ordinal)
            ?? new Dictionary<string, string>(StringComparer.Ordinal);

        var spineContainer = packageElem.Descendants().FirstOrDefault(e => e.Name.LocalName == "spine");
        var spineRefs = spineContainer?.Descendants().Where(e => e.Name.LocalName == "itemref");
        if (spineRefs is not null)
        {
            foreach (var itemref in spineRefs)
            {
                var idref = itemref.Attribute("idref")?.Value;
                if (!string.IsNullOrEmpty(idref) && manifestItems.TryGetValue(idref!, out var href))
                {
                    sectionPaths.Add(CombineRoot(rootHpfPath, href));
                }
            }
        }

        if (sectionPaths.Count == 0)
        {
            foreach (var (id, href) in manifestItems)
            {
                if (id.StartsWith("section", StringComparison.OrdinalIgnoreCase)
                    || (href is { Length: > 0 } && System.IO.Path.GetFileName(href).StartsWith("section", StringComparison.OrdinalIgnoreCase)))
                {
                    sectionPaths.Add(CombineRoot(rootHpfPath, href));
                }
            }
        }

        return (sectionPaths, metadata);
    }

    private static string CombineRoot(string rootHpfPath, string relative)
    {
        var slash = rootHpfPath.LastIndexOf('/');
        var dir = slash >= 0 ? rootHpfPath[..(slash + 1)] : string.Empty;
        return dir + relative;
    }

    private static Section ReadSection(ZipArchive archive, string path)
    {
        var section = new Section();
        var doc = LoadXml(archive, path);
        if (doc?.Root is null)
        {
            return section;
        }

        // 한컴 오피스가 만든 HWPX 는 sec 안에 추가 wrapper(예: subList) 를 둘 수도 있어
        // root 직속이 아닌 깊이의 hp:p 도 찾아야 한다. 표 안의 셀 paragraph 까지 모두 평탄화로
        // 흡수해 사용자가 본문 텍스트는 잃지 않도록 한다 (1차 사이클의 의도적 단순화).
        foreach (var elem in doc.Root.Descendants())
        {
            if (elem.Name.LocalName == "p")
            {
                section.Blocks.Add(ReadParagraph(elem));
            }
        }
        return section;
    }

    private static Paragraph ReadParagraph(XElement wp)
    {
        var paragraph = new Paragraph();
        if (int.TryParse(wp.Attribute("paraPrIDRef")?.Value, out var paraPrId))
        {
            paragraph.Style.Alignment = ParaPrIdToAlignment(paraPrId);
        }
        if (int.TryParse(wp.Attribute("styleIDRef")?.Value, out var styleId)
            && styleId is >= 1 and <= 6)
        {
            paragraph.Style.Outline = (OutlineLevel)styleId;
        }

        // <hp:p> 는 직속 자식으로 <hp:run> 들을 갖는 것이 표준이지만, 한컴 변종에선
        // 중간 wrapper(예: <hp:linesegarray> 다음 위치 등) 를 둘 수 있어 descendants 로 안전 매칭.
        foreach (var elem in wp.Descendants())
        {
            if (elem.Name.LocalName == "run")
            {
                ReadRun(paragraph, elem);
            }
        }

        if (paragraph.Runs.Count == 0)
        {
            paragraph.AddText(string.Empty);
        }
        return paragraph;
    }

    private static void ReadRun(Paragraph paragraph, XElement run)
    {
        var style = new RunStyle();
        if (int.TryParse(run.Attribute("charPrIDRef")?.Value, out var charPrId))
        {
            ApplyCharPrIdToStyle(charPrId, style);
        }

        // <hp:t> 는 보통 직속 자식이지만, 한컴 변종에선 <hp:t> 가 더 깊은 위치에 있을 수도 있고
        // <hp:tab>·<hp:lineBreak> 같은 형제와 섞일 수도 있어 descendants 로 텍스트 노드만 모음.
        var sb = new StringBuilder();
        foreach (var elem in run.Descendants())
        {
            switch (elem.Name.LocalName)
            {
                case "t":
                    sb.Append(elem.Value);
                    break;
                case "tab":
                    sb.Append('\t');
                    break;
                case "lineBreak":
                    sb.Append('\n');
                    break;
            }
        }
        if (sb.Length > 0)
        {
            paragraph.AddText(sb.ToString(), style);
        }
    }

    private static void ApplyCharPrIdToStyle(int id, RunStyle style)
    {
        // HwpxWriter 의 charPr id 매핑과 정확히 대응:
        // 0=plain, 1=bold, 2=italic, 3=bold+italic, 4=underline, 5=strikeout
        switch (id)
        {
            case 1: style.Bold = true; break;
            case 2: style.Italic = true; break;
            case 3: style.Bold = true; style.Italic = true; break;
            case 4: style.Underline = true; break;
            case 5: style.Strikethrough = true; break;
        }
    }

    private static Alignment ParaPrIdToAlignment(int id) => id switch
    {
        1 => Alignment.Center,
        2 => Alignment.Right,
        3 => Alignment.Justify,
        _ => Alignment.Left,
    };

    private static XDocument? LoadXml(ZipArchive archive, string path)
    {
        var entry = archive.GetEntry(path);
        if (entry is null) return null;
        using var stream = entry.Open();
        return XDocument.Load(stream);
    }
}

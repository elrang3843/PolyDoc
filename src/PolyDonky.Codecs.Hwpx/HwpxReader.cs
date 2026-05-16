using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using PolyDonky.Core;

namespace PolyDonky.Codecs.Hwpx;

/// <summary>
/// HWPX (KS X 6101) → PolyDonkyument 리더.
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

    private static readonly HashSet<string> s_shapeLocalNames = new(StringComparer.Ordinal)
    {
        "rect", "line", "ellipse", "arc", "polygon", "textBox", "connector",
    };

    public PolyDonkyument Read(Stream input)
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

            var parseErrors = new List<string>();
            var rootHpfPath = ResolveContentHpf(archive, parseErrors);
            var (sectionPaths, metadata) = ReadOpfManifest(archive, rootHpfPath, parseErrors);

            // OPF 가 변종 namespace 또는 빠진 spine 으로 비어 있을 수 있어 ZIP 직접 스캔으로 fallback.
            if (sectionPaths.Count == 0)
            {
                sectionPaths.AddRange(FallbackSectionPaths(archive));
            }

            // OPF manifest href 가 절대/상대 어느 쪽이든 우리 추정 경로가 실제 ZIP entry 와 안 맞을 수 있다.
            // basename + case-insensitive fallback 으로 정규화한다.
            for (int idx = 0; idx < sectionPaths.Count; idx++)
            {
                sectionPaths[idx] = ResolveEntryPath(archive, sectionPaths[idx]);
            }

            // header.xml — 한컴 / 우리 자체 codec 모두 charPr/paraPr/style 정의를 여기 둔다.
            // 못 찾거나 비어 있으면 빈 컨텍스트로 graceful degradation (paragraph 텍스트는 여전히 회수).
            var headerPath = ResolveEntryPath(archive, HwpxPaths.HeaderXml);
            var headerDoc = LoadXml(archive, headerPath, parseErrors);
            var header = HwpxHeaderReader.Parse(headerDoc);

            var document = new PolyDonkyument { Metadata = metadata };
            var ctx = new ReadContext(archive, header, parseErrors, document);
            int totalParagraphs = 0;
            int totalTextRuns = 0;
            string? firstSectionRoot = null;
            var firstSectionTagCounts = new Dictionary<string, int>(StringComparer.Ordinal);

            for (int i = 0; i < sectionPaths.Count; i++)
            {
                var path = sectionPaths[i];
                var sectionDoc = LoadXml(archive, path, parseErrors);
                var section = ReadSectionFromDoc(sectionDoc, ctx);
                document.Sections.Add(section);

                if (i == 0 && sectionDoc?.Root is { } root)
                {
                    firstSectionRoot = root.Name.LocalName;
                    foreach (var d in root.DescendantsAndSelf())
                    {
                        var name = d.Name.LocalName;
                        firstSectionTagCounts[name] = firstSectionTagCounts.GetValueOrDefault(name, 0) + 1;
                    }
                }

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

    /// <summary>
    /// 추정 경로(예: OPF manifest 의 href + content.hpf 의 dirname 합산)가 실제 ZIP entry 와 안 맞을 때
    /// basename + case-insensitive 매칭으로 보정한다.
    /// </summary>
    private static string ResolveEntryPath(ZipArchive archive, string proposed)
    {
        if (string.IsNullOrEmpty(proposed))
        {
            return proposed;
        }
        if (archive.GetEntry(proposed) is not null)
        {
            return proposed;
        }
        // ZIP 표준은 forward slash 만 쓰지만 일부 변환 도구가 backslash 를 남길 수 있다.
        var normalized = proposed.Replace('\\', '/').TrimStart('/');
        if (archive.GetEntry(normalized) is not null)
        {
            return normalized;
        }
        // case-insensitive 정확 매치
        var ciMatch = archive.Entries
            .FirstOrDefault(e => string.Equals(e.FullName, normalized, StringComparison.OrdinalIgnoreCase));
        if (ciMatch is not null)
        {
            return ciMatch.FullName;
        }
        // basename 매치 (마지막 fallback)
        var basename = System.IO.Path.GetFileName(normalized);
        var byBasename = archive.Entries
            .FirstOrDefault(e => string.Equals(System.IO.Path.GetFileName(e.FullName), basename, StringComparison.OrdinalIgnoreCase));
        return byBasename?.FullName ?? proposed;
    }

    private static void ValidateMimetype(ZipArchive archive)
    {
        var entry = archive.GetEntry(HwpxPaths.Mimetype);
        if (entry is null)
        {
            // 일부 한컴 변종/자가 변환 도구가 mimetype 엔트리를 누락하기도 한다.
            // 본문(Contents/header.xml + section*.xml) 존재 여부로 graceful 통과.
            if (HasCoreHwpxContent(archive)) return;
            throw new InvalidDataException("HWPX package is missing 'mimetype' entry.");
        }
        using var stream = entry.Open();
        // BOM 자동 감지 — 일부 한컴 HWPX 가 mimetype 에 UTF-8 BOM(EF BB BF) 을 붙여 와
        // ASCII reader 가 이를 데이터로 읽어 비교 실패하던 문제 회피.
        using var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
        var content = reader.ReadToEnd().Trim().Trim('﻿');
        if (content == HwpxPaths.MimetypeContent) return;
        // 한컴 변종이 표기를 약간 달리 쓰는 경우(예: x-hwp+zip) 라도 본문이 정상이면 통과.
        if (HasCoreHwpxContent(archive)) return;
        throw new InvalidDataException(
            $"Unexpected HWPX mimetype: '{content}'. Expected '{HwpxPaths.MimetypeContent}'.");
    }

    private static bool HasCoreHwpxContent(ZipArchive archive)
    {
        bool hasHeader  = archive.Entries.Any(e =>
            e.FullName.Equals("Contents/header.xml", StringComparison.OrdinalIgnoreCase));
        bool hasSection = archive.Entries.Any(e =>
            e.FullName.StartsWith("Contents/section", StringComparison.OrdinalIgnoreCase) &&
            e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase));
        return hasHeader && hasSection;
    }

    private static string ResolveContentHpf(ZipArchive archive, List<string>? errors = null)
    {
        var container = LoadXml(archive, HwpxPaths.ContainerXml, errors);
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

    private static (List<string> sectionPaths, DocumentMetadata metadata) ReadOpfManifest(ZipArchive archive, string rootHpfPath, List<string>? errors = null)
    {
        var doc = LoadXml(archive, rootHpfPath, errors);
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
                if (string.IsNullOrEmpty(idref) || !manifestItems.TryGetValue(idref!, out var href))
                {
                    continue;
                }
                // 한컴 hwpx 의 spine 은 header.xml + 본문 sections + 기타(.js·이미지) 를 참조할 수 있어
                // .xml 항목만 section 후보로 채택하되, header.xml 은 구조 파일이므로 제외.
                if (!href.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }
                if (System.IO.Path.GetFileName(href).Equals("header.xml", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }
                sectionPaths.Add(CombineRoot(rootHpfPath, href));
            }
        }

        if (sectionPaths.Count == 0)
        {
            foreach (var (id, href) in manifestItems)
            {
                if (!href.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }
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

    private sealed class ReadContext
    {
        public ReadContext(ZipArchive archive, HwpxHeader header, List<string> parseErrors, PolyDonkyument document)
        {
            Archive = archive;
            Header = header;
            ParseErrors = parseErrors;
            BinDataIndex = BuildBinDataIndex(archive);
            Document = document;
        }

        public ZipArchive Archive { get; }
        public HwpxHeader Header { get; }
        public List<string> ParseErrors { get; }
        public PolyDonkyument Document { get; }
        /// <summary>BinData 안의 파일들을 'basename(확장자 제외)' → FullName 으로 사전화.</summary>
        public Dictionary<string, string> BinDataIndex { get; }

        private static Dictionary<string, string> BuildBinDataIndex(ZipArchive archive)
        {
            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var entry in archive.Entries)
            {
                if (entry.FullName.StartsWith(HwpxPaths.BinDataDir, StringComparison.OrdinalIgnoreCase)
                    && !string.IsNullOrEmpty(System.IO.Path.GetFileName(entry.FullName)))
                {
                    var stem = System.IO.Path.GetFileNameWithoutExtension(entry.FullName);
                    if (!dict.ContainsKey(stem))
                    {
                        dict[stem] = entry.FullName;
                    }
                }
            }
            return dict;
        }
    }

    private static Section ReadSectionFromDoc(XDocument? doc, ReadContext ctx)
    {
        var section = new Section();
        if (doc?.Root is null)
        {
            return section;
        }

        // secPr 에서 페이지 크기·여백 회수 (첫 번째 secPr 만 처리).
        var secPrElem = doc.Root.Descendants().FirstOrDefault(e => e.Name.LocalName == "secPr");
        if (secPrElem is not null)
        {
            ApplySecPr(secPrElem, section);
        }

        // ── 머리말/꼬리말 ─────────────────────────────────────────────────────
        // <hp:header> / <hp:footer> 요소들을 찾아 section.Page에 저장하고,
        // 그 안의 모든 자손 요소들을 insideHeaderFooter로 마킹해 본문 순회에서 제외.
        var insideHeaderFooter = new HashSet<XElement>();
        foreach (var hdrElem in doc.Root.Descendants().Where(e => e.Name.LocalName == "header"))
        {
            foreach (var d in hdrElem.DescendantsAndSelf()) insideHeaderFooter.Add(d);
            ReadHeaderOrFooter(hdrElem, section.Page.Header, ctx);
        }
        foreach (var ftrElem in doc.Root.Descendants().Where(e => e.Name.LocalName == "footer"))
        {
            foreach (var d in ftrElem.DescendantsAndSelf()) insideHeaderFooter.Add(d);
            ReadHeaderOrFooter(ftrElem, section.Page.Footer, ctx);
        }

        // ── 글상자 (drawText) ─────────────────────────────────────────────────
        // 한컴 HWPX 에서 글상자는 <hp:rect>/<hp:textBox> 안의 <hp:drawText> 자식으로 구현.
        // drawText 의 subList 자손을 마킹해 본문 평탄 순회에서 제외.
        var insideTextBox = new HashSet<XElement>();
        var seenTextBoxes = new HashSet<XElement>();
        foreach (var dtElem in doc.Root.Descendants()
            .Where(e => e.Name.LocalName == "drawText" && !insideHeaderFooter.Contains(e)))
        {
            var dtSubList = dtElem.Elements().FirstOrDefault(e => e.Name.LocalName == "subList");
            if (dtSubList is not null)
                foreach (var d in dtSubList.DescendantsAndSelf()) insideTextBox.Add(d);
        }

        // 한컴 hwpx 는 root 안에 wrapper 를 둘 수도 있고 표 안에 셀 paragraph 가 중첩되므로
        // 1) 모든 hp:tbl 안의 paragraph/pic 을 마킹해 평탄 순회에서 제외하고
        // 2) descendants 평탄 순회로 hp:p / hp:tbl / hp:pic 을 본문 block 으로 추출.
        // 표 안 paragraph 는 ReadTable 이 셀에 모은다.
        var insideTable = new HashSet<XElement>();
        foreach (var tbl in doc.Root.Descendants().Where(e => e.Name.LocalName == "tbl"))
        {
            foreach (var d in tbl.Descendants())
            {
                insideTable.Add(d);
            }
        }

        var seenPics    = new HashSet<XElement>();
        var seenShapes  = new HashSet<XElement>();
        foreach (var elem in doc.Root.Descendants())
        {
            if (insideHeaderFooter.Contains(elem)) continue;
            if (insideTextBox.Contains(elem)) continue;
            if (insideTable.Contains(elem)) continue;

            switch (elem.Name.LocalName)
            {
                case "p":
                {
                    // floating 도형·이미지의 앵커 단락(hosting paragraph) 판별:
                    // hp:t 에 실제 텍스트가 없고 hp:tab/hp:lineBreak 도 없는 상태에서
                    // 도형·그림만 들어 있으면 HwpxWriter 의 BuildShapeHostingParagraph /
                    // BuildImageHostingParagraph 가 만든 빈 단락이거나 한컴이 생성한
                    // 앵커 단락이다. 이런 빈 단락을 그대로 추가하면 실제 본문이 아래로 밀린다.
                    var shapeLocal = s_shapeLocalNames;
                    var embeddedShapes = elem.Descendants()
                        .Where(d => shapeLocal.Contains(d.Name.LocalName)
                                 && !insideHeaderFooter.Contains(d)
                                 && !insideTextBox.Contains(d))
                        .ToList();
                    var embeddedPics = elem.Descendants()
                        .Where(d => d.Name.LocalName == "pic"
                                 && !insideHeaderFooter.Contains(d)
                                 && !insideTextBox.Contains(d))
                        .ToList();
                    bool hasRealText = elem.Descendants().Any(d =>
                        !insideHeaderFooter.Contains(d)
                     && !insideTextBox.Contains(d)
                     && ((d.Name.LocalName == "t"   && !string.IsNullOrWhiteSpace(d.Value))
                      || d.Name.LocalName == "tab"
                      || d.Name.LocalName == "lineBreak"));
                    bool isHostingParagraph = (embeddedShapes.Count > 0 || embeddedPics.Count > 0)
                                             && !hasRealText;
                    if (!isHostingParagraph)
                    {
                        section.Blocks.Add(ReadParagraph(elem, ctx));
                    }
                    foreach (var pic in embeddedPics)
                    {
                        if (seenPics.Add(pic) && TryReadPicture(pic, ctx, out var img))
                        {
                            section.Blocks.Add(img);
                        }
                    }
                    foreach (var shape in embeddedShapes)
                    {
                        if (!seenShapes.Add(shape)) continue;
                        // drawText 자식이 있으면 글상자(TextBoxObject), 없으면 일반 도형
                        var drawText = shape.Elements()
                            .FirstOrDefault(e => e.Name.LocalName == "drawText");
                        if (drawText is not null)
                        {
                            if (seenTextBoxes.Add(shape))
                                section.Blocks.Add(ReadTextBox(shape, drawText, ctx));
                        }
                        else
                        {
                            section.Blocks.Add(ReadShape(shape));
                        }
                    }
                    break;
                }
                case "tbl":
                    section.Blocks.Add(ReadTable(elem, ctx));
                    break;
                case "pic":
                    if (seenPics.Add(elem) && TryReadPicture(elem, ctx, out var pictureBlock))
                    {
                        section.Blocks.Add(pictureBlock);
                    }
                    break;
            }
        }
        return section;
    }

    /// <summary>
    /// hp:rect+hp:drawText → TextBoxObject 변환.
    /// shapeElem 에서 크기·위치를, drawText 의 subList 에서 내용을 읽는다.
    /// </summary>
    private static TextBoxObject ReadTextBox(XElement shapeElem, XElement drawText, ReadContext ctx)
    {
        var textBox = new TextBoxObject
        {
            WrapMode = shapeElem.Attribute("textWrap")?.Value?.ToUpperInvariant() switch
            {
                "BEHIND_TEXT"    => ImageWrapMode.BehindText,
                "TOP_AND_BOTTOM" => ImageWrapMode.Inline,
                "FLOAT_LEFT"     => ImageWrapMode.WrapLeft,
                "FLOAT_RIGHT"    => ImageWrapMode.WrapRight,
                _                => ImageWrapMode.InFrontOfText,
            },
        };

        // 크기: curSz → orgSz
        foreach (var szName in new[] { "curSz", "orgSz" })
        {
            var sz = shapeElem.Descendants().FirstOrDefault(e => e.Name.LocalName == szName);
            if (sz is null) continue;
            if (TryParseDouble(sz.Attribute("width")?.Value, out var w) && w > 0)
                textBox.WidthMm = UnitConverter.HwpUnitToMm(w);
            if (TryParseDouble(sz.Attribute("height")?.Value, out var h) && h > 0)
                textBox.HeightMm = UnitConverter.HwpUnitToMm(h);
            if (textBox.WidthMm > 0 && textBox.HeightMm > 0) break;
        }

        // 위치
        var pos = shapeElem.Descendants().FirstOrDefault(e => e.Name.LocalName == "pos");
        if (pos is not null)
        {
            if (TryParseDouble(pos.Attribute("horzOffset")?.Value, out var ox)) textBox.OverlayXMm = UnitConverter.HwpUnitToMm(ox);
            if (TryParseDouble(pos.Attribute("vertOffset")?.Value, out var oy)) textBox.OverlayYMm = UnitConverter.HwpUnitToMm(oy);
        }

        // 내용: drawText 의 subList 안의 hp:p 들
        var subList = drawText.Elements().FirstOrDefault(e => e.Name.LocalName == "subList");
        if (subList is not null)
        {
            foreach (var p in subList.Elements().Where(e => e.Name.LocalName == "p"))
                textBox.Content.Add(ReadParagraph(p, ctx));
        }

        if (textBox.Content.Count == 0)
            textBox.Content.Add(new Paragraph());

        return textBox;
    }

    /// <summary>
    /// hp:header 또는 hp:footer 의 subList 단락들을 HeaderFooterContent 의 Left/Center/Right 슬롯으로 파싱.
    ///
    /// 지원 구조:
    ///  • 1행 3열 표 (HwpxWriter 가 2개 이상 슬롯 시 생성) → 각 셀 → Left/Center/Right
    ///  • 단락들 (정렬로 슬롯 구분: LEFT→Left, CENTER→Center, RIGHT→Right)
    /// </summary>
    private static void ReadHeaderOrFooter(XElement hdrElem, HeaderFooterContent content, ReadContext ctx)
    {
        var subList = hdrElem.Elements().FirstOrDefault(e => e.Name.LocalName == "subList");
        if (subList is null) return;

        // 1×3 표로 저장된 경우 — HwpxWriter 의 BuildHeaderFooterTablePara 구조
        var tbl = subList.Descendants().FirstOrDefault(e => e.Name.LocalName == "tbl");
        if (tbl is not null)
        {
            var cells = tbl.Descendants().Where(e => e.Name.LocalName == "tc").ToList();
            if (cells.Count >= 3)
            {
                ReadHeaderFooterSlotFromCell(cells[0], content.Left,   ctx);
                ReadHeaderFooterSlotFromCell(cells[1], content.Center, ctx);
                ReadHeaderFooterSlotFromCell(cells[2], content.Right,  ctx);
                return;
            }
        }

        // 단락들로 저장된 경우 — 정렬로 슬롯 구분
        foreach (var p in subList.Elements().Where(e => e.Name.LocalName == "p"))
        {
            var para = ReadParagraph(p, ctx);
            if (para.Runs.Count == 0 || (para.Runs.Count == 1 && string.IsNullOrEmpty(para.Runs[0].Text)))
                continue;

            var slot = para.Style.Alignment switch
            {
                Alignment.Right  => content.Right,
                Alignment.Center => content.Center,
                _                => content.Left,
            };
            slot.Paragraphs.Add(para);
        }
    }

    private static void ReadHeaderFooterSlotFromCell(XElement cell, HeaderFooterSlot slot, ReadContext ctx)
    {
        var subList = cell.Descendants().FirstOrDefault(e => e.Name.LocalName == "subList");
        if (subList is null) return;
        foreach (var p in subList.Elements().Where(e => e.Name.LocalName == "p"))
        {
            var para = ReadParagraph(p, ctx);
            if (para.Runs.Count == 0 || (para.Runs.Count == 1 && string.IsNullOrEmpty(para.Runs[0].Text)))
                continue;
            slot.Paragraphs.Add(para);
        }
    }

    private static void ApplySecPr(XElement secPr, Section section)
    {
        var pagePr = secPr.Descendants().FirstOrDefault(e => e.Name.LocalName == "pagePr");
        if (pagePr is null) return;

        if (TryParseDouble(pagePr.Attribute("width")?.Value, out var w) && w > 0)
            section.Page.WidthMm = UnitConverter.HwpUnitToMm(w);
        if (TryParseDouble(pagePr.Attribute("height")?.Value, out var h) && h > 0)
            section.Page.HeightMm = UnitConverter.HwpUnitToMm(h);

        // 가로/세로 방향: 너비 > 높이 면 Landscape, 또는 landscape 속성이 "LANDSCAPE" 인 경우.
        var landscapeAttr = pagePr.Attribute("landscape")?.Value?.ToUpperInvariant();
        if (landscapeAttr == "LANDSCAPE"
            || (section.Page.WidthMm > 0 && section.Page.HeightMm > 0
                && section.Page.WidthMm > section.Page.HeightMm))
        {
            section.Page.Orientation = PageOrientation.Landscape;
        }

        var margin = pagePr.Descendants().FirstOrDefault(e => e.Name.LocalName == "margin");
        if (margin is null) return;

        if (TryParseDouble(margin.Attribute("left")?.Value,   out var ml)) section.Page.MarginLeftMm   = UnitConverter.HwpUnitToMm(ml);
        if (TryParseDouble(margin.Attribute("right")?.Value,  out var mr)) section.Page.MarginRightMm  = UnitConverter.HwpUnitToMm(mr);
        if (TryParseDouble(margin.Attribute("top")?.Value,    out var mt)) section.Page.MarginTopMm    = UnitConverter.HwpUnitToMm(mt);
        if (TryParseDouble(margin.Attribute("bottom")?.Value, out var mb)) section.Page.MarginBottomMm = UnitConverter.HwpUnitToMm(mb);
        if (TryParseDouble(margin.Attribute("header")?.Value, out var mh)) section.Page.MarginHeaderMm = UnitConverter.HwpUnitToMm(mh);
        if (TryParseDouble(margin.Attribute("footer")?.Value, out var mf)) section.Page.MarginFooterMm = UnitConverter.HwpUnitToMm(mf);
    }

    private static Table ReadTable(XElement wtbl, ReadContext ctx)
    {
        var table = new Table();

        // colCnt 속성으로 컬럼 수만 추정 (너비는 셀 단위로 주어지는 경우가 많음).
        if (TryParseInt(wtbl.Attribute("colCnt")?.Value) is { } colCount && colCount > 0)
        {
            for (int c = 0; c < colCount; c++)
            {
                table.Columns.Add(new TableColumn());
            }
        }

        // hp:tbl borderFillIDRef → header borderFill 정의에서 외곽선 색·두께·배경 회수.
        // per-side spec 인 경우 top/left 쪽 값 우선 (외곽 셀 외곽쪽 면이 표 외곽선 spec 임).
        if (TryParseInt(wtbl.Attribute("borderFillIDRef")?.Value) is { } tblBfId
            && ctx.Header.BorderFills.TryGetValue(tblBfId, out var tblBf))
        {
            ApplyTableBorderFromDef(table, tblBf);
        }

        // 열 너비 집계용 임시 버퍼 — 첫 본문 행의 셀 WidthMm 에서 역산.
        // 파싱 완료 후 Columns 에 반영한다.
        var colWidthBuffer = new List<double>();
        bool colWidthCaptured = false;

        foreach (var row in wtbl.Elements().Where(e => e.Name.LocalName == "tr"))
        {
            var tableRow = new TableRow();

            foreach (var cell in row.Elements().Where(e => e.Name.LocalName == "tc"))
            {
                var tableCell = new TableCell();

                // hp:tc borderFillIDRef → 셀 외곽선 색·두께·배경 회수.
                if (TryParseInt(cell.Attribute("borderFillIDRef")?.Value) is { } cellBfId
                    && ctx.Header.BorderFills.TryGetValue(cellBfId, out var cellBf))
                {
                    ApplyCellBorderFromDef(tableCell, cellBf);
                }

                // cellSpan 의 colSpan/rowSpan
                var span = cell.Elements().FirstOrDefault(e => e.Name.LocalName == "cellSpan");
                if (span is not null)
                {
                    if (TryParseInt(span.Attribute("colSpan")?.Value) is { } cs && cs > 0)
                        tableCell.ColumnSpan = cs;
                    if (TryParseInt(span.Attribute("rowSpan")?.Value) is { } rs && rs > 0)
                        tableCell.RowSpan = rs;
                }

                // cellSz 의 width (hwpunit) → mm
                var sz = cell.Elements().FirstOrDefault(e => e.Name.LocalName == "cellSz");
                if (sz is not null
                    && double.TryParse(sz.Attribute("width")?.Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var widthHwp))
                {
                    tableCell.WidthMm = UnitConverter.HwpUnitToMm(widthHwp);
                }

                // 셀 본문 — subList 안의 hp:p 들을 셀에 모은다 (인라인 그림·도형 포함).
                var seenInCell = new HashSet<XElement>();
                foreach (var d in cell.Descendants())
                {
                    switch (d.Name.LocalName)
                    {
                        case "p":
                            tableCell.Blocks.Add(ReadParagraph(d, ctx));
                            foreach (var pic in d.Descendants().Where(x => x.Name.LocalName == "pic"))
                            {
                                if (seenInCell.Add(pic) && TryReadPicture(pic, ctx, out var img))
                                {
                                    tableCell.Blocks.Add(img);
                                }
                            }
                            foreach (var shape in d.Descendants().Where(x => s_shapeLocalNames.Contains(x.Name.LocalName)))
                            {
                                if (seenInCell.Add(shape))
                                {
                                    tableCell.Blocks.Add(ReadShape(shape));
                                }
                            }
                            break;
                        case "tbl":
                            // 중첩 표
                            tableCell.Blocks.Add(ReadTable(d, ctx));
                            break;
                    }
                }

                if (tableCell.Blocks.Count == 0)
                {
                    tableCell.Blocks.Add(Paragraph.Of(string.Empty));
                }
                tableRow.Cells.Add(tableCell);
            }

            // 첫 본문 행에서 열 너비를 수집 (colspan 고려해 균등 분배)
            if (!colWidthCaptured && !tableRow.IsHeader && tableRow.Cells.Count > 0)
            {
                foreach (var c in tableRow.Cells)
                {
                    int cs = Math.Max(c.ColumnSpan, 1);
                    double perCol = c.WidthMm > 0 ? c.WidthMm / cs : 0;
                    for (int i = 0; i < cs; i++) colWidthBuffer.Add(perCol);
                }
                colWidthCaptured = true;
            }

            table.Rows.Add(tableRow);
        }

        // 열 너비 Columns 에 반영 (기존 컬럼 수 확장 포함)
        if (colWidthBuffer.Count > 0)
        {
            while (table.Columns.Count < colWidthBuffer.Count)
                table.Columns.Add(new TableColumn());
            for (int i = 0; i < colWidthBuffer.Count && i < table.Columns.Count; i++)
                if (colWidthBuffer[i] > 0) table.Columns[i].WidthMm = colWidthBuffer[i];

            // 표 총 너비 = 첫 본문 행 셀 너비 합계
            double totalW = colWidthBuffer.Sum();
            if (totalW > 0) table.WidthMm = totalW;
        }

        return table;
    }

    /// <summary>
    /// 표 borderFill 정의 → Table.BorderColor/BorderThicknessPt/BackgroundColor 매핑.
    /// per-side spec 일 수 있어 top 면을 대표로 사용 (writer 가 표 외곽 spec 으로 4면 동일 값).
    /// </summary>
    private static void ApplyTableBorderFromDef(Table table, HwpxBorderFillDef bf)
    {
        if (!string.IsNullOrEmpty(bf.TopColor))
            table.BorderColor = bf.TopColor;
        if (bf.TopWidthPt > 0)
            table.BorderThicknessPt = bf.TopWidthPt;
        if (!string.IsNullOrEmpty(bf.FillFaceColor))
            table.BackgroundColor = bf.FillFaceColor;
    }

    /// <summary>
    /// 셀 borderFill 정의 → TableCell 면별 테두리 프로퍼티 매핑.
    /// 4면이 모두 동일하면 공통값(BorderThicknessPt/BorderColor)만 세팅하고 면별 값은 null 유지.
    /// 면마다 다르면 per-side 프로퍼티를 채우고 공통값은 top 대표값으로 유지.
    /// </summary>
    private static void ApplyCellBorderFromDef(TableCell cell, HwpxBorderFillDef bf)
    {
        bool sideUniform = bf.TopColor == bf.BottomColor && bf.TopColor == bf.LeftColor && bf.TopColor == bf.RightColor
                        && Math.Abs(bf.TopWidthPt - bf.BottomWidthPt) < 0.01
                        && Math.Abs(bf.TopWidthPt - bf.LeftWidthPt)   < 0.01
                        && Math.Abs(bf.TopWidthPt - bf.RightWidthPt)  < 0.01;

        if (sideUniform)
        {
            // 4면 동일 — 공통값만 세팅, per-side 는 null 유지.
            if (!string.IsNullOrEmpty(bf.TopColor)) cell.BorderColor = bf.TopColor;
            if (bf.TopWidthPt > 0)                  cell.BorderThicknessPt = bf.TopWidthPt;
        }
        else
        {
            // 면마다 다름 — per-side 채우기. 공통 대표값은 inner(bottom) 면 사용 (HWPX 관례).
            cell.BorderTop    = MakeSide(bf.TopWidthPt,    bf.TopColor);
            cell.BorderBottom = MakeSide(bf.BottomWidthPt, bf.BottomColor);
            cell.BorderLeft   = MakeSide(bf.LeftWidthPt,   bf.LeftColor);
            cell.BorderRight  = MakeSide(bf.RightWidthPt,  bf.RightColor);

            var (repColor, repPt) = (bf.BottomColor, bf.BottomWidthPt);
            if (!string.IsNullOrEmpty(repColor)) cell.BorderColor = repColor;
            if (repPt > 0)                        cell.BorderThicknessPt = repPt;
        }

        if (!string.IsNullOrEmpty(bf.FillFaceColor))
            cell.BackgroundColor = bf.FillFaceColor;

        static CellBorderSide? MakeSide(double pt, string? color)
            => (pt > 0 || !string.IsNullOrEmpty(color)) ? new CellBorderSide(pt, color) : null;
    }

    private static bool TryReadPicture(XElement pic, ReadContext ctx, out ImageBlock image)
    {
        image = null!;

        // <hp:pic> 안의 <hp:img binaryItemIDRef="..."> 를 찾는다.
        var img = pic.Descendants().FirstOrDefault(e => e.Name.LocalName == "img");
        var binId = img?.Attribute("binaryItemIDRef")?.Value;
        if (string.IsNullOrEmpty(binId))
        {
            return false;
        }

        // BinData 인덱스에서 같은 stem 을 가진 entry 매칭.
        if (!ctx.BinDataIndex.TryGetValue(binId!, out var entryPath))
        {
            // 'image1' 형태 외에 'BinData/image1.png' 같은 풀 경로가 들어올 수도 있어 직접 시도.
            var direct = ctx.Archive.GetEntry(binId!);
            if (direct is null)
            {
                return false;
            }
            entryPath = direct.FullName;
        }

        var entry = ctx.Archive.GetEntry(entryPath);
        if (entry is null)
        {
            return false;
        }

        byte[] bytes;
        using (var stream = entry.Open())
        using (var ms = new MemoryStream())
        {
            stream.CopyTo(ms);
            bytes = ms.ToArray();
        }

        // curSz 의 width/height (hwpunit) → mm
        double widthMm = 0;
        double heightMm = 0;
        var curSz = pic.Descendants().FirstOrDefault(e => e.Name.LocalName == "curSz");
        if (curSz is not null)
        {
            if (double.TryParse(curSz.Attribute("width")?.Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var w))
                widthMm = UnitConverter.HwpUnitToMm(w);
            if (double.TryParse(curSz.Attribute("height")?.Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var h))
                heightMm = UnitConverter.HwpUnitToMm(h);
        }

        // WrapMode 로 textWrap 속성 매핑.
        var wrapMode = ImageWrapMode.Inline;
        var textWrap = pic.Attribute("textWrap")?.Value?.ToUpperInvariant();
        if (textWrap is "IN_FRONT_OF_TEXT")
            wrapMode = ImageWrapMode.InFrontOfText;
        else if (textWrap is "BEHIND_TEXT")
            wrapMode = ImageWrapMode.BehindText;
        else if (textWrap is "FLOAT_LEFT")
            wrapMode = ImageWrapMode.WrapLeft;
        else if (textWrap is "FLOAT_RIGHT")
            wrapMode = ImageWrapMode.WrapRight;

        // 오버레이 위치 (IN_FRONT_OF_TEXT / BEHIND_TEXT 일 때 의미 있음).
        double overlayX = 0, overlayY = 0;
        var posElem = pic.Descendants().FirstOrDefault(e => e.Name.LocalName == "pos");
        if (posElem is not null)
        {
            if (TryParseDouble(posElem.Attribute("horzOffset")?.Value, out var ox)) overlayX = UnitConverter.HwpUnitToMm(ox);
            if (TryParseDouble(posElem.Attribute("vertOffset")?.Value, out var oy)) overlayY = UnitConverter.HwpUnitToMm(oy);
        }

        image = new ImageBlock
        {
            MediaType = GuessMediaType(entry.FullName),
            Data = bytes,
            WidthMm = widthMm,
            HeightMm = heightMm,
            WrapMode = wrapMode,
            OverlayXMm = overlayX,
            OverlayYMm = overlayY,
        };
        return true;
    }

    private static string GuessMediaType(string path)
    {
        var ext = System.IO.Path.GetExtension(path).TrimStart('.').ToLowerInvariant();
        return ext switch
        {
            "png" => "image/png",
            "jpg" or "jpeg" => "image/jpeg",
            "gif" => "image/gif",
            "bmp" => "image/bmp",
            "tif" or "tiff" => "image/tiff",
            "svg" => "image/svg+xml",
            "webp" => "image/webp",
            _ => "application/octet-stream",
        };
    }

    private static Paragraph ReadParagraph(XElement wp, ReadContext ctx)
    {
        var header = ctx.Header;
        var paragraph = new Paragraph();

        // 1) styleIDRef 의 정의를 우선 적용 — outline + style 기본 paraPr/charPr 베이스로 둔다.
        var styleDef = header.GetStyle(TryParseInt(wp.Attribute("styleIDRef")?.Value));
        int? defaultCharPrId = null;
        if (styleDef is not null)
        {
            paragraph.Style.Outline = styleDef.Outline;
            var styleParaStyle = header.GetParagraphStyle(styleDef.ParaPrIdRef);
            if (styleParaStyle is not null)
            {
                CopyParagraphStyle(styleParaStyle, paragraph.Style);
            }
            defaultCharPrId = styleDef.CharPrIdRef;
        }

        if (wp.Attribute("pageBreak")?.Value == "1")
        {
            paragraph.Style.ForcePageBreakBefore = true;
        }

        // 2) paragraph 자신의 paraPrIDRef 가 있으면 그 위에 override.
        var directParaPrId = TryParseInt(wp.Attribute("paraPrIDRef")?.Value);
        var directParaStyle = header.GetParagraphStyle(directParaPrId);
        if (directParaStyle is not null)
        {
            CopyParagraphStyle(directParaStyle, paragraph.Style);
        }
        // header 가 비었거나 매핑이 없을 때 — 우리 자체 codec 의 0~3 약속을 fallback 으로 쓴다.
        if (directParaStyle is null && directParaPrId is { } pp)
        {
            paragraph.Style.Alignment = ParaPrIdToAlignment(pp);
        }
        // styleIDRef 직접 매핑(우리 자체 codec 의 1~6 = Heading) 도 보조로 유지.
        if (paragraph.Style.Outline == OutlineLevel.Body
            && TryParseInt(wp.Attribute("styleIDRef")?.Value) is { } sid && sid is >= 1 and <= 6)
        {
            paragraph.Style.Outline = (OutlineLevel)sid;
        }

        // <hp:p> 는 직속 자식으로 <hp:run> 들을 갖는 것이 표준이지만, 한컴 변종에선
        // 중간 wrapper(예: <hp:linesegarray> 다음 위치 등) 를 둘 수 있어 descendants 로 안전 매칭.
        // 단, hp:header/footer 서브리스트 안의 run/ctrl 은 이미 ReadHeaderOrFooter 가 처리했으므로
        // 여기서는 제외해야 한다.
        var hfDescInPara = new HashSet<XElement>();
        foreach (var hf in wp.Descendants().Where(e => e.Name.LocalName is "header" or "footer"))
            foreach (var d in hf.DescendantsAndSelf()) hfDescInPara.Add(d);
        // 글상자(drawText) subList 도 제외 (ReadTextBox 가 별도 처리)
        foreach (var tb in wp.Descendants().Where(e => e.Name.LocalName == "drawText"))
        {
            var sl = tb.Elements().FirstOrDefault(e => e.Name.LocalName == "subList");
            if (sl is not null)
                foreach (var d in sl.DescendantsAndSelf()) hfDescInPara.Add(d);
        }

        foreach (var elem in wp.Descendants())
        {
            if (hfDescInPara.Contains(elem)) continue;
            switch (elem.Name.LocalName)
            {
                case "run":
                    ReadRun(paragraph, elem, ctx, defaultCharPrId);
                    break;
                case "ctrl":
                    ReadCtrl(paragraph, elem, ctx);
                    break;
            }
        }

        if (paragraph.Runs.Count == 0)
        {
            paragraph.AddText(string.Empty);
        }
        return paragraph;
    }

    private static void ReadCtrl(Paragraph paragraph, XElement ctrl, ReadContext ctx)
    {
        var ctrlId = ctrl.Attribute("ctrlID")?.Value;
        switch (ctrlId)
        {
            case "FOOT_NOTE":
            case "END_NOTE":
                ReadNoteCtrl(paragraph, ctrl, ctx, ctrlId);
                break;

            // 페이지 번호 필드.
            case "PGNUM":
                paragraph.Runs.Add(new Run { Field = FieldType.Page });
                break;

            // 전체 페이지 수 필드.
            case "NPAGNUM":
            case "TOTAL_PGNUM":
                paragraph.Runs.Add(new Run { Field = FieldType.NumPages });
                break;

            // 날짜/시간 필드 — KS X 6101 의 DATE_TIME ctrl.
            case "DATE_TIME":
            case "DATE":
                paragraph.Runs.Add(new Run { Field = FieldType.Date });
                break;

            // 하이퍼링크 — url 속성(또는 href)으로 URL 전달.
            case "HYPERLINK":
            {
                var url = ctrl.Attribute("url")?.Value
                       ?? ctrl.Attribute("href")?.Value
                       ?? ctrl.Attribute("uri")?.Value;
                if (url is null) break;

                var subList = ctrl.Elements().FirstOrDefault(e => e.Name.LocalName == "subList");
                if (subList is null) break;

                foreach (var p in subList.Elements().Where(e => e.Name.LocalName == "p"))
                {
                    var linkPara = ReadParagraph(p, ctx);
                    foreach (var r in linkPara.Runs)
                    {
                        r.Url = url;
                        paragraph.Runs.Add(r);
                    }
                }
                break;
            }
        }
    }

    private static void ReadNoteCtrl(Paragraph paragraph, XElement ctrl, ReadContext ctx, string ctrlId)
    {
        var subList = ctrl.Elements().FirstOrDefault(e => e.Name.LocalName == "subList");
        if (subList is null) return;

        var entry = new FootnoteEntry { Id = Guid.NewGuid().ToString("N")[..8] };
        foreach (var elem in subList.Elements())
        {
            if (elem.Name.LocalName == "p")
                entry.Blocks.Add(ReadParagraph(elem, ctx));
        }
        if (entry.Blocks.Count == 0)
            entry.Blocks.Add(new Paragraph());

        if (ctrlId == "FOOT_NOTE")
        {
            ctx.Document.Footnotes.Add(entry);
            paragraph.Runs.Add(new Run { FootnoteId = entry.Id });
        }
        else
        {
            ctx.Document.Endnotes.Add(entry);
            paragraph.Runs.Add(new Run { EndnoteId = entry.Id });
        }
    }

    private static void ReadRun(Paragraph paragraph, XElement run, ReadContext ctx, int? defaultCharPrId)
    {
        var header = ctx.Header;
        var directCharPrId = TryParseInt(run.Attribute("charPrIDRef")?.Value);
        // 우선순위: run 의 charPrIDRef → style 의 charPrIDRef → 빈 RunStyle.
        var resolvedId = directCharPrId ?? defaultCharPrId;
        RunStyle style;
        if (resolvedId is { } id)
        {
            style = header.GetRunStyle(id);
            // header 에 charPr 정의가 없으면 — 우리 자체 codec 의 0~5 약속을 fallback 으로.
            if (!header.CharProperties.ContainsKey(id))
            {
                ApplyCharPrIdToStyle(id, style);
            }
        }
        else
        {
            style = new RunStyle();
        }

        // <hp:t> 는 보통 직속 자식이지만, 한컴 변종에선 <hp:t> 가 더 깊은 위치에 있을 수도 있고
        // <hp:tab>·<hp:lineBreak> 같은 형제와 섞일 수도 있어 descendants 로 텍스트 노드만 모음.
        // 단, run 안에 hp:header/footer/secPr/colPr 같은 특수 메타 ctrl 이 중첩된 경우
        // 그 안의 텍스트는 이미 ReadHeaderOrFooter 가 처리했으므로 제외.
        var skipDescendants = new HashSet<XElement>();
        foreach (var skip in run.Descendants()
            .Where(e => e.Name.LocalName is "header" or "footer" or "secPr" or "colPr"))
        {
            foreach (var d in skip.DescendantsAndSelf()) skipDescendants.Add(d);
        }
        // 글상자(drawText) subList 안의 텍스트도 건너뜀 (ReadTextBox 가 별도 처리)
        foreach (var tb in run.Descendants().Where(e => e.Name.LocalName == "drawText"))
        {
            var sl = tb.Elements().FirstOrDefault(e => e.Name.LocalName == "subList");
            if (sl is not null)
                foreach (var d in sl.DescendantsAndSelf()) skipDescendants.Add(d);
        }

        var sb = new StringBuilder();
        foreach (var elem in run.Descendants())
        {
            if (skipDescendants.Contains(elem)) continue;
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

    private static int? TryParseInt(string? raw)
        => int.TryParse(raw, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out var v) ? v : null;

    private static bool TryParseDouble(string? raw, out double value)
        => double.TryParse(raw, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out value);

    // ── 도형(ShapeObject) 파싱 ──────────────────────────────────────────────

    private static ShapeObject ReadShape(XElement shapeElem)
    {
        var localName = shapeElem.Name.LocalName;
        var shape = new ShapeObject();

        // WrapMode (textWrap 속성)
        shape.WrapMode = shapeElem.Attribute("textWrap")?.Value?.ToUpperInvariant() switch
        {
            "BEHIND_TEXT"       => ImageWrapMode.BehindText,
            "TOP_AND_BOTTOM"    => ImageWrapMode.Inline,
            "FLOAT_LEFT"        => ImageWrapMode.WrapLeft,
            "FLOAT_RIGHT"       => ImageWrapMode.WrapRight,
            _                   => ImageWrapMode.InFrontOfText,
        };

        // 크기: curSz → orgSz → sz 순서
        foreach (var sizeName in new[] { "curSz", "orgSz", "sz" })
        {
            var sz = shapeElem.Descendants().FirstOrDefault(e => e.Name.LocalName == sizeName);
            if (sz is null) continue;
            if (TryParseDouble(sz.Attribute("width")?.Value, out var sw) && sw > 0)
                shape.WidthMm = UnitConverter.HwpUnitToMm(sw);
            if (TryParseDouble(sz.Attribute("height")?.Value, out var sh) && sh > 0)
                shape.HeightMm = UnitConverter.HwpUnitToMm(sh);
            if (shape.WidthMm > 0 && shape.HeightMm > 0) break;
        }

        // 오버레이 위치 (hp:pos/@horzOffset, @vertOffset)
        var pos = shapeElem.Descendants().FirstOrDefault(e => e.Name.LocalName == "pos");
        if (pos is not null)
        {
            if (TryParseDouble(pos.Attribute("horzOffset")?.Value, out var ox)) shape.OverlayXMm = UnitConverter.HwpUnitToMm(ox);
            if (TryParseDouble(pos.Attribute("vertOffset")?.Value, out var oy)) shape.OverlayYMm = UnitConverter.HwpUnitToMm(oy);
        }

        // 회전 (hp:rotationInfo/@angle, HWPX 단위 = 1/10도)
        var rot = shapeElem.Descendants().FirstOrDefault(e => e.Name.LocalName == "rotationInfo");
        if (rot is not null && TryParseDouble(rot.Attribute("angle")?.Value, out var ang) && ang != 0)
        {
            shape.RotationAngleDeg = ang / 10.0;
        }

        // 선 속성 (hp:lineShape)
        var lineShape = shapeElem.Descendants().FirstOrDefault(e => e.Name.LocalName == "lineShape");
        if (lineShape is not null)
        {
            var sc = lineShape.Attribute("color")?.Value;
            if (!string.IsNullOrEmpty(sc))
                shape.StrokeColor = sc.StartsWith('#') ? sc : "#" + sc;
            if (TryParseDouble(lineShape.Attribute("width")?.Value, out var sw))
                shape.StrokeThicknessPt = sw / 100.0;
            shape.StrokeDash = lineShape.Attribute("style")?.Value?.ToUpperInvariant() switch
            {
                "DASHED" or "DASH"         => StrokeDash.Dashed,
                "DOTTED" or "DOT"          => StrokeDash.Dotted,
                "DASH_DOT" or "DASHDOT"    => StrokeDash.DashDot,
                _                          => StrokeDash.Solid,
            };
            shape.StartArrow = ParseShapeArrow(lineShape.Attribute("headStyle")?.Value);
            shape.EndArrow   = ParseShapeArrow(lineShape.Attribute("tailStyle")?.Value);
        }

        // 채우기 (hc:winBrush)
        var winBrush = shapeElem.Descendants().FirstOrDefault(e => e.Name.LocalName == "winBrush");
        if (winBrush is not null)
        {
            var fc = winBrush.Attribute("faceColor")?.Value;
            if (!string.IsNullOrEmpty(fc) && !fc.Equals("none", StringComparison.OrdinalIgnoreCase))
                shape.FillColor = fc.StartsWith('#') ? fc : "#" + fc;
            if (TryParseDouble(winBrush.Attribute("alpha")?.Value, out var alpha))
                shape.FillOpacity = 1.0 - Math.Clamp(alpha / 255.0, 0.0, 1.0);
        }

        // 도형 종류별 파싱
        switch (localName)
        {
            case "line":
            case "connector":
                shape.Kind = ShapeKind.Line;
                ReadLinePoints(shapeElem, shape);
                break;

            case "rect":
                if (TryParseInt(shapeElem.Attribute("ratio")?.Value) is { } ratio && ratio > 0)
                {
                    shape.Kind = ShapeKind.RoundedRect;
                    var minSide = Math.Min(
                        shape.WidthMm  > 0 ? shape.WidthMm  : 40,
                        shape.HeightMm > 0 ? shape.HeightMm : 30);
                    shape.CornerRadiusMm = ratio / 50.0 * (minSide / 2.0);
                }
                else
                {
                    shape.Kind = ShapeKind.Rectangle;
                }
                break;

            case "ellipse":
            case "arc":
                shape.Kind = ShapeKind.Ellipse;
                break;

            case "polygon":
                ReadPolygonPoints(shapeElem, shape);
                break;

            case "textBox":
                shape.Kind = ShapeKind.Rectangle;
                break;

            default:
                shape.Kind = ShapeKind.Rectangle;
                break;
        }

        return shape;
    }

    private static void ReadLinePoints(XElement shapeElem, ShapeObject shape)
    {
        var startPt = shapeElem.Descendants().FirstOrDefault(e => e.Name.LocalName == "startPt");
        var endPt   = shapeElem.Descendants().FirstOrDefault(e => e.Name.LocalName == "endPt");
        if (startPt is null || endPt is null) return;

        TryParseDouble(startPt.Attribute("x")?.Value, out var sxRaw);
        TryParseDouble(startPt.Attribute("y")?.Value, out var syRaw);
        TryParseDouble(endPt.Attribute("x")?.Value,   out var exRaw);
        TryParseDouble(endPt.Attribute("y")?.Value,   out var eyRaw);

        shape.Points = new List<ShapePoint>
        {
            new() { X = UnitConverter.HwpUnitToMm(sxRaw), Y = UnitConverter.HwpUnitToMm(syRaw) },
            new() { X = UnitConverter.HwpUnitToMm(exRaw), Y = UnitConverter.HwpUnitToMm(eyRaw) },
        };
    }

    private static void ReadPolygonPoints(XElement shapeElem, ShapeObject shape)
    {
        var rawPts = shapeElem.Descendants()
            .Where(e => e.Name.LocalName == "pt")
            .Select(pt =>
            {
                TryParseDouble(pt.Attribute("x")?.Value, out var px);
                TryParseDouble(pt.Attribute("y")?.Value, out var py);
                return new ShapePoint { X = UnitConverter.HwpUnitToMm(px), Y = UnitConverter.HwpUnitToMm(py) };
            })
            .ToList();

        if (rawPts.Count == 0)
        {
            shape.Kind = ShapeKind.Polygon;
            return;
        }

        // 마지막 점이 첫 점과 같으면 닫힌 다각형, 아니면 폴리선.
        bool closed = rawPts.Count >= 3
            && Math.Abs(rawPts[0].X - rawPts[^1].X) < 0.01
            && Math.Abs(rawPts[0].Y - rawPts[^1].Y) < 0.01;

        if (closed)
        {
            rawPts.RemoveAt(rawPts.Count - 1);
            shape.Kind = ShapeKind.Polygon;
        }
        else
        {
            shape.Kind = ShapeKind.Polyline;
        }
        shape.Points = rawPts;
    }

    private static ShapeArrow ParseShapeArrow(string? style)
        => style?.ToUpperInvariant() switch
        {
            "OPEN" or "ARROW"      => ShapeArrow.Open,
            "FILLED" or "SOLID"    => ShapeArrow.Filled,
            "DIAMOND"              => ShapeArrow.Diamond,
            "CIRCLE"               => ShapeArrow.Circle,
            _                      => ShapeArrow.None,
        };

    private static void CopyParagraphStyle(ParagraphStyle src, ParagraphStyle dst)
    {
        dst.Alignment = src.Alignment;
        dst.LineHeightFactor = src.LineHeightFactor;
        dst.SpaceBeforePt = src.SpaceBeforePt;
        dst.SpaceAfterPt = src.SpaceAfterPt;
        dst.IndentFirstLineMm = src.IndentFirstLineMm;
        dst.IndentLeftMm = src.IndentLeftMm;
        dst.IndentRightMm = src.IndentRightMm;
        // Outline / ListMarker 는 styleDef 또는 별도 신호로 결정하므로 여기서는 덮지 않는다.
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

    private static XDocument? LoadXml(ZipArchive archive, string path, List<string>? errors = null)
    {
        var entry = archive.GetEntry(path);
        if (entry is null)
        {
            errors?.Add($"missing entry: {path}");
            return null;
        }
        try
        {
            using var stream = entry.Open();
            // BOM 자동 감지 + UTF-8 기본. 일부 한컴 hwpx 의 packaging XML 이 BOM 으로 시작해
            // raw stream → XDocument.Load 가 'Data at the root level is invalid' 로 throw 하는 경우 회피.
            using var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
            return XDocument.Load(reader);
        }
        catch (Exception ex)
        {
            errors?.Add($"{path}: {ex.GetType().Name}: {ex.Message}");
            return null;
        }
    }
}

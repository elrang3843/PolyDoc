using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using PolyDonky.Core;

namespace PolyDonky.Codecs.Hwpx;

public sealed class HwpxWriter : IDocumentWriter
{
    public string FormatId => "hwpx";

    // BuildNoteCtrlRun 에서 참조하는 각주/미주 목록 — Write() 호출 간 스레드 안전하지 않으므로 인스턴스 필드.
    private IList<FootnoteEntry> _pendingFootnotes = Array.Empty<FootnoteEntry>();
    private IList<FootnoteEntry> _pendingEndnotes  = Array.Empty<FootnoteEntry>();

    private static readonly XNamespace OpfContainer = HwpxNamespaces.OpfContainer;
    private static readonly XNamespace Opf = HwpxNamespaces.OpfPackage;
    private static readonly XNamespace Dc = HwpxNamespaces.DcMetadata;
    private static readonly XNamespace Hh = HwpxNamespaces.Head;
    private static readonly XNamespace Hp = HwpxNamespaces.Paragraph;
    private static readonly XNamespace Hs = HwpxNamespaces.Section;
    private static readonly XNamespace Hc = HwpxNamespaces.Common;
    private static readonly XNamespace Ha = HwpxNamespaces.Application;
    private static readonly XNamespace Hv = HwpxNamespaces.HcfVersion;

    // Additional namespaces found in real Hancom-generated files
    private static readonly XNamespace Hp10 = "http://www.hancom.co.kr/hwpml/2016/paragraph";
    private static readonly XNamespace HpfSchema = "http://www.hancom.co.kr/schema/2011/hpf";
    private static readonly XNamespace HwpUnitChar = "http://www.hancom.co.kr/hwpml/2016/HwpUnitChar";
    private static readonly XNamespace Hm = "http://www.hancom.co.kr/hwpml/2011/master-page";
    private static readonly XNamespace Hhs = "http://www.hancom.co.kr/hwpml/2011/history";
    private static readonly XNamespace OoxmlChart = "http://www.hancom.co.kr/hwpml/2016/ooxmlchart";
    private static readonly XNamespace Epub = "http://www.idpf.org/2007/ops";
    private static readonly XNamespace Config = "urn:oasis:names:tc:opendocument:xmlns:config:1.0";
    private static readonly XNamespace Rdf = "http://www.w3.org/1999/02/22-rdf-syntax-ns#";
    private static readonly XNamespace PkgMeta = "http://www.hancom.co.kr/hwpml/2016/meta/pkg#";

    // ── key types for dynamic registry ──────────────────────────────────────

    private readonly record struct RunStyleKey(
        string? FontFamily, long HeightUnit,
        bool Bold, bool Italic, bool Underline, bool Strikethrough,
        bool Overline, bool Superscript, bool Subscript,
        string? FgHex, string? BgHex)
    {
        internal static RunStyleKey From(RunStyle s) => new(
            s.FontFamily,
            (long)Math.Round(s.FontSizePt * 100),
            s.Bold, s.Italic, s.Underline, s.Strikethrough,
            s.Overline, s.Superscript, s.Subscript,
            s.Foreground?.ToHex(), s.Background?.ToHex());
    }

    private readonly record struct ParaStyleKey(
        Alignment Alignment, long LineHeightX1000,
        long SpBeforeX100, long SpAfterX100,
        long IndFirstX100, long IndLeftX100, long IndRightX100)
    {
        internal static ParaStyleKey From(ParagraphStyle s) => new(
            s.Alignment,
            (long)Math.Round(s.LineHeightFactor * 1000),
            (long)Math.Round(s.SpaceBeforePt * 100),
            (long)Math.Round(s.SpaceAfterPt * 100),
            (long)Math.Round(s.IndentFirstLineMm * 100),
            (long)Math.Round(s.IndentLeftMm * 100),
            (long)Math.Round(s.IndentRightMm * 100));
    }

    // ── WriteContext ─────────────────────────────────────────────────────────

    private sealed class WriteContext
    {
        private int _nextImageId = 1;
        private int _nextParaId = 0;
        private long _nextObjId  = 2042848500;
        private int  _nextZOrder = 0;
        private long _nextInstId = 969106680;
        private readonly Dictionary<string, string> _imageIdByHash = new(StringComparer.Ordinal);
        // Registration order, for binDataList in header.xml + opf:item entries in content.hpf.
        private readonly List<BinDataInfo> _binData = new();
        public IReadOnlyList<BinDataInfo> BinData => _binData;
        public sealed record BinDataInfo(string Id, string MediaType, string Extension, string ZipPath);

        // 표·셀의 외곽선/배경색을 위한 동적 borderFill 등록.
        // id 1=page, 2=char, 3=table-default(고정 SOLID 0.12mm), 4 부터 custom.
        public const int FirstCustomBorderFillId = 4;
        private readonly List<BorderFillSpec> _customBorderFills = new();
        public IReadOnlyList<BorderFillSpec> CustomBorderFills => _customBorderFills;
        // 4면 각각 type/color/width 가 다를 수 있도록 확장 — 표 외곽 cell 의 외곽쪽
        // 면이 inner 면과 다른 색·두께를 가질 수 있게 함.
        public sealed record BorderFillSpec(
            string TopType,    string TopColor,    string TopWidth,
            string BottomType, string BottomColor, string BottomWidth,
            string LeftType,   string LeftColor,   string LeftWidth,
            string RightType,  string RightColor,  string RightWidth,
            string? FillColor)
        {
            // 4면 동일 spec 헬퍼 (대부분 케이스).
            public static BorderFillSpec Uniform(string type, string color, string width, string? fill = null)
                => new(type, color, width, type, color, width, type, color, width, type, color, width, fill);
        }
        public int RegisterCustomBorderFill(BorderFillSpec spec)
        {
            int idx = _customBorderFills.IndexOf(spec);
            if (idx >= 0) return idx + FirstCustomBorderFillId;
            _customBorderFills.Add(spec);
            return _customBorderFills.Count - 1 + FirstCustomBorderFillId;
        }

        private readonly Dictionary<RunStyleKey, int> _runStyleIds = new();
        private readonly List<RunStyle> _runStyles = new();

        private readonly Dictionary<ParaStyleKey, int> _paraStyleIds = new();
        private readonly List<ParagraphStyle> _paraStyles = new();

        private readonly Dictionary<string, int> _fontIds = new(StringComparer.OrdinalIgnoreCase);
        private readonly List<string> _fonts = new();

        public WriteContext(ZipArchive archive)
        {
            Archive = archive;
            // id 0 = default font/run/para — always present
            InternalRegisterFont("맑은 고딕");
            RegisterRunStyle(new RunStyle());
            RegisterParagraphStyle(new ParagraphStyle());
        }

        public ZipArchive Archive { get; }
        public IReadOnlyList<string> Fonts => _fonts;
        public IReadOnlyList<RunStyle> RunStyles => _runStyles;
        public IReadOnlyList<ParagraphStyle> ParaStyles => _paraStyles;

        private int _nextCtrlId = 1;
        private int _nextSubListId = 1;

        public void ResetParaId() => _nextParaId = 0;
        public int  NextParaId()  => _nextParaId++;
        public long NextObjId()   => _nextObjId++;
        public int  NextZOrder()  => _nextZOrder++;
        public long NextInstId()  => _nextInstId++;
        public int  NextCtrlId()  => _nextCtrlId++;
        public int  NextSubListId() => _nextSubListId++;

        private int InternalRegisterFont(string family)
        {
            if (_fontIds.TryGetValue(family, out var id)) return id;
            id = _fonts.Count;
            _fontIds[family] = id;
            _fonts.Add(family);
            return id;
        }

        public int RegisterRunStyle(RunStyle s)
        {
            var key = RunStyleKey.From(s);
            if (_runStyleIds.TryGetValue(key, out var id)) return id;
            id = _runStyles.Count;
            _runStyleIds[key] = id;
            _runStyles.Add(HwpxHeader.CloneRunStyle(s));
            if (!string.IsNullOrEmpty(s.FontFamily)) InternalRegisterFont(s.FontFamily);
            return id;
        }

        public int RegisterParagraphStyle(ParagraphStyle s)
        {
            var key = ParaStyleKey.From(s);
            if (_paraStyleIds.TryGetValue(key, out var id)) return id;
            id = _paraStyles.Count;
            _paraStyleIds[key] = id;
            _paraStyles.Add(HwpxHeader.CloneParagraphStyle(s));
            return id;
        }

        public int RunStyleId(RunStyle s) =>
            _runStyleIds.GetValueOrDefault(RunStyleKey.From(s), 0);

        public int ParaStyleId(ParagraphStyle s) =>
            _paraStyleIds.GetValueOrDefault(ParaStyleKey.From(s), 0);

        public int FontId(string? family) =>
            string.IsNullOrEmpty(family) ? 0 : _fontIds.GetValueOrDefault(family, 0);

        public void RegisterFromBlock(Block block)
        {
            switch (block)
            {
                case Paragraph p:
                    RegisterParagraphStyle(p.Style);
                    foreach (var run in p.Runs) RegisterRunStyle(run.Style);
                    break;
                case Table t:
                    foreach (var row in t.Rows)
                        foreach (var cell in row.Cells)
                            foreach (var inner in cell.Blocks)
                                RegisterFromBlock(inner);
                    break;
                // ImageBlock / OpaqueBlock → default styles (id 0 already registered)
            }
        }

        public string AddImage(byte[] data, string mediaType)
        {
            Span<byte> hash = stackalloc byte[SHA256.HashSizeInBytes];
            SHA256.HashData(data, hash);
            var hashKey = Convert.ToHexStringLower(hash);
            if (_imageIdByHash.TryGetValue(hashKey, out var existing)) return existing;
            var id = $"image{_nextImageId++}";
            var ext = ExtensionForMediaType(mediaType);
            var path = $"{HwpxPaths.BinDataDir}{id}{ext}";
            var entry = Archive.CreateEntry(path, CompressionLevel.Optimal);
            using (var stream = entry.Open())
                stream.Write(data, 0, data.Length);
            _imageIdByHash[hashKey] = id;
            _binData.Add(new BinDataInfo(id, mediaType, ext, path));
            return id;
        }

        // Pre-pass: register all ImageBlock binaries before header.xml/content.hpf are written.
        // Required so that <hp:img binaryItemIDRef="N"> resolves via header binDataList,
        // and so content.hpf manifest contains BinData/* entries (else Hangul rejects the file).
        public void PreRegisterImages(PolyDonkyument document)
        {
            foreach (var section in document.Sections)
                foreach (var block in section.Blocks)
                    PreRegisterImagesFromBlock(block);
        }

        private void PreRegisterImagesFromBlock(Block block)
        {
            switch (block)
            {
                case ImageBlock img when img.Data.Length > 0:
                    AddImage(img.Data, img.MediaType);
                    break;
                case Table t:
                    foreach (var row in t.Rows)
                        foreach (var cell in row.Cells)
                            foreach (var inner in cell.Blocks)
                                PreRegisterImagesFromBlock(inner);
                    break;
            }
        }

        // 표·셀의 동적 borderFill 사전 등록 — header.xml 쓰기 전에 호출되어야
        // 모든 borderFill id 가 헤더에 포함되고 section 의 borderFillIDRef 가 dangling 되지 않는다.
        public void PreRegisterTableBorderFills(PolyDonkyument document)
        {
            foreach (var section in document.Sections)
                foreach (var block in section.Blocks)
                    PreRegisterBorderFillsFromBlock(block);
        }

        private static string SnapBorderWidthMm(double pt)
        {
            if (pt <= 0) return "0.12 mm";
            double mm = pt * 0.3527777778;
            double[] standard = { 0.1, 0.12, 0.15, 0.2, 0.25, 0.3, 0.4, 0.5, 0.6, 0.7, 1.0, 1.5, 2.0, 3.0, 4.0, 5.0 };
            double best = standard[0];
            double bestDiff = Math.Abs(best - mm);
            foreach (var s in standard)
            {
                double d = Math.Abs(s - mm);
                if (d < bestDiff) { best = s; bestDiff = d; }
            }
            return $"{best:0.##} mm";
        }

        private void PreRegisterBorderFillsFromBlock(Block block)
        {
            if (block is Table table)
            {
                // 표 외곽 (전체 표 borderFillIDRef) — 4면 동일 spec.
                string tableColor = string.IsNullOrEmpty(table.BorderColor) ? "#000000" : table.BorderColor!;
                string tableWidth = SnapBorderWidthMm(table.BorderThicknessPt);
                RegisterCustomBorderFill(BorderFillSpec.Uniform(
                    "SOLID", tableColor, tableWidth, table.BackgroundColor));

                int rowCount = table.Rows.Count;
                int colCount = table.Columns.Count;
                if (colCount == 0 && rowCount > 0)
                    colCount = table.Rows.Max(r => r.Cells.Sum(c => Math.Max(c.ColumnSpan, 1)));

                for (int r = 0; r < rowCount; r++)
                {
                    int c = 0;
                    foreach (var cell in table.Rows[r].Cells)
                    {
                        int colSpan = Math.Max(cell.ColumnSpan, 1);
                        int rowSpan = Math.Max(cell.RowSpan, 1);
                        bool top    = r == 0;
                        bool bottom = r + rowSpan >= rowCount;
                        bool left   = c == 0;
                        bool right  = c + colSpan >= colCount;

                        string innerColor = !string.IsNullOrEmpty(cell.BorderColor) ? cell.BorderColor! : tableColor;
                        string innerWidth = cell.BorderThicknessPt > 0
                            ? SnapBorderWidthMm(cell.BorderThicknessPt)
                            : tableWidth;
                        string? cellFill = !string.IsNullOrEmpty(cell.BackgroundColor) ? cell.BackgroundColor
                                         : table.BackgroundColor;

                        // 면별 per-side 값 우선, 없으면 inner 대표값 → 표 외곽이면 표 값으로 덮어씀.
                        (string tc, string tw) = SideOrFallback(top,    cell.BorderTop,    tableColor, tableWidth, innerColor, innerWidth);
                        (string bc, string bw) = SideOrFallback(bottom, cell.BorderBottom, tableColor, tableWidth, innerColor, innerWidth);
                        (string lc, string lw) = SideOrFallback(left,   cell.BorderLeft,   tableColor, tableWidth, innerColor, innerWidth);
                        (string rc, string rw) = SideOrFallback(right,  cell.BorderRight,  tableColor, tableWidth, innerColor, innerWidth);
                        RegisterCustomBorderFill(new BorderFillSpec(
                            "SOLID", tc, tw,
                            "SOLID", bc, bw,
                            "SOLID", lc, lw,
                            "SOLID", rc, rw,
                            cellFill));
                        foreach (var inner in cell.Blocks)
                            PreRegisterBorderFillsFromBlock(inner);
                        c += colSpan;
                    }
                }
            }
        }

        // 외곽 여부 isOuter=true 이면 표 외곽선 값 고정, 아니면 per-side 값 → 공통 inner 값 순서로 사용.
        private static (string color, string width) SideOrFallback(
            bool isOuter, CellBorderSide? side,
            string tableColor, string tableWidth,
            string innerColor, string innerWidth)
        {
            if (isOuter) return (tableColor, tableWidth);
            var c = side.HasValue && !string.IsNullOrEmpty(side.Value.Color) ? side.Value.Color! : innerColor;
            var w = side.HasValue && side.Value.ThicknessPt > 0 ? SnapBorderWidthMm(side.Value.ThicknessPt) : innerWidth;
            return (c, w);
        }

        private static string ExtensionForMediaType(string mt) => mt switch
        {
            "image/png"     => ".png",
            "image/jpeg"    => ".jpg",
            "image/gif"     => ".gif",
            "image/bmp"     => ".bmp",
            "image/tiff"    => ".tif",
            "image/svg+xml" => ".svg",
            "image/webp"    => ".webp",
            _               => ".bin",
        };
    }

    // ── public Write ─────────────────────────────────────────────────────────

    public void Write(PolyDonkyument document, Stream output)
    {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(output);

        _pendingFootnotes = document.Footnotes;
        _pendingEndnotes  = document.Endnotes;

        using var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true);
        WriteRawText(archive, HwpxPaths.Mimetype, HwpxPaths.MimetypeContent, CompressionLevel.NoCompression);
        WriteContainerXml(archive);
        WriteContainerRdf(archive, document);
        WriteManifestXml(archive);
        WriteSettingsXml(archive);
        WriteVersionXml(archive);
        WritePreviewText(archive, document);

        var ctx = new WriteContext(archive);

        // Style/font registration pass.
        foreach (var section in document.Sections)
            foreach (var block in section.Blocks)
                ctx.RegisterFromBlock(block);

        // Image pre-registration: assigns binData ids and writes BinData/* entries to ZIP.
        // Must precede WriteContentHpf (manifest needs BinData items) and WriteHeaderXml
        // (refList needs binDataList for binaryItemIDRef resolution).
        ctx.PreRegisterImages(document);

        // 표 동적 borderFill 사전 등록 — header.xml 쓰기 전에 모든 borderFill id 확정.
        ctx.PreRegisterTableBorderFills(document);

        WriteContentHpf(archive, document, ctx);

        int cnt = document.Sections.Count;
        WriteHeaderXml(archive, ctx, Math.Max(cnt, 1));
        for (int i = 0; i < cnt; i++)
            WriteSectionXml(archive, i, document.Sections[i], ctx);
        if (cnt == 0)
            WriteSectionXml(archive, 0, new Section(), ctx);
    }

    // ── static XML helpers ───────────────────────────────────────────────────

    private static void WriteRawText(ZipArchive archive, string path, string content, CompressionLevel level)
    {
        var entry = archive.CreateEntry(path, level);
        using var stream = entry.Open();
        stream.Write(Encoding.UTF8.GetBytes(content));
    }

    private static void WriteXml(ZipArchive archive, string path, XDocument doc)
    {
        var entry = archive.CreateEntry(path, CompressionLevel.Optimal);
        using var stream = entry.Open();
        using var writer = new StreamWriter(stream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        doc.Save(writer, SaveOptions.None);
    }

    // ── package structure ────────────────────────────────────────────────────

    private void WriteContainerXml(ZipArchive archive)
    {
        // Real Hancom format: ocf: prefix, xmlns:hpf extra ns, THREE rootfiles.
        var ocf = XNamespace.Get("urn:oasis:names:tc:opendocument:xmlns:container");
        var doc = new XDocument(new XDeclaration("1.0", "utf-8", null),
            new XElement(ocf + "container",
                new XAttribute(XNamespace.Xmlns + "ocf", ocf.NamespaceName),
                new XAttribute(XNamespace.Xmlns + "hpf", HpfSchema.NamespaceName),
                new XElement(ocf + "rootfiles",
                    new XElement(ocf + "rootfile",
                        new XAttribute("full-path", HwpxPaths.ContentHpf),
                        new XAttribute("media-type", "application/hwpml-package+xml")),
                    new XElement(ocf + "rootfile",
                        new XAttribute("full-path", "Preview/PrvText.txt"),
                        new XAttribute("media-type", "text/plain")),
                    new XElement(ocf + "rootfile",
                        new XAttribute("full-path", "META-INF/container.rdf"),
                        new XAttribute("media-type", "application/rdf+xml")))));
        WriteXml(archive, HwpxPaths.ContainerXml, doc);
    }

    private void WriteContainerRdf(ZipArchive archive, PolyDonkyument document)
    {
        // RDF metadata describing which files are the header and section files.
        // Use plain string URIs for rdf:resource values (not XNamespace + localname notation).
        int sectionCount = Math.Max(document.Sections.Count, 1);
        const string PkgNs = "http://www.hancom.co.kr/hwpml/2016/meta/pkg#";

        var rdf = new XElement(Rdf + "RDF",
            new XAttribute(XNamespace.Xmlns + "rdf", Rdf.NamespaceName));

        // Header file descriptor
        rdf.Add(new XElement(Rdf + "Description",
            new XAttribute(Rdf + "about", ""),
            new XElement(PkgMeta + "hasPart",
                new XAttribute(XNamespace.Xmlns + "ns0", PkgNs),
                new XAttribute(Rdf + "resource", "Contents/header.xml"))));
        rdf.Add(new XElement(Rdf + "Description",
            new XAttribute(Rdf + "about", "Contents/header.xml"),
            new XElement(Rdf + "type",
                new XAttribute(Rdf + "resource", PkgNs + "HeaderFile"))));

        // Section file descriptors
        for (int i = 0; i < sectionCount; i++)
        {
            var sectionPath = HwpxPaths.SectionXml(i);
            rdf.Add(new XElement(Rdf + "Description",
                new XAttribute(Rdf + "about", ""),
                new XElement(PkgMeta + "hasPart",
                    new XAttribute(XNamespace.Xmlns + "ns0", PkgNs),
                    new XAttribute(Rdf + "resource", sectionPath))));
            rdf.Add(new XElement(Rdf + "Description",
                new XAttribute(Rdf + "about", sectionPath),
                new XElement(Rdf + "type",
                    new XAttribute(Rdf + "resource", PkgNs + "SectionFile"))));
        }

        // Document type descriptor
        rdf.Add(new XElement(Rdf + "Description",
            new XAttribute(Rdf + "about", ""),
            new XElement(Rdf + "type",
                new XAttribute(Rdf + "resource", PkgNs + "Document"))));

        WriteXml(archive, "META-INF/container.rdf",
            new XDocument(new XDeclaration("1.0", "utf-8", null), rdf));
    }

    private void WriteManifestXml(ZipArchive archive)
    {
        // ODF-style empty manifest — required by Hangul's package validator.
        var ns = XNamespace.Get("urn:oasis:names:tc:opendocument:xmlns:manifest:1.0");
        var doc = new XDocument(new XDeclaration("1.0", "utf-8", null),
            new XElement(ns + "manifest",
                new XAttribute(XNamespace.Xmlns + "odf", ns.NamespaceName)));
        WriteXml(archive, HwpxPaths.ManifestXml, doc);
    }

    private void WriteSettingsXml(ZipArchive archive)
    {
        // Real Hancom uses ha:HWPApplicationSetting (singular, not plural) with config ns.
        var doc = new XDocument(new XDeclaration("1.0", "utf-8", null),
            new XElement(Ha + "HWPApplicationSetting",
                new XAttribute(XNamespace.Xmlns + HwpxNamespaces.PrefixHa, Ha.NamespaceName),
                new XAttribute(XNamespace.Xmlns + "config", Config.NamespaceName)));
        WriteXml(archive, HwpxPaths.SettingsXml, doc);
    }

    private void WriteContentHpf(ZipArchive archive, PolyDonkyument document, WriteContext ctx)
    {
        // All paths in manifest are ABSOLUTE from ZIP root (not relative to content.hpf location).
        // Real Hancom includes header in spine with linear="yes", sections without.
        int sectionCount = Math.Max(document.Sections.Count, 1);

        var manifest = new XElement(Opf + "manifest");
        manifest.Add(new XElement(Opf + "item",
            new XAttribute("id", "header"),
            new XAttribute("href", "Contents/header.xml"),
            new XAttribute("media-type", "application/xml")));
        manifest.Add(new XElement(Opf + "item",
            new XAttribute("id", "settings"),
            new XAttribute("href", "settings.xml"),
            new XAttribute("media-type", "application/xml")));

        // Embedded images registered by PreRegisterImages. Hangul rejects the package
        // if BinData/* files exist on disk but aren't listed in the manifest.
        // 실 한컴 파일은 isEmbeded="1" 속성을 포함 (오타 그대로 — 한 'd').
        foreach (var bd in ctx.BinData)
            manifest.Add(new XElement(Opf + "item",
                new XAttribute("id",         bd.Id),
                new XAttribute("href",       bd.ZipPath),
                new XAttribute("media-type", bd.MediaType),
                new XAttribute("isEmbeded",  "1")));

        var spine = new XElement(Opf + "spine");
        // Header in spine with linear="yes" (matches real Hancom format)
        spine.Add(new XElement(Opf + "itemref",
            new XAttribute("idref", "header"),
            new XAttribute("linear", "yes")));

        for (int i = 0; i < sectionCount; i++)
        {
            var id = $"section{i}";
            manifest.Add(new XElement(Opf + "item",
                new XAttribute("id", id),
                new XAttribute("href", $"Contents/section{i}.xml"),
                new XAttribute("media-type", "application/xml")));
            spine.Add(new XElement(Opf + "itemref", new XAttribute("idref", id)));
        }

        var metadata = new XElement(Opf + "metadata");
        if (!string.IsNullOrEmpty(document.Metadata.Title))
            metadata.Add(new XElement(Opf + "title", document.Metadata.Title));
        if (!string.IsNullOrEmpty(document.Metadata.Author))
            metadata.Add(new XElement(Opf + "creator", document.Metadata.Author));
        metadata.Add(new XElement(Opf + "language", document.Metadata.Language ?? "ko"));

        // All namespaces that real Hancom includes on opf:package root
        var package = new XElement(Opf + "package",
            new XAttribute(XNamespace.Xmlns + "ha",   Ha.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "hp",   Hp.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "hp10", Hp10.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "hs",   Hs.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "hc",   Hc.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "hh",   Hh.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "hhs",  Hhs.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "hm",   Hm.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "hpf",  HpfSchema.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "dc",   Dc.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "opf",  Opf.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "ooxmlchart", OoxmlChart.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "hwpunitchar", HwpUnitChar.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "epub", Epub.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "config", Config.NamespaceName),
            new XAttribute("version", string.Empty),
            new XAttribute("unique-identifier", string.Empty),
            new XAttribute("id", string.Empty),
            metadata, manifest, spine);

        WriteXml(archive, HwpxPaths.ContentHpf, new XDocument(new XDeclaration("1.0", "utf-8", null), package));
    }

    private void WriteVersionXml(ZipArchive archive)
    {
        // "tagetApplication" is the actual Hancom typo in the real format — must match exactly.
        var doc = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"),
            new XElement(Hv + "HCFVersion",
                new XAttribute(XNamespace.Xmlns + HwpxNamespaces.PrefixHv, Hv.NamespaceName),
                new XAttribute("tagetApplication", "WORDPROCESSOR"),
                new XAttribute("major", "5"),
                new XAttribute("minor", "1"),
                new XAttribute("micro", "0"),
                new XAttribute("buildNumber", "1"),
                new XAttribute("os", "1"),
                new XAttribute("xmlVersion", "1.4"),
                new XAttribute("application", "Hancom Office Hangul"),
                new XAttribute("appVersion", $"{HwpxFormat.ProducedBy} {HwpxFormat.Version}")));
        WriteXml(archive, HwpxPaths.VersionXml, doc);
    }

    private void WritePreviewText(ZipArchive archive, PolyDonkyument document)
    {
        var text = new StringBuilder();
        foreach (var section in document.Sections)
            foreach (var block in section.Blocks)
                if (block is Paragraph p)
                    text.AppendLine(p.GetPlainText());
        WriteRawText(archive, "Preview/PrvText.txt",
            text.Length > 0 ? text.ToString() : string.Empty,
            CompressionLevel.Optimal);
    }

    // ── header.xml (dynamic registry) ────────────────────────────────────────

    // KS X 6101 language groups — charPr.fontRef has one slot per group.
    private static readonly string[] FontLangs =
        { "HANGUL", "LATIN", "HANJA", "JAPANESE", "OTHER", "SYMBOL", "USER" };

    // A4 page in hwpunit (1 hwpunit = 1/7200 inch).  210mm×297mm.
    private const long A4W = 59528;
    private const long A4H = 84188;
    // Standard Korean default margins (mm → hwpunit: mm / 25.4 × 7200).
    private const long MarginLeft   = 8503;   // ~30 mm
    private const long MarginRight  = 8503;   // ~30 mm
    private const long MarginTop    = 5669;   // ~20 mm
    private const long MarginBottom = 4961;   // ~17.5 mm
    private const long MarginHead   = 4252;   // ~15 mm
    private const long MarginFoot   = 4252;   // ~15 mm
    // Content area width (used for lineseg horzsize)
    private const long ContentWidth = A4W - MarginLeft - MarginRight; // = 42522

    // Shared namespace declarations for header/section XML root elements (matching real Hancom)
    private static XAttribute[] AllNamespaceAttrs() => [
        new(XNamespace.Xmlns + "ha",          Ha.NamespaceName),
        new(XNamespace.Xmlns + "hp",          Hp.NamespaceName),
        new(XNamespace.Xmlns + "hp10",        Hp10.NamespaceName),
        new(XNamespace.Xmlns + "hs",          Hs.NamespaceName),
        new(XNamespace.Xmlns + "hc",          Hc.NamespaceName),
        new(XNamespace.Xmlns + "hh",          Hh.NamespaceName),
        new(XNamespace.Xmlns + "hhs",         Hhs.NamespaceName),
        new(XNamespace.Xmlns + "hm",          Hm.NamespaceName),
        new(XNamespace.Xmlns + "hpf",         HpfSchema.NamespaceName),
        new(XNamespace.Xmlns + "dc",          Dc.NamespaceName),
        new(XNamespace.Xmlns + "opf",         Opf.NamespaceName),
        new(XNamespace.Xmlns + "ooxmlchart",  OoxmlChart.NamespaceName),
        new(XNamespace.Xmlns + "hwpunitchar", HwpUnitChar.NamespaceName),
        new(XNamespace.Xmlns + "epub",        Epub.NamespaceName),
        new(XNamespace.Xmlns + "config",      Config.NamespaceName),
    ];

    private void WriteHeaderXml(ZipArchive archive, WriteContext ctx, int sectionCount)
    {
        // fontfaces — 7 language groups, each carrying the full font registry.
        var fontFaces = new XElement(Hh + "fontfaces",
            new XAttribute("itemCnt", FontLangs.Length.ToString()));
        foreach (var lang in FontLangs)
        {
            // Real Hancom uses fontCnt (not count) attribute name.
            var ff = new XElement(Hh + "fontface",
                new XAttribute("lang", lang),
                new XAttribute("fontCnt", ctx.Fonts.Count.ToString()));
            for (int i = 0; i < ctx.Fonts.Count; i++)
            {
                ff.Add(new XElement(Hh + "font",
                    new XAttribute("id", i.ToString()),
                    new XAttribute("face", ctx.Fonts[i]),
                    new XAttribute("type", "TTF"),
                    new XAttribute("isEmbedded", "0")));
            }
            fontFaces.Add(ff);
        }

        // borderFills — HwpForge / Hancom OWPML 관례를 따라 3개 등록.
        //   id=1: page border (NONE 4면, secPr/pageBorderFill 참조)
        //   id=2: char background (NONE + winBrush face=none alpha=0, charPr 참조)
        //   id=3: table border (SOLID 0.12mm 4면, 표/셀 참조)
        XElement BuildBorder(string borderType) => new(Hh + "borderFill",
            new XAttribute("threeD", "0"),
            new XAttribute("shadow", "0"),
            new XAttribute("centerLine", "NONE"),
            new XAttribute("breakCellSeparateLine", "0"),
            new XElement(Hh + "slash",     new XAttribute("type", "NONE"), new XAttribute("Crooked", "0"), new XAttribute("isCounter", "0")),
            new XElement(Hh + "backSlash", new XAttribute("type", "NONE"), new XAttribute("Crooked", "0"), new XAttribute("isCounter", "0")),
            new XElement(Hh + "leftBorder",   new XAttribute("type", borderType), new XAttribute("width", "0.12 mm"), new XAttribute("color", "#000000")),
            new XElement(Hh + "rightBorder",  new XAttribute("type", borderType), new XAttribute("width", "0.12 mm"), new XAttribute("color", "#000000")),
            new XElement(Hh + "topBorder",    new XAttribute("type", borderType), new XAttribute("width", "0.12 mm"), new XAttribute("color", "#000000")),
            new XElement(Hh + "bottomBorder", new XAttribute("type", borderType), new XAttribute("width", "0.12 mm"), new XAttribute("color", "#000000")),
            new XElement(Hh + "diagonal",     new XAttribute("type", "SOLID"),    new XAttribute("width", "0.1 mm"),  new XAttribute("color", "#000000")));

        var bf1 = BuildBorder("NONE");  bf1.SetAttributeValue("id", "1");
        var bf2 = BuildBorder("NONE");  bf2.SetAttributeValue("id", "2");
        // 문자 배경: NONE 4면 + winBrush (face=none, alpha=0).
        bf2.Add(new XElement(Hc + "fillBrush",
            new XElement(Hc + "winBrush",
                new XAttribute("faceColor",  "none"),
                new XAttribute("hatchColor", "#FF000000"),
                new XAttribute("alpha",      "0"))));
        var bf3 = BuildBorder("SOLID"); bf3.SetAttributeValue("id", "3");

        // 동적 등록된 custom borderFill (표/셀 색·두께·배경) 추가.
        XElement BuildCustomBorder(int id, WriteContext.BorderFillSpec spec)
        {
            var bf = new XElement(Hh + "borderFill",
                new XAttribute("id", id.ToString()),
                new XAttribute("threeD", "0"),
                new XAttribute("shadow", "0"),
                new XAttribute("centerLine", "NONE"),
                new XAttribute("breakCellSeparateLine", "0"),
                new XElement(Hh + "slash",     new XAttribute("type", "NONE"), new XAttribute("Crooked", "0"), new XAttribute("isCounter", "0")),
                new XElement(Hh + "backSlash", new XAttribute("type", "NONE"), new XAttribute("Crooked", "0"), new XAttribute("isCounter", "0")),
                new XElement(Hh + "leftBorder",   new XAttribute("type", spec.LeftType),   new XAttribute("width", spec.LeftWidth),   new XAttribute("color", spec.LeftColor)),
                new XElement(Hh + "rightBorder",  new XAttribute("type", spec.RightType),  new XAttribute("width", spec.RightWidth),  new XAttribute("color", spec.RightColor)),
                new XElement(Hh + "topBorder",    new XAttribute("type", spec.TopType),    new XAttribute("width", spec.TopWidth),    new XAttribute("color", spec.TopColor)),
                new XElement(Hh + "bottomBorder", new XAttribute("type", spec.BottomType), new XAttribute("width", spec.BottomWidth), new XAttribute("color", spec.BottomColor)),
                new XElement(Hh + "diagonal",     new XAttribute("type", "SOLID"),         new XAttribute("width", "0.1 mm"),         new XAttribute("color", "#000000")));
            if (!string.IsNullOrEmpty(spec.FillColor))
            {
                bf.Add(new XElement(Hc + "fillBrush",
                    new XElement(Hc + "winBrush",
                        new XAttribute("faceColor",  spec.FillColor),
                        new XAttribute("hatchColor", "#000000"),
                        new XAttribute("alpha",      "0"))));
            }
            return bf;
        }

        int totalBorderFills = 3 + ctx.CustomBorderFills.Count;
        var borderFills = new XElement(Hh + "borderFills",
            new XAttribute("itemCnt", totalBorderFills.ToString()),
            bf1, bf2, bf3);
        for (int i = 0; i < ctx.CustomBorderFills.Count; i++)
            borderFills.Add(BuildCustomBorder(i + WriteContext.FirstCustomBorderFillId, ctx.CustomBorderFills[i]));

        // charProperties — one entry per unique RunStyle
        var charProps = new XElement(Hh + "charProperties",
            new XAttribute("itemCnt", ctx.RunStyles.Count.ToString()));
        for (int id = 0; id < ctx.RunStyles.Count; id++)
            charProps.Add(BuildDynamicCharPr(id, ctx.RunStyles[id], ctx));

        // tabProperties — one default entry with no custom tab stops
        var tabProperties = new XElement(Hh + "tabProperties",
            new XAttribute("itemCnt", "1"),
            new XElement(Hh + "tabPr",
                new XAttribute("id", "0"),
                new XAttribute("autoTabLeft", "0"),
                new XAttribute("autoTabRight", "0")));

        // paraProperties — one entry per unique ParagraphStyle
        var paraProps = new XElement(Hh + "paraProperties",
            new XAttribute("itemCnt", ctx.ParaStyles.Count.ToString()));
        for (int id = 0; id < ctx.ParaStyles.Count; id++)
            paraProps.Add(BuildDynamicParaPr(id, ctx.ParaStyles[id]));

        // numberings and bullets — empty lists
        var numberings = new XElement(Hh + "numberings", new XAttribute("itemCnt", "0"));
        var bullets     = new XElement(Hh + "bullets",    new XAttribute("itemCnt", "0"));

        // styles: 0=바탕글(Normal), 1~6=개요 1~6(Heading1~6)
        // Real Hancom uses langID (not langIDRef).
        var styles = new XElement(Hh + "styles", new XAttribute("itemCnt", "7"));
        styles.Add(new XElement(Hh + "style",
            new XAttribute("id", "0"), new XAttribute("type", "PARA"),
            new XAttribute("name", "바탕글"), new XAttribute("engName", "Normal"),
            new XAttribute("paraPrIDRef", "0"), new XAttribute("charPrIDRef", "0"),
            new XAttribute("nextStyleIDRef", "0"), new XAttribute("langID", "1042"),
            new XAttribute("lockForm", "0")));
        for (int level = 1; level <= 6; level++)
        {
            styles.Add(new XElement(Hh + "style",
                new XAttribute("id", level.ToString()), new XAttribute("type", "PARA"),
                new XAttribute("name", $"개요 {level}"), new XAttribute("engName", $"Outline {level}"),
                new XAttribute("paraPrIDRef", "0"), new XAttribute("charPrIDRef", "0"),
                new XAttribute("nextStyleIDRef", "0"), new XAttribute("langID", "1042"),
                new XAttribute("lockForm", "0")));
        }

        // KS X 6101 §5 order: fontfaces → borderFills → charProperties → tabProperties
        //   → paraProperties → numberings → bullets → styles
        // 주의: 실 한컴 파일에는 hh:binDataList 가 없다. binaryItemIDRef 는
        // content.hpf 의 opf:item id (예: "image1") 로 직접 해석된다.
        // binDataList 를 추가하면 한컴이 파일을 거부하므로 생략.
        var refList = new XElement(Hh + "refList",
            fontFaces, borderFills, charProps, tabProperties, paraProps,
            numberings, bullets, styles);

        // Real Hancom: targetProgram="HWP201X" with <hh:layoutCompatibility/> child.
        var compatDoc = new XElement(Hh + "compatibleDocument",
            new XAttribute("targetProgram", "HWP201X"),
            new XElement(Hh + "layoutCompatibility"));

        // docOption and trackchageConfig required by real Hancom format.
        var docOption = new XElement(Hh + "docOption",
            new XElement(Hh + "linkinfo",
                new XAttribute("path", ""),
                new XAttribute("pageInherit", "0"),
                new XAttribute("footnoteInherit", "0")));

        var trackChange = new XElement(Hh + "trackchageConfig",
            new XAttribute("flags", "56"));

        // hh:head version="1.4" (real Hancom value, not "1.5" which was wrong).
        // Order per real Hancom: beginNum → refList → compatibleDocument → docOption → trackchageConfig.
        var nsAttrs = AllNamespaceAttrs();
        var head = new XElement(Hh + "head");
        head.Add(new XAttribute("version", "1.4"));
        head.Add(new XAttribute("secCnt", sectionCount.ToString()));
        foreach (var a in nsAttrs) head.Add(a);
        head.Add(new XElement(Hh + "beginNum",
            new XAttribute("page", "1"), new XAttribute("footnote", "1"),
            new XAttribute("endnote", "1"), new XAttribute("pic", "1"),
            new XAttribute("tbl", "1"), new XAttribute("equation", "1")));
        head.Add(refList);
        head.Add(compatDoc);
        head.Add(docOption);
        head.Add(trackChange);

        WriteXml(archive, HwpxPaths.HeaderXml, new XDocument(new XDeclaration("1.0", "utf-8", null), head));
    }

    private XElement BuildDynamicCharPr(int id, RunStyle s, WriteContext ctx)
    {
        long heightUnit = (long)Math.Round(s.FontSizePt * 100);
        if (heightUnit <= 0) heightUnit = 1100; // fallback 11pt

        var el = new XElement(Hh + "charPr",
            new XAttribute("id", id.ToString()),
            new XAttribute("height", heightUnit.ToString()),
            new XAttribute("textColor", s.Foreground is { } fg ? fg.ToHex() : "#000000"),
            new XAttribute("shadeColor", s.Background is { } bg ? bg.ToHex() : "#FFFFFF"),
            new XAttribute("useFontSpace", "0"),
            new XAttribute("useKerning", "0"),
            new XAttribute("symMark", "NONE"),
            new XAttribute("borderFillIDRef", "2")); // char background border (id=2)

        if (s.Bold)          el.Add(new XElement(Hh + "bold"));
        if (s.Italic)        el.Add(new XElement(Hh + "italic"));
        if (s.Underline)     el.Add(new XElement(Hh + "underline",
                                 new XAttribute("type", "BOTTOM"),
                                 new XAttribute("shape", "SOLID"),
                                 new XAttribute("color", "#000000")));
        if (s.Strikethrough) el.Add(new XElement(Hh + "strikeout",
                                 new XAttribute("shape", "SOLID"),
                                 new XAttribute("color", "#000000")));
        if (s.Overline)      el.Add(new XElement(Hh + "overline",
                                 new XAttribute("type", "TOP"),
                                 new XAttribute("shape", "SOLID"),
                                 new XAttribute("color", "#000000")));
        if (s.Superscript)   el.Add(new XElement(Hh + "supScript"));
        if (s.Subscript)     el.Add(new XElement(Hh + "subScript"));

        var fontId = ctx.FontId(s.FontFamily).ToString();
        el.Add(new XElement(Hh + "fontRef",
            new XAttribute("hangul",   fontId),
            new XAttribute("latin",    fontId),
            new XAttribute("hanja",    fontId),
            new XAttribute("japanese", fontId),
            new XAttribute("other",    fontId),
            new XAttribute("symbol",   fontId),
            new XAttribute("user",     fontId)));

        el.Add(new XElement(Hh + "ratio",   new XAttribute("value", "100")));
        el.Add(new XElement(Hh + "spacing", new XAttribute("value", "0")));
        el.Add(new XElement(Hh + "relSz",   new XAttribute("size", "100")));
        el.Add(new XElement(Hh + "offset",  new XAttribute("value", "0")));

        return el;
    }

    private XElement BuildDynamicParaPr(int id, ParagraphStyle s)
    {
        var alignStr = s.Alignment switch
        {
            Alignment.Center                            => "CENTER",
            Alignment.Right                             => "RIGHT",
            Alignment.Justify or Alignment.Distributed => "JUSTIFY",
            _                                           => "LEFT",
        };

        long lineSpacingPct = (long)Math.Round(s.LineHeightFactor * 100);
        if (lineSpacingPct <= 0) lineSpacingPct = 160;

        long indentHwp = UnitConverter.MmToHwpUnit(s.IndentFirstLineMm);
        long leftHwp   = UnitConverter.MmToHwpUnit(s.IndentLeftMm     );
        long rightHwp  = UnitConverter.MmToHwpUnit(s.IndentRightMm    );
        long prevHwp   = (long)Math.Round(s.SpaceBeforePt * 100);
        long nextHwp   = (long)Math.Round(s.SpaceAfterPt  * 100);

        // Real Hancom wraps margin + lineSpacing in hp:switch/hp:case/hp:default
        // to support both HwpUnitChar (character-unit indents) and legacy HWPUNIT modes.
        var marginElem = new XElement(Hh + "margin",
            new XElement(Hc + "intent", new XAttribute("value", indentHwp.ToString()), new XAttribute("unit", "HWPUNIT")),
            new XElement(Hc + "left",   new XAttribute("value", leftHwp.ToString()),   new XAttribute("unit", "HWPUNIT")),
            new XElement(Hc + "right",  new XAttribute("value", rightHwp.ToString()),  new XAttribute("unit", "HWPUNIT")),
            new XElement(Hc + "prev",   new XAttribute("value", prevHwp.ToString()),   new XAttribute("unit", "HWPUNIT")),
            new XElement(Hc + "next",   new XAttribute("value", nextHwp.ToString()),   new XAttribute("unit", "HWPUNIT")));

        var lineSpacingElem = new XElement(Hh + "lineSpacing",
            new XAttribute("type", "PERCENT"),
            new XAttribute("value", lineSpacingPct.ToString()),
            new XAttribute("unit", "HWPUNIT"));

        var switchElem = new XElement(Hp + "switch",
            new XElement(Hp + "case",
                new XAttribute(Hp + "required-namespace", HwpUnitChar.NamespaceName),
                new XElement(Hh + "margin",
                    new XElement(Hc + "intent", new XAttribute("value", indentHwp.ToString()), new XAttribute("unit", "HWPUNIT")),
                    new XElement(Hc + "left",   new XAttribute("value", leftHwp.ToString()),   new XAttribute("unit", "HWPUNIT")),
                    new XElement(Hc + "right",  new XAttribute("value", rightHwp.ToString()),  new XAttribute("unit", "HWPUNIT")),
                    new XElement(Hc + "prev",   new XAttribute("value", prevHwp.ToString()),   new XAttribute("unit", "HWPUNIT")),
                    new XElement(Hc + "next",   new XAttribute("value", nextHwp.ToString()),   new XAttribute("unit", "HWPUNIT"))),
                new XElement(Hh + "lineSpacing",
                    new XAttribute("type", "PERCENT"),
                    new XAttribute("value", lineSpacingPct.ToString()),
                    new XAttribute("unit", "HWPUNIT"))),
            new XElement(Hp + "default",
                new XElement(Hh + "margin",
                    new XElement(Hc + "intent", new XAttribute("value", indentHwp.ToString()), new XAttribute("unit", "HWPUNIT")),
                    new XElement(Hc + "left",   new XAttribute("value", leftHwp.ToString()),   new XAttribute("unit", "HWPUNIT")),
                    new XElement(Hc + "right",  new XAttribute("value", rightHwp.ToString()),  new XAttribute("unit", "HWPUNIT")),
                    new XElement(Hc + "prev",   new XAttribute("value", prevHwp.ToString()),   new XAttribute("unit", "HWPUNIT")),
                    new XElement(Hc + "next",   new XAttribute("value", nextHwp.ToString()),   new XAttribute("unit", "HWPUNIT"))),
                new XElement(Hh + "lineSpacing",
                    new XAttribute("type", "PERCENT"),
                    new XAttribute("value", lineSpacingPct.ToString()),
                    new XAttribute("unit", "HWPUNIT"))));

        return new XElement(Hh + "paraPr",
            new XAttribute("id", id.ToString()),
            new XAttribute("tabPrIDRef", "0"),
            new XAttribute("condense", "0"),
            new XAttribute("fontLineHeight", "0"),
            new XAttribute("snapToGrid", "1"),
            new XAttribute("suppressLineNumbers", "0"),
            new XAttribute("checked", "0"),
            new XElement(Hh + "align",
                new XAttribute("horizontal", alignStr),
                new XAttribute("vertical", "BASELINE")),
            new XElement(Hh + "heading",
                new XAttribute("type", "NONE"),
                new XAttribute("idRef", "0"),
                new XAttribute("level", "0")),
            // Real Hancom: lineWrap="BREAK" not columnBreakBefore="0"
            new XElement(Hh + "breakSetting",
                new XAttribute("breakLatinWord", "KEEP_WORD"),
                new XAttribute("breakNonLatinWord", "KEEP_WORD"),
                new XAttribute("widowOrphan", "0"),
                new XAttribute("keepWithNext", "0"),
                new XAttribute("keepLines", "0"),
                new XAttribute("pageBreakBefore", "0"),
                new XAttribute("lineWrap", "BREAK")),
            new XElement(Hh + "autoSpacing",
                new XAttribute("eAsianEng", "0"),
                new XAttribute("eAsianNum", "0")),
            switchElem,
            new XElement(Hh + "border",
                new XAttribute("borderFillIDRef", "1"), // 1-indexed
                new XAttribute("offsetLeft", "0"),
                new XAttribute("offsetRight", "0"),
                new XAttribute("offsetTop", "0"),
                new XAttribute("offsetBottom", "0"),
                new XAttribute("connect", "0"),
                new XAttribute("ignoreMargin", "0")));
    }

    // ── section XML ──────────────────────────────────────────────────────────

    private void WriteSectionXml(ZipArchive archive, int sectionIndex, Section section, WriteContext ctx)
    {
        ctx.ResetParaId();

        // Include all namespace declarations on hs:sec root (matches real Hancom format).
        var nsAttrs = AllNamespaceAttrs();
        var sec = new XElement(Hs + "sec");
        foreach (var a in nsAttrs) sec.Add(a);

        // secPr must be the first run of the first paragraph of EVERY section (KS X 6101).
        // Without it Hangul Office rejects the file with "파일을 읽거나 저장하는데 오류가 있습니다".
        bool injectSecPr = true;

        // 도형(ShapeObject) 은 anchored overlay 로 출력되지만 hosting paragraph 의
        // 위치가 anchor 의 페이지를 결정한다. IWPF 문서 순서상 마지막에 두면 host
        // paragraph 가 마지막 페이지로 밀려 도형이 그 페이지에 그려진다.
        // 사용자 보고: 도형이 다음 페이지로 넘어감 — 모든 도형의 anchor 가 끝쪽
        // 페이지에 놓이기 때문.
        // 해결: AnchorPageIndex==0 인 도형을 첫 비-도형 단락 직후로 재배치.
        // 첫 단락은 secPr 을 보유해야 해서 도형으로 대체 안 함 — 두번째 위치부터
        // 도형 삽입. 결과적으로 도형 anchor 는 page 1 첫 단락 다음에 위치.
        var blocks = section.Blocks.ToList();
        var page0Shapes = blocks
            .OfType<ShapeObject>()
            .Where(s => s.AnchorPageIndex == 0)
            .ToList();
        if (page0Shapes.Count > 0 && blocks.Count > page0Shapes.Count)
        {
            foreach (var s in page0Shapes) blocks.Remove(s);
            // 첫 비-도형 블록 직후 (인덱스 1) 에 도형들을 삽입.
            blocks.InsertRange(1, page0Shapes);
        }

        if (blocks.Count == 0)
        {
            var para = BuildEmptyParagraph(ctx);
            if (injectSecPr) PrependSecPrRun(para);
            sec.Add(para);
        }
        else
        {
            foreach (var block in blocks)
                AppendBlock(sec, block, ctx, ref injectSecPr);
        }

        WriteXml(archive, HwpxPaths.SectionXml(sectionIndex),
            new XDocument(new XDeclaration("1.0", "utf-8", null), sec));
    }

    private void AppendBlock(XElement target, Block block, WriteContext ctx, ref bool injectSecPr)
    {
        XElement para = block switch
        {
            Paragraph p     => BuildParagraph(p, ctx),
            Table t         => BuildTableHostingParagraph(t, ctx),
            ImageBlock img  => BuildImageHostingParagraph(img, ctx),
            ShapeObject sh  => BuildShapeHostingParagraph(sh, ctx),
            OpaqueBlock op  => BuildOpaqueHostingParagraph(op, ctx),
            _               => BuildEmptyParagraph(ctx),
        };
        if (injectSecPr) { PrependSecPrRun(para); injectSecPr = false; }
        target.Add(para);
    }

    // Overload for table-cell inner blocks (no secPr injection, uses local dummy).
    private void AppendBlock(XElement target, Block block, WriteContext ctx)
    {
        bool dummy = false;
        AppendBlock(target, block, ctx, ref dummy);
    }

    // Injects the full hp:secPr control run as the very first child of the paragraph.
    // Structure per real Hancom format: secPr + ctrl in same run.
    private void PrependSecPrRun(XElement para)
    {
        var secPr = new XElement(Hp + "secPr",
            new XAttribute("id", ""),
            new XAttribute("textDirection", "HORIZONTAL"),
            new XAttribute("spaceColumns", "1134"),
            new XAttribute("tabStop", "8000"),
            new XAttribute("tabStopVal", "4000"),
            new XAttribute("tabStopUnit", "HWPUNIT"),
            new XAttribute("outlineShapeIDRef", "1"), // 1-indexed, matches borderFill IDs
            new XAttribute("memoShapeIDRef", "0"),
            new XAttribute("textVerticalWidthHead", "0"),
            new XAttribute("masterPageCnt", "0"),
            // Ordered children per real Hancom: grid, startNum, visibility, lineNumberShape,
            // pagePr, footNotePr, endNotePr, pageBorderFill×3
            new XElement(Hp + "grid",
                new XAttribute("lineGrid", "0"),
                new XAttribute("charGrid", "0"),
                new XAttribute("wonggojiFormat", "0")),
            new XElement(Hp + "startNum",
                new XAttribute("pageStartsOn", "BOTH"),
                new XAttribute("page", "0"),
                new XAttribute("pic", "0"),
                new XAttribute("tbl", "0"),
                new XAttribute("equation", "0")),
            new XElement(Hp + "visibility",
                new XAttribute("hideFirstHeader", "0"),
                new XAttribute("hideFirstFooter", "0"),
                new XAttribute("hideFirstMasterPage", "0"),
                new XAttribute("border", "SHOW_ALL"),
                new XAttribute("fill", "SHOW_ALL"),
                new XAttribute("hideFirstPageNum", "0"),
                new XAttribute("hideFirstEmptyLine", "0"),
                new XAttribute("showLineNumber", "0")),
            new XElement(Hp + "lineNumberShape",
                new XAttribute("restartType", "0"),
                new XAttribute("countBy", "0"),
                new XAttribute("distance", "0"),
                new XAttribute("startNumber", "0")),
            new XElement(Hp + "pagePr",
                new XAttribute("landscape", "WIDELY"),
                new XAttribute("width",  A4W.ToString()),
                new XAttribute("height", A4H.ToString()),
                new XAttribute("gutterType", "LEFT_ONLY"),
                new XElement(Hp + "margin",
                    new XAttribute("header", MarginHead.ToString()),
                    new XAttribute("footer", MarginFoot.ToString()),
                    new XAttribute("gutter", "0"),
                    new XAttribute("left",   MarginLeft.ToString()),
                    new XAttribute("right",  MarginRight.ToString()),
                    new XAttribute("top",    MarginTop.ToString()),
                    new XAttribute("bottom", MarginBottom.ToString()))),
            new XElement(Hp + "footNotePr",
                new XElement(Hp + "autoNumFormat",
                    new XAttribute("type", "DIGIT"),
                    new XAttribute("userChar", ""),
                    new XAttribute("prefixChar", ""),
                    new XAttribute("suffixChar", ")"),
                    new XAttribute("supscript", "0")),
                new XElement(Hp + "noteLine",
                    new XAttribute("length", "-1"),
                    new XAttribute("type", "SOLID"),
                    new XAttribute("width", "0.12 mm"),
                    new XAttribute("color", "#000000")),
                new XElement(Hp + "noteSpacing",
                    new XAttribute("betweenNotes", "283"),
                    new XAttribute("belowLine", "567"),
                    new XAttribute("aboveLine", "850")),
                new XElement(Hp + "numbering",
                    new XAttribute("type", "CONTINUOUS"),
                    new XAttribute("newNum", "1")),
                new XElement(Hp + "placement",
                    new XAttribute("place", "EACH_COLUMN"),
                    new XAttribute("beneathText", "0"))),
            new XElement(Hp + "endNotePr",
                new XElement(Hp + "autoNumFormat",
                    new XAttribute("type", "DIGIT"),
                    new XAttribute("userChar", ""),
                    new XAttribute("prefixChar", ""),
                    new XAttribute("suffixChar", ")"),
                    new XAttribute("supscript", "0")),
                new XElement(Hp + "noteLine",
                    new XAttribute("length", "14692344"),
                    new XAttribute("type", "SOLID"),
                    new XAttribute("width", "0.12 mm"),
                    new XAttribute("color", "#000000")),
                new XElement(Hp + "noteSpacing",
                    new XAttribute("betweenNotes", "0"),
                    new XAttribute("belowLine", "567"),
                    new XAttribute("aboveLine", "850")),
                new XElement(Hp + "numbering",
                    new XAttribute("type", "CONTINUOUS"),
                    new XAttribute("newNum", "1")),
                new XElement(Hp + "placement",
                    new XAttribute("place", "END_OF_DOCUMENT"),
                    new XAttribute("beneathText", "0"))),
            // Three pageBorderFill entries: BOTH/EVEN/ODD all referencing borderFill id=1
            new XElement(Hp + "pageBorderFill",
                new XAttribute("type", "BOTH"),
                new XAttribute("borderFillIDRef", "1"),
                new XAttribute("textBorder", "PAPER"),
                new XAttribute("headerInside", "0"),
                new XAttribute("footerInside", "0"),
                new XAttribute("fillArea", "PAPER"),
                new XElement(Hp + "offset",
                    new XAttribute("left", "1417"), new XAttribute("right", "1417"),
                    new XAttribute("top", "1417"), new XAttribute("bottom", "1417"))),
            new XElement(Hp + "pageBorderFill",
                new XAttribute("type", "EVEN"),
                new XAttribute("borderFillIDRef", "1"),
                new XAttribute("textBorder", "PAPER"),
                new XAttribute("headerInside", "0"),
                new XAttribute("footerInside", "0"),
                new XAttribute("fillArea", "PAPER"),
                new XElement(Hp + "offset",
                    new XAttribute("left", "1417"), new XAttribute("right", "1417"),
                    new XAttribute("top", "1417"), new XAttribute("bottom", "1417"))),
            new XElement(Hp + "pageBorderFill",
                new XAttribute("type", "ODD"),
                new XAttribute("borderFillIDRef", "1"),
                new XAttribute("textBorder", "PAPER"),
                new XAttribute("headerInside", "0"),
                new XAttribute("footerInside", "0"),
                new XAttribute("fillArea", "PAPER"),
                new XElement(Hp + "offset",
                    new XAttribute("left", "1417"), new XAttribute("right", "1417"),
                    new XAttribute("top", "1417"), new XAttribute("bottom", "1417"))));

        var ctrl = new XElement(Hp + "ctrl",
            new XElement(Hp + "colPr",
                new XAttribute("id", ""),
                new XAttribute("type", "NEWSPAPER"),
                new XAttribute("layout", "LEFT"),
                new XAttribute("colCount", "1"),
                new XAttribute("sameSz", "1"),
                new XAttribute("sameGap", "0")));

        var secPrRun = new XElement(Hp + "run",
            new XAttribute("charPrIDRef", "0"),
            secPr,
            ctrl);
        para.AddFirst(secPrRun);
    }

    private XElement BuildParagraph(Paragraph p, WriteContext ctx)
    {
        int paraPrID = ctx.ParaStyleId(p.Style);
        int styleID = p.Style.Outline switch
        {
            OutlineLevel.H1 => 1,
            OutlineLevel.H2 => 2,
            OutlineLevel.H3 => 3,
            OutlineLevel.H4 => 4,
            OutlineLevel.H5 => 5,
            OutlineLevel.H6 => 6,
            _               => 0,
        };

        var para = new XElement(Hp + "p",
            new XAttribute("id",          ctx.NextParaId().ToString()),
            new XAttribute("paraPrIDRef", paraPrID.ToString()),
            new XAttribute("styleIDRef",  styleID.ToString()),
            new XAttribute("pageBreak",   p.Style.ForcePageBreakBefore ? "1" : "0"),
            new XAttribute("columnBreak", "0"),
            new XAttribute("merged",      "0"));

        foreach (var run in p.Runs)
            para.Add(BuildRun(run, ctx));
        if (p.Runs.Count == 0)
            para.Add(new XElement(Hp + "run", new XAttribute("charPrIDRef", "0")));

        double maxPt = p.Runs.Count > 0
            ? p.Runs.Max(r => r.Style.FontSizePt > 0 ? r.Style.FontSizePt : 10.0)
            : 10.0;
        double lsFactor = p.Style.LineHeightFactor > 0 ? p.Style.LineHeightFactor : 1.6;

        para.Add(BuildLineseg(maxPt, lsFactor));
        return para;
    }

    private XElement BuildRun(Run run, WriteContext ctx)
    {
        // 각주/미주 참조 run — hp:ctrl 을 품은 run 으로 출력.
        if (run.FootnoteId is { Length: > 0 } fnId)
            return BuildNoteCtrlRun("FOOT_NOTE", fnId, ctx);
        if (run.EndnoteId is { Length: > 0 } enId)
            return BuildNoteCtrlRun("END_NOTE", enId, ctx);

        int charPrID = ctx.RunStyleId(run.Style);
        var elem = new XElement(Hp + "run",
            new XAttribute("charPrIDRef", charPrID.ToString()));
        if (run.Text.Length > 0)
            elem.Add(new XElement(Hp + "t", run.Text));
        return elem;
    }

    private XElement BuildNoteCtrlRun(string ctrlID, string noteId, WriteContext ctx)
    {
        // 각주/미주 내용을 document 에서 찾아 subList 로 직렬화.
        var entries = ctrlID == "FOOT_NOTE"
            ? (IEnumerable<FootnoteEntry>)_pendingFootnotes
            : _pendingEndnotes;
        var entry = entries.FirstOrDefault(e => e.Id == noteId);

        var subListId = ctx.NextSubListId();
        var subList = new XElement(Hp + "subList",
            new XAttribute("id",               subListId.ToString()),
            new XAttribute("textDirection",    "HORIZONTAL"),
            new XAttribute("lineWrap",         "BREAK"),
            new XAttribute("vertAlign",        "NORMAL"),
            new XAttribute("linkListIDRef",    "0"),
            new XAttribute("linkListNextIDRef","0"),
            new XAttribute("textWidth",        "0"),
            new XAttribute("textHeight",       "0"));

        if (entry is not null)
        {
            ctx.ResetParaId();
            foreach (var block in entry.Blocks)
            {
                if (block is Paragraph p)
                    subList.Add(BuildParagraph(p, ctx));
                else
                    subList.Add(BuildEmptyParagraph(ctx));
            }
        }
        else
        {
            subList.Add(BuildEmptyParagraph(ctx));
        }

        var ctrl = new XElement(Hp + "ctrl",
            new XAttribute("id",           ctx.NextCtrlId().ToString()),
            new XAttribute("ctrlID",       ctrlID),
            new XAttribute("numberingType","BOTH_ALIGN"),
            new XAttribute("textRef",      "0"),
            subList);

        return new XElement(Hp + "run",
            new XAttribute("charPrIDRef", "0"),
            ctrl,
            new XElement(Hp + "t"));
    }

    private XElement BuildEmptyParagraph(WriteContext ctx)
    {
        var para = new XElement(Hp + "p",
            new XAttribute("id",          ctx.NextParaId().ToString()),
            new XAttribute("paraPrIDRef", "0"),
            new XAttribute("styleIDRef",  "0"),
            new XAttribute("pageBreak",   "0"),
            new XAttribute("columnBreak", "0"),
            new XAttribute("merged",      "0"),
            new XElement(Hp + "run", new XAttribute("charPrIDRef", "0")));
        para.Add(BuildLineseg(10.0, 1.6));
        return para;
    }

    // Build lineseg with new (real Hancom) attribute names.
    // vertpos=0 (simplified; Hangul recalculates), others calculated from font metrics.
    private static XElement BuildLineseg(double fontPt, double lsFactor)
    {
        long textH   = (long)Math.Round(fontPt * 100);
        if (textH <= 0) textH = 1000;
        long vertSize  = textH;
        long baseline  = (long)Math.Round(textH * 0.85);
        long spacing   = (long)Math.Round(textH * (lsFactor - 1.0));

        return new XElement(Hp + "linesegarray",
            new XElement(Hp + "lineseg",
                new XAttribute("textpos",   "0"),
                new XAttribute("vertpos",   "0"),
                new XAttribute("vertsize",  vertSize.ToString()),
                new XAttribute("textheight", textH.ToString()),
                new XAttribute("baseline",  baseline.ToString()),
                new XAttribute("spacing",   spacing.ToString()),
                new XAttribute("horzpos",   "0"),
                new XAttribute("horzsize",  ContentWidth.ToString()),
                new XAttribute("flags",     "393216")));
    }

    // ── opaque island ────────────────────────────────────────────────────────

    private XElement BuildOpaqueHostingParagraph(OpaqueBlock op, WriteContext ctx)
    {
        // hwpx 포맷 + XML 있으면 원본 재출력; 그 외 포맷은 placeholder 단락.
        if (op.Format == "hwpx" && !string.IsNullOrEmpty(op.Xml))
        {
            try
            {
                var shapeElem = XElement.Parse(op.Xml);
                // 보존된 opaque XML 의 *IDRef 가 원본 HWPX 의 charPr/paraPr 테이블을
                // 가리키는데 우리 새 헤더에는 그 id 가 없을 수 있다 — 한컴이 lookup
                // 하다 무한 루프에 빠지므로 모든 ID 참조를 "0" 으로 재작성한다.
                SanitizeIdRefs(shapeElem);
                // 한컴 실파일 패턴: 도형 → 빈 hp:t → 단락 끝에 linesegarray.
                // linesegarray 가 없으면 한컴이 레이아웃 계산 중 무한 루프에 빠진다.
                var run = new XElement(Hp + "run",
                    new XAttribute("charPrIDRef", "0"),
                    shapeElem,
                    new XElement(Hp + "t"));
                var para = new XElement(Hp + "p",
                    new XAttribute("id",          ctx.NextParaId().ToString()),
                    new XAttribute("paraPrIDRef", "0"),
                    new XAttribute("styleIDRef",  "0"),
                    new XAttribute("pageBreak",   "0"),
                    new XAttribute("columnBreak", "0"),
                    new XAttribute("merged",      "0"),
                    run);
                para.Add(BuildLineseg(10.0, 1.6));
                return para;
            }
            catch (System.Xml.XmlException)
            {
                // 파싱 실패 시 placeholder 로 fallback.
            }
        }
        return BuildParagraph(Paragraph.Of(op.DisplayLabel), ctx);
    }

    // *IDRef 들이 원본 헤더의 ID 를 가리키므로 새 헤더에서 빈 id 가 되면 한컴이 거부.
    // 안전하게 0 으로 재작성. (style 손실보다 한컴 호환이 우선)
    private static readonly string[] s_idRefAttrs = [
        "paraPrIDRef", "charPrIDRef", "styleIDRef",
        "borderFillIDRef", "tabPrIDRef",
        "linkListIDRef", "linkListNextIDRef",
        "numberingIDRef", "bulletIDRef",
        "outlineShapeIDRef", "memoShapeIDRef",
        "masterPageIDRef",
    ];

    private static void SanitizeIdRefs(XElement root)
    {
        foreach (var elem in root.DescendantsAndSelf())
        {
            foreach (var name in s_idRefAttrs)
            {
                var attr = elem.Attribute(name);
                if (attr is not null) attr.Value = "0";
            }
        }
    }

    // ── table ────────────────────────────────────────────────────────────────

    private XElement BuildTableHostingParagraph(Table table, WriteContext ctx)
    {
        var run = new XElement(Hp + "run", new XAttribute("charPrIDRef", "0"));
        run.Add(BuildTable(table, ctx));
        run.Add(new XElement(Hp + "t"));

        var para = new XElement(Hp + "p",
            new XAttribute("id",          ctx.NextParaId().ToString()),
            new XAttribute("paraPrIDRef", "0"),
            new XAttribute("styleIDRef",  "0"),
            new XAttribute("pageBreak",   "0"),
            new XAttribute("columnBreak", "0"),
            new XAttribute("merged",      "0"),
            run);
        para.Add(BuildLineseg(10.0, 1.6));
        return para;
    }

    private XElement BuildTable(Table table, WriteContext ctx)
    {
        int rowCount = table.Rows.Count;
        int colCount = table.Columns.Count;
        if (colCount == 0 && rowCount > 0)
            colCount = table.Rows.Max(r => r.Cells.Sum(c => Math.Max(c.ColumnSpan, 1)));
        if (colCount <= 0) colCount = 1;

        // 컬럼 너비 (HWPUNIT) — 정의 없으면 ContentWidth 균등.
        long[] colWidths = new long[colCount];
        bool colWidthsDefined = table.Columns.Count == colCount &&
                                table.Columns.All(c => c.WidthMm > 0);
        if (colWidthsDefined)
            for (int i = 0; i < colCount; i++)
                colWidths[i] = UnitConverter.MmToHwpUnit(table.Columns[i].WidthMm);
        else
        {
            long even = ContentWidth / colCount;
            for (int i = 0; i < colCount; i++) colWidths[i] = even;
        }
        long tableWidth = colWidths.Sum();
        if (tableWidth <= 0) tableWidth = ContentWidth;

        long[] rowHeights = new long[rowCount];
        for (int i = 0; i < rowCount; i++)
            rowHeights[i] = table.Rows[i].HeightMm > 0
                ? UnitConverter.MmToHwpUnit(table.Rows[i].HeightMm)
                : 1000;

        // 한컴 기본 마진 (HwpForge DEFAULT_CELL_MARGIN/OUT_MARGIN)
        const long CellMarginLeftRight = 510;
        const long CellMarginTopBottom = 141;
        const long OuterMarginAll      = 283;

        // 표 외곽선·배경색을 동적 borderFill 로 등록 (모델값 반영).
        // pt 두께를 mm 로 (1 pt ≈ 0.3528 mm). 한컴이 인식하는 표준 너비 집합으로 snap.
        // "0.71 mm" 같은 비표준 값은 한컴이 기본 thin (0.12 mm) 으로 fallback.
        static string BorderWidthMm(double pt)
        {
            if (pt <= 0) return "0.12 mm";
            double mm = pt * 0.3527777778;
            double[] standard = { 0.1, 0.12, 0.15, 0.2, 0.25, 0.3, 0.4, 0.5, 0.6, 0.7, 1.0, 1.5, 2.0, 3.0, 4.0, 5.0 };
            double best = standard[0];
            double bestDiff = Math.Abs(best - mm);
            foreach (var s in standard)
            {
                double d = Math.Abs(s - mm);
                if (d < bestDiff) { best = s; bestDiff = d; }
            }
            return $"{best:0.##} mm";
        }
        string tableBorderColor = string.IsNullOrEmpty(table.BorderColor) ? "#000000" : table.BorderColor!;
        string tableBorderWidth = BorderWidthMm(table.BorderThicknessPt);
        // 표 자체의 borderFillIDRef — 4면 동일 (외곽선 spec).
        int tableBorderFillId = ctx.RegisterCustomBorderFill(WriteContext.BorderFillSpec.Uniform(
            "SOLID", tableBorderColor, tableBorderWidth, table.BackgroundColor));

        // 외곽 속성 — CTableType (CAbstractShapeObjectType 상속) 정의대로.
        var wtbl = new XElement(Hp + "tbl",
            new XAttribute("id",              ctx.NextObjId().ToString()),
            new XAttribute("zOrder",          ctx.NextZOrder().ToString()),
            new XAttribute("numberingType",   "TABLE"),
            new XAttribute("textWrap",        "TOP_AND_BOTTOM"),
            new XAttribute("textFlow",        "BOTH_SIDES"),
            new XAttribute("lock",            "0"),
            new XAttribute("dropcapstyle",    "None"),
            new XAttribute("pageBreak",       "TABLE"),
            new XAttribute("repeatHeader",    "1"),
            new XAttribute("rowCnt",          rowCount.ToString()),
            new XAttribute("colCnt",          colCount.ToString()),
            new XAttribute("cellSpacing",     "0"),
            new XAttribute("borderFillIDRef", tableBorderFillId.ToString()),
            new XAttribute("noAdjust",        "0"));

        wtbl.Add(new XElement(Hp + "sz",
            new XAttribute("width",       tableWidth.ToString()),
            new XAttribute("widthRelTo",  "ABSOLUTE"),
            new XAttribute("height",      "0"),
            new XAttribute("heightRelTo", "ABSOLUTE"),
            new XAttribute("protect",     "0")));
        // 표 위치 — anchored (treatAsChar=0, vertRelTo=PARA, horzRelTo=COLUMN).
        wtbl.Add(new XElement(Hp + "pos",
            new XAttribute("treatAsChar",     "0"),
            new XAttribute("affectLSpacing",  "0"),
            new XAttribute("flowWithText",    "1"),
            new XAttribute("allowOverlap",    "0"),
            new XAttribute("holdAnchorAndSO", "0"),
            new XAttribute("vertRelTo",       "PARA"),
            new XAttribute("horzRelTo",       "COLUMN"),
            new XAttribute("vertAlign",       "TOP"),
            new XAttribute("horzAlign",       "LEFT"),
            new XAttribute("vertOffset",      "0"),
            new XAttribute("horzOffset",      "0")));
        wtbl.Add(new XElement(Hp + "outMargin",
            new XAttribute("left",   OuterMarginAll.ToString()),
            new XAttribute("right",  OuterMarginAll.ToString()),
            new XAttribute("top",    OuterMarginAll.ToString()),
            new XAttribute("bottom", OuterMarginAll.ToString())));
        wtbl.Add(new XElement(Hp + "inMargin",
            new XAttribute("left",   CellMarginLeftRight.ToString()),
            new XAttribute("right",  CellMarginLeftRight.ToString()),
            new XAttribute("top",    CellMarginTopBottom.ToString()),
            new XAttribute("bottom", CellMarginTopBottom.ToString())));

        for (int r = 0; r < rowCount; r++)
        {
            var row = table.Rows[r];
            var wrow = new XElement(Hp + "tr");
            int c = 0;
            foreach (var cell in row.Cells)
            {
                int colSpan = Math.Max(cell.ColumnSpan, 1);
                int rowSpan = Math.Max(cell.RowSpan, 1);

                long cellW = 0;
                for (int k = 0; k < colSpan && c + k < colCount; k++) cellW += colWidths[c + k];
                if (cell.WidthMm > 0)
                    cellW = UnitConverter.MmToHwpUnit(cell.WidthMm);
                if (cellW <= 0) cellW = ContentWidth / Math.Max(colCount, 1);

                long cellH = 0;
                for (int k = 0; k < rowSpan && r + k < rowCount; k++) cellH += rowHeights[r + k];
                if (cellH <= 0) cellH = 1000;

                // 셀 외곽선·배경색 — 셀이 명시 안하면 표 외곽선 설정 상속.
                // 외곽쪽 면은 표 외곽선 spec, 그 외는 cell inner spec → 표 외곽선 가시화.
                string innerColor = !string.IsNullOrEmpty(cell.BorderColor) ? cell.BorderColor!
                                  : tableBorderColor;
                string innerWidth = cell.BorderThicknessPt > 0 ? BorderWidthMm(cell.BorderThicknessPt)
                                  : tableBorderWidth;
                string? cellFill  = !string.IsNullOrEmpty(cell.BackgroundColor) ? cell.BackgroundColor
                                  : table.BackgroundColor;
                bool topEdge    = r == 0;
                bool bottomEdge = r + rowSpan >= rowCount;
                bool leftEdge   = c == 0;
                bool rightEdge  = c + colSpan >= colCount;
                int cellBorderFillId = ctx.RegisterCustomBorderFill(new WriteContext.BorderFillSpec(
                    "SOLID", topEdge    ? tableBorderColor : innerColor, topEdge    ? tableBorderWidth : innerWidth,
                    "SOLID", bottomEdge ? tableBorderColor : innerColor, bottomEdge ? tableBorderWidth : innerWidth,
                    "SOLID", leftEdge   ? tableBorderColor : innerColor, leftEdge   ? tableBorderWidth : innerWidth,
                    "SOLID", rightEdge  ? tableBorderColor : innerColor, rightEdge  ? tableBorderWidth : innerWidth,
                    cellFill));

                var wcell = new XElement(Hp + "tc",
                    new XAttribute("name",            string.Empty),
                    new XAttribute("header",          row.IsHeader ? "1" : "0"),
                    new XAttribute("hasMargin",       "0"),
                    new XAttribute("protect",         "0"),
                    new XAttribute("editable",        "0"),
                    new XAttribute("dirty",           "0"),
                    new XAttribute("borderFillIDRef", cellBorderFillId.ToString()));

                // hp:subList — 셀 본문 컨테이너. 한컴 필수 속성 전체 포함.
                var subList = new XElement(Hp + "subList",
                    new XAttribute("id",                ""),
                    new XAttribute("textDirection",     "HORIZONTAL"),
                    new XAttribute("lineWrap",          "BREAK"),
                    new XAttribute("vertAlign",         "CENTER"),
                    new XAttribute("linkListIDRef",     "0"),
                    new XAttribute("linkListNextIDRef", "0"),
                    new XAttribute("textWidth",         "0"),
                    new XAttribute("textHeight",        "0"),
                    new XAttribute("hasTextRef",        "0"),
                    new XAttribute("hasNumRef",         "0"));
                if (cell.Blocks.Count == 0)
                    subList.Add(BuildEmptyParagraph(ctx));
                else
                    foreach (var inner in cell.Blocks)
                        AppendBlock(subList, inner, ctx);

                wcell.Add(subList);
                wcell.Add(new XElement(Hp + "cellAddr",
                    new XAttribute("colAddr", c.ToString()),
                    new XAttribute("rowAddr", r.ToString())));
                wcell.Add(new XElement(Hp + "cellSpan",
                    new XAttribute("colSpan", colSpan.ToString()),
                    new XAttribute("rowSpan", rowSpan.ToString())));
                wcell.Add(new XElement(Hp + "cellSz",
                    new XAttribute("width",  cellW.ToString()),
                    new XAttribute("height", cellH.ToString())));
                wcell.Add(new XElement(Hp + "cellMargin",
                    new XAttribute("left",   CellMarginLeftRight.ToString()),
                    new XAttribute("right",  CellMarginLeftRight.ToString()),
                    new XAttribute("top",    CellMarginTopBottom.ToString()),
                    new XAttribute("bottom", CellMarginTopBottom.ToString())));

                wrow.Add(wcell);
                c += colSpan;
            }
            wtbl.Add(wrow);
        }
        return wtbl;
    }

    // ── image ────────────────────────────────────────────────────────────────

    private XElement BuildImageHostingParagraph(ImageBlock image, WriteContext ctx)
    {
        var run = new XElement(Hp + "run", new XAttribute("charPrIDRef", "0"));
        if (image.Data.Length > 0)
            run.Add(BuildPicture(image, ctx));
        run.Add(new XElement(Hp + "t"));

        var para = new XElement(Hp + "p",
            new XAttribute("id",          ctx.NextParaId().ToString()),
            new XAttribute("paraPrIDRef", "0"),
            new XAttribute("styleIDRef",  "0"),
            new XAttribute("pageBreak",   "0"),
            new XAttribute("columnBreak", "0"),
            new XAttribute("merged",      "0"),
            run);
        para.Add(BuildLineseg(10.0, 1.6));
        return para;
    }

    private XElement BuildPicture(ImageBlock image, WriteContext ctx)
    {
        var binId = ctx.AddImage(image.Data, image.MediaType);
        long w = UnitConverter.MmToHwpUnit((image.WidthMm  > 0 ? image.WidthMm  : 80));
        long h = UnitConverter.MmToHwpUnit((image.HeightMm > 0 ? image.HeightMm : 60));

        // 외곽 속성과 sz/pos/outMargin 이 없으면 한컴 오피스가 거부.
        // treatAsChar="1" 로 인라인 배치 — 위치 계산 불필요.
        return new XElement(Hp + "pic",
            new XAttribute("id",            ctx.NextObjId().ToString()),
            new XAttribute("zOrder",        ctx.NextZOrder().ToString()),
            new XAttribute("numberingType", "PICTURE"),
            new XAttribute("textWrap",      "TOP_AND_BOTTOM"),
            new XAttribute("textFlow",      "BOTH_SIDES"),
            new XAttribute("lock",          "0"),
            new XAttribute("dropcapstyle",  "None"),
            new XAttribute("href",          ""),
            new XAttribute("groupLevel",    "0"),
            new XAttribute("instid",        ctx.NextInstId().ToString()),
            new XAttribute("reverse",       "0"),

            new XElement(Hp + "offset", new XAttribute("x", "0"), new XAttribute("y", "0")),
            new XElement(Hp + "orgSz",  new XAttribute("width", w.ToString()), new XAttribute("height", h.ToString())),
            new XElement(Hp + "curSz",  new XAttribute("width", w.ToString()), new XAttribute("height", h.ToString())),
            new XElement(Hp + "flip",   new XAttribute("horizontal", "0"), new XAttribute("vertical", "0")),
            new XElement(Hp + "rotationInfo",
                new XAttribute("angle",       "0"),
                new XAttribute("centerX",     (w / 2).ToString()),
                new XAttribute("centerY",     (h / 2).ToString()),
                new XAttribute("rotateimage", "1")),
            new XElement(Hp + "renderingInfo",
                new XElement(Hc + "transMatrix",
                    new XAttribute("e1", "1"), new XAttribute("e2", "0"), new XAttribute("e3", "0"),
                    new XAttribute("e4", "0"), new XAttribute("e5", "1"), new XAttribute("e6", "0")),
                new XElement(Hc + "scaMatrix",
                    new XAttribute("e1", "1"), new XAttribute("e2", "0"), new XAttribute("e3", "0"),
                    new XAttribute("e4", "0"), new XAttribute("e5", "1"), new XAttribute("e6", "0")),
                new XElement(Hc + "rotMatrix",
                    new XAttribute("e1", "1"), new XAttribute("e2", "0"), new XAttribute("e3", "0"),
                    new XAttribute("e4", "0"), new XAttribute("e5", "1"), new XAttribute("e6", "0"))),
            // HwpForge ground truth 순서: hc:img → imgRect → imgClip → inMargin → imgDim → effects.
            // img 는 hc: 네임스페이스 (이전엔 hp: 로 잘못 사용 → 한컴 거부 원인).
            // Picture 는 lineShape/shadow/fillBrush 없음 (drawing object 아님).
            new XElement(Hc + "img",
                new XAttribute("binaryItemIDRef", binId),
                new XAttribute("bright",   "0"),
                new XAttribute("contrast", "0"),
                new XAttribute("effect",   "REAL_PIC"),
                new XAttribute("alpha",    "0")),
            new XElement(Hp + "imgRect",
                new XElement(Hc + "pt0", new XAttribute("x", "0"),         new XAttribute("y", "0")),
                new XElement(Hc + "pt1", new XAttribute("x", w.ToString()), new XAttribute("y", "0")),
                new XElement(Hc + "pt2", new XAttribute("x", w.ToString()), new XAttribute("y", h.ToString())),
                new XElement(Hc + "pt3", new XAttribute("x", "0"),         new XAttribute("y", h.ToString()))),
            new XElement(Hp + "imgClip",
                new XAttribute("left",   "0"),         new XAttribute("top",    "0"),
                new XAttribute("right",  w.ToString()), new XAttribute("bottom", h.ToString())),
            new XElement(Hp + "inMargin",
                new XAttribute("left", "0"), new XAttribute("right",  "0"),
                new XAttribute("top",  "0"), new XAttribute("bottom", "0")),
            new XElement(Hp + "imgDim",
                new XAttribute("dimwidth",  w.ToString()),
                new XAttribute("dimheight", h.ToString())),
            new XElement(Hp + "effects"),
            new XElement(Hp + "sz",
                new XAttribute("width",       w.ToString()),
                new XAttribute("widthRelTo",  "ABSOLUTE"),
                new XAttribute("height",      h.ToString()),
                new XAttribute("heightRelTo", "ABSOLUTE"),
                new XAttribute("protect",     "0")),
            new XElement(Hp + "pos",
                new XAttribute("treatAsChar",     "1"),
                new XAttribute("affectLSpacing",  "0"),
                new XAttribute("flowWithText",    "1"),
                new XAttribute("allowOverlap",    "0"),
                new XAttribute("holdAnchorAndSO", "0"),
                new XAttribute("vertRelTo",       "PARA"),
                new XAttribute("horzRelTo",       "PARA"),
                new XAttribute("vertAlign",       "TOP"),
                new XAttribute("horzAlign",       "LEFT"),
                new XAttribute("vertOffset",      "0"),
                new XAttribute("horzOffset",      "0")),
            new XElement(Hp + "outMargin",
                new XAttribute("left", "0"), new XAttribute("right",  "0"),
                new XAttribute("top",  "0"), new XAttribute("bottom", "0")));
    }

    // ── shape (PolyDonky 네이티브 ShapeObject → HWPX hp:line/hp:rect/...) ────────

    private XElement BuildShapeHostingParagraph(ShapeObject shape, WriteContext ctx)
    {
        var run = new XElement(Hp + "run", new XAttribute("charPrIDRef", "0"));
        run.Add(BuildShape(shape, ctx));
        run.Add(new XElement(Hp + "t"));

        var para = new XElement(Hp + "p",
            new XAttribute("id",          ctx.NextParaId().ToString()),
            new XAttribute("paraPrIDRef", "0"),
            new XAttribute("styleIDRef",  "0"),
            new XAttribute("pageBreak",   "0"),
            new XAttribute("columnBreak", "0"),
            new XAttribute("merged",      "0"),
            run);
        para.Add(BuildLineseg(10.0, 1.6));
        return para;
    }

    private XElement BuildShape(ShapeObject shape, WriteContext ctx)
    {
        // 종류별 hp: 요소 매핑 (공식 OWPML 클래스명 참고).
        // Spline/ClosedSpline 은 hp:curve(<hp:seg> Bezier 구조) 가 필요한데 PolyDonky
        // 모델엔 control point 가 없어 정확히 재현 불가 — 폴리곤으로 폴백 (직선 연결).
        string elemName = shape.Kind switch
        {
            ShapeKind.Line       => "line",
            ShapeKind.Ellipse    => "ellipse",
            ShapeKind.Polyline
            or ShapeKind.Polygon
            or ShapeKind.Triangle
            or ShapeKind.RegularPolygon
            or ShapeKind.Star
            or ShapeKind.Spline
            or ShapeKind.ClosedSpline => "polygon",
            _                    => "rect", // Rectangle/RoundedRect
        };

        // 한컴 ground truth: 닫힌 도형은 numberingType="PICTURE", 선은 "NONE".
        string numType = elemName == "line" ? "NONE" : "PICTURE";

        long w = UnitConverter.MmToHwpUnit((shape.WidthMm  > 0 ? shape.WidthMm  : 40));
        long h = UnitConverter.MmToHwpUnit((shape.HeightMm > 0 ? shape.HeightMm : 30));
        if (w <= 0) w = 1000;
        if (h <= 0) h = 1;

        // 선·테두리 색·두께. 한컴 ground truth: 두께 0 이어도 style="SOLID" 유지.
        string strokeColor = string.IsNullOrEmpty(shape.StrokeColor) ? "#000000" : shape.StrokeColor;
        long strokeWidth   = (long)Math.Round(shape.StrokeThicknessPt * 100);
        if (strokeWidth < 0) strokeWidth = 0;
        string strokeStyle = "SOLID";

        var elem = new XElement(Hp + elemName,
            new XAttribute("id",            ctx.NextObjId().ToString()),
            new XAttribute("zOrder",        ctx.NextZOrder().ToString()),
            new XAttribute("numberingType", numType),
            // 도형은 anchored overlay 라서 textWrap=IN_FRONT_OF_TEXT 로 본문 위에 그림.
            // BehindText 면 BEHIND_TEXT, 그 외는 IN_FRONT_OF_TEXT.
            // 이전 TOP_AND_BOTTOM 은 본문이 도형 위·아래로 밀려나 페이지 흐름 망가짐.
            new XAttribute("textWrap",
                shape.WrapMode == ImageWrapMode.BehindText ? "BEHIND_TEXT" : "IN_FRONT_OF_TEXT"),
            new XAttribute("textFlow",      "BOTH_SIDES"),
            new XAttribute("lock",          "0"),
            new XAttribute("dropcapstyle",  "None"),
            new XAttribute("href",          ""),
            new XAttribute("groupLevel",    "0"),
            new XAttribute("instid",        ctx.NextInstId().ToString()));

        // 도형별 추가 속성.
        if (elemName == "line")
            elem.Add(new XAttribute("isReverseHV", "0"));
        else if (elemName == "rect")
        {
            // ratio = 모서리 반지름 비율 (0~50, 단위는 셈프상 ‱ 같음).
            // RoundedRect 면 CornerRadiusMm 를 width/height 의 짧은 변 대비 비율로 변환.
            int ratio = 0;
            if (shape.Kind == ShapeKind.RoundedRect && shape.CornerRadiusMm > 0)
            {
                double minSide = Math.Min(shape.WidthMm, shape.HeightMm);
                if (minSide > 0)
                {
                    double frac = shape.CornerRadiusMm / (minSide / 2.0);
                    ratio = (int)Math.Round(Math.Clamp(frac, 0.0, 1.0) * 50);
                }
            }
            elem.Add(new XAttribute("ratio", ratio.ToString()));
        }
        else if (elemName == "ellipse")
        {
            elem.Add(new XAttribute("intervalDirty", "0"));
            elem.Add(new XAttribute("hasArcPr",      "0"));
            elem.Add(new XAttribute("arcType",       "NORMAL"));
        }

        elem.Add(new XElement(Hp + "offset", new XAttribute("x", "0"), new XAttribute("y", "0")));
        elem.Add(new XElement(Hp + "orgSz",  new XAttribute("width", w.ToString()), new XAttribute("height", h.ToString())));
        elem.Add(new XElement(Hp + "curSz",  new XAttribute("width", w.ToString()), new XAttribute("height", h.ToString())));
        elem.Add(new XElement(Hp + "flip",   new XAttribute("horizontal", "0"), new XAttribute("vertical", "0")));
        elem.Add(new XElement(Hp + "rotationInfo",
            new XAttribute("angle",       "0"),
            new XAttribute("centerX",     (w / 2).ToString()),
            new XAttribute("centerY",     (h / 2).ToString()),
            new XAttribute("rotateimage", "1")));
        elem.Add(new XElement(Hp + "renderingInfo",
            new XElement(Hc + "transMatrix",
                new XAttribute("e1", "1"), new XAttribute("e2", "0"), new XAttribute("e3", "0"),
                new XAttribute("e4", "0"), new XAttribute("e5", "1"), new XAttribute("e6", "0")),
            new XElement(Hc + "scaMatrix",
                new XAttribute("e1", "1"), new XAttribute("e2", "0"), new XAttribute("e3", "0"),
                new XAttribute("e4", "0"), new XAttribute("e5", "1"), new XAttribute("e6", "0")),
            new XElement(Hc + "rotMatrix",
                new XAttribute("e1", "1"), new XAttribute("e2", "0"), new XAttribute("e3", "0"),
                new XAttribute("e4", "0"), new XAttribute("e5", "1"), new XAttribute("e6", "0"))));
        elem.Add(new XElement(Hp + "lineShape",
            new XAttribute("color",        strokeColor),
            new XAttribute("width",        strokeWidth.ToString()),
            new XAttribute("style",        strokeStyle),
            new XAttribute("endCap",       "FLAT"),
            new XAttribute("headStyle",    "NORMAL"),
            new XAttribute("tailStyle",    "NORMAL"),
            new XAttribute("headfill",     "1"),
            new XAttribute("tailfill",     "1"),
            new XAttribute("headSz",       "SMALL_SMALL"),
            new XAttribute("tailSz",       "SMALL_SMALL"),
            new XAttribute("outlineStyle", "NORMAL"),
            new XAttribute("alpha",        "0")));

        // hc:fillBrush 는 채울 색이 있을 때만 추가. 한컴 ground truth 도 빈 채우기는 생략.
        // Line 은 시각적으로 채우기 없으므로 fillBrush 자체를 출력하지 않는다.
        bool hasFill = !string.IsNullOrEmpty(shape.FillColor);
        if (hasFill && shape.Kind != ShapeKind.Line && shape.Kind != ShapeKind.Polyline)
        {
            // FillOpacity (0.0=완전 투명 ~ 1.0=완전 불투명) → winBrush.alpha (0=불투명 ~ 255=투명).
            // condition.hwpx 는 불투명 채움을 alpha="0" 로 표현 — 즉 alpha 는 \"투명도\".
            double opacity = Math.Clamp(shape.FillOpacity, 0.0, 1.0);
            int alpha = (int)Math.Round((1.0 - opacity) * 255);
            elem.Add(new XElement(Hc + "fillBrush",
                new XElement(Hc + "winBrush",
                    new XAttribute("faceColor",  shape.FillColor!),
                    new XAttribute("hatchColor", "#000000"),
                    new XAttribute("alpha",      alpha.ToString()))));
        }

        elem.Add(new XElement(Hp + "shadow",
            new XAttribute("type",    "NONE"),
            new XAttribute("color",   "#B2B2B2"),
            new XAttribute("offsetX", "0"),
            new XAttribute("offsetY", "0"),
            new XAttribute("alpha",   "0")));

        // Line 만 startPt/endPt 좌표 추가.
        if (shape.Kind == ShapeKind.Line)
        {
            // Points[0] = 시작, Points[1] = 끝 (mm). 없으면 좌상단→우하단 기본.
            double sx = 0, sy = 0, ex = shape.WidthMm, ey = shape.HeightMm;
            if (shape.Points.Count >= 2)
            {
                sx = shape.Points[0].X; sy = shape.Points[0].Y;
                ex = shape.Points[1].X; ey = shape.Points[1].Y;
            }
            long sxU = UnitConverter.MmToHwpUnit(sx);
            long syU = UnitConverter.MmToHwpUnit(sy);
            long exU = UnitConverter.MmToHwpUnit(ex);
            long eyU = UnitConverter.MmToHwpUnit(ey);
            elem.Add(new XElement(Hc + "startPt",
                new XAttribute("x", sxU.ToString()), new XAttribute("y", syU.ToString())));
            elem.Add(new XElement(Hc + "endPt",
                new XAttribute("x", exU.ToString()), new XAttribute("y", eyU.ToString())));
        }
        else if (elemName == "polygon")
        {
            // hp:polygon: 꼭짓점 hc:pt 시리즈.
            // 닫힌 도형 (Polygon/Triangle/RegularPolygon/Star/ClosedSpline) → 첫 점 반복으로 닫음.
            // 열린 도형 (Polyline/Spline) → 닫힘점 추가하지 않음.
            var verts = ResolvePolygonVertices(shape, w, h);
            foreach (var (vx, vy) in verts)
                elem.Add(new XElement(Hc + "pt",
                    new XAttribute("x", vx.ToString()),
                    new XAttribute("y", vy.ToString())));
            bool isOpen = shape.Kind == ShapeKind.Polyline || shape.Kind == ShapeKind.Spline;
            if (!isOpen && verts.Count > 0
                && (verts[0].X != verts[^1].X || verts[0].Y != verts[^1].Y))
            {
                elem.Add(new XElement(Hc + "pt",
                    new XAttribute("x", verts[0].X.ToString()),
                    new XAttribute("y", verts[0].Y.ToString())));
            }
        }
        else if (elemName == "ellipse")
        {
            // hp:ellipse: center / ax1 / ax2 / start1 / start2 / end1 / end2 (모두 hc:).
            // 전체 타원: start = end = ax1 (시작 = 끝 같은 점이면 닫힌 타원으로 그려짐).
            long cx = w / 2, cy = h / 2;
            // ax1 = 가로축 끝점 (오른쪽 중간). ax2 = 세로축 끝점 (아래 중간).
            elem.Add(new XElement(Hc + "center", new XAttribute("x", cx.ToString()), new XAttribute("y", cy.ToString())));
            elem.Add(new XElement(Hc + "ax1",    new XAttribute("x", w.ToString()),  new XAttribute("y", cy.ToString())));
            elem.Add(new XElement(Hc + "ax2",    new XAttribute("x", cx.ToString()), new XAttribute("y", h.ToString())));
            elem.Add(new XElement(Hc + "start1", new XAttribute("x", w.ToString()),  new XAttribute("y", cy.ToString())));
            elem.Add(new XElement(Hc + "start2", new XAttribute("x", w.ToString()),  new XAttribute("y", cy.ToString())));
            elem.Add(new XElement(Hc + "end1",   new XAttribute("x", w.ToString()),  new XAttribute("y", cy.ToString())));
            elem.Add(new XElement(Hc + "end2",   new XAttribute("x", w.ToString()),  new XAttribute("y", cy.ToString())));
        }
        else
        {
            // Rectangle / RoundedRect: 4 코너 좌표 (HWPUNIT, 좌상단 기준).
            elem.Add(new XElement(Hc + "pt0", new XAttribute("x", "0"),         new XAttribute("y", "0")));
            elem.Add(new XElement(Hc + "pt1", new XAttribute("x", w.ToString()), new XAttribute("y", "0")));
            elem.Add(new XElement(Hc + "pt2", new XAttribute("x", w.ToString()), new XAttribute("y", h.ToString())));
            elem.Add(new XElement(Hc + "pt3", new XAttribute("x", "0"),         new XAttribute("y", h.ToString())));
        }

        elem.Add(new XElement(Hp + "sz",
            new XAttribute("width",       w.ToString()),
            new XAttribute("widthRelTo",  "ABSOLUTE"),
            new XAttribute("height",      h.ToString()),
            new XAttribute("heightRelTo", "ABSOLUTE"),
            new XAttribute("protect",     "0")));

        // 도형 위치 — 모든 도형을 anchored overlay 로 출력.
        // 이유: WrapMode=Inline 기본값으로 만들면 treatAsChar=1 → 도형 높이가
        // 단락 라인 높이가 되어 큰 도형은 다음 페이지로 밀림. 사용자 보고:
        // 직선만 InFrontOfText 라서 정상이고, 사각형/타원/폴리곤 등 Inline 기본값
        // 도형은 모두 다음 페이지로 넘어감. 한컴 ground truth 도 도형은 anchored.
        // OverlayXMm/Y 가 설정되어 있으면 그 값, 아니면 (0,0) 기본 위치.
        long xOff = UnitConverter.MmToHwpUnit(shape.OverlayXMm);
        long yOff = UnitConverter.MmToHwpUnit(shape.OverlayYMm);
        elem.Add(new XElement(Hp + "pos",
            new XAttribute("treatAsChar",     "0"),
            new XAttribute("affectLSpacing",  "0"),
            new XAttribute("flowWithText",    "0"),
            new XAttribute("allowOverlap",    "1"),
            new XAttribute("holdAnchorAndSO", "0"),
            new XAttribute("vertRelTo",       "PAPER"),
            new XAttribute("horzRelTo",       "PAPER"),
            new XAttribute("vertAlign",       "TOP"),
            new XAttribute("horzAlign",       "LEFT"),
            new XAttribute("vertOffset",      yOff.ToString()),
            new XAttribute("horzOffset",      xOff.ToString())));
        elem.Add(new XElement(Hp + "outMargin",
            new XAttribute("left", "0"), new XAttribute("right",  "0"),
            new XAttribute("top",  "0"), new XAttribute("bottom", "0")));

        return elem;
    }

    // Polygon/Curve 의 꼭짓점 좌표 (HWPUNIT, 바운딩 박스 좌상단 0,0 기준).
    // ShapeKind 별 보강:
    //   - Polyline/Polygon/Spline: shape.Points 가 정의되어 있으면 그대로 사용
    //   - Triangle: Points 가 비어있으면 등변 삼각형 (위 꼭짓점→오른쪽 아래→왼쪽 아래)
    //   - RegularPolygon: Points 가 비어있으면 SideCount 개 정다각형 자동 계산
    //   - Star: 외/내부 반지름 교차로 별 모양 자동 계산
    private static List<(long X, long Y)> ResolvePolygonVertices(ShapeObject shape, long w, long h)
    {
        var result = new List<(long, long)>();

        if (shape.Points.Count >= 2)
        {
            foreach (var p in shape.Points)
            {
                long px = UnitConverter.MmToHwpUnit(p.X);
                long py = UnitConverter.MmToHwpUnit(p.Y);
                if (px < 0) px = 0; if (py < 0) py = 0;
                if (px > w) px = w; if (py > h) py = h;
                result.Add((px, py));
            }
            return result;
        }

        switch (shape.Kind)
        {
            case ShapeKind.Triangle:
                result.Add((w / 2, 0));   // top
                result.Add((w, h));       // bottom-right
                result.Add((0, h));       // bottom-left
                break;

            case ShapeKind.RegularPolygon:
            case ShapeKind.Star:
            {
                int n = Math.Max(3, shape.SideCount);
                long cx = w / 2, cy = h / 2;
                long rx = w / 2, ry = h / 2;
                bool isStar = shape.Kind == ShapeKind.Star;
                int total = isStar ? n * 2 : n;
                double inner = Math.Clamp(shape.InnerRadiusRatio, 0.1, 0.95);
                for (int i = 0; i < total; i++)
                {
                    double t = (Math.PI * 2 * i) / total - Math.PI / 2; // start at top
                    double radiusFactor = isStar && (i % 2 == 1) ? inner : 1.0;
                    long vx = cx + (long)Math.Round(Math.Cos(t) * rx * radiusFactor);
                    long vy = cy + (long)Math.Round(Math.Sin(t) * ry * radiusFactor);
                    if (vx < 0) vx = 0; if (vy < 0) vy = 0;
                    if (vx > w) vx = w; if (vy > h) vy = h;
                    result.Add((vx, vy));
                }
                break;
            }

            default:
                // Polyline/Polygon with no Points: rectangle outline as fallback.
                result.Add((0, 0));
                result.Add((w, 0));
                result.Add((w, h));
                result.Add((0, h));
                break;
        }
        return result;
    }
}

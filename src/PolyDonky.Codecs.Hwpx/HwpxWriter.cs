using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using PolyDonky.Core;

namespace PolyDonky.Codecs.Hwpx;

public sealed class HwpxWriter : IDocumentWriter
{
    public string FormatId => "hwpx";

    private const double HwpUnitToMm = 25.4 / 7200.0;

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

        public void ResetParaId() => _nextParaId = 0;
        public int  NextParaId()  => _nextParaId++;
        public long NextObjId()   => _nextObjId++;
        public int  NextZOrder()  => _nextZOrder++;
        public long NextInstId()  => _nextInstId++;

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
        var borderFills = new XElement(Hh + "borderFills",
            new XAttribute("itemCnt", "3"),
            bf1, bf2, bf3);

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

        long indentHwp = (long)Math.Round(s.IndentFirstLineMm / HwpUnitToMm);
        long leftHwp   = (long)Math.Round(s.IndentLeftMm      / HwpUnitToMm);
        long rightHwp  = (long)Math.Round(s.IndentRightMm     / HwpUnitToMm);
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

        if (section.Blocks.Count == 0)
        {
            var para = BuildEmptyParagraph(ctx);
            if (injectSecPr) PrependSecPrRun(para);
            sec.Add(para);
        }
        else
        {
            foreach (var block in section.Blocks)
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
            new XAttribute("pageBreak",   "0"),
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
        int charPrID = ctx.RunStyleId(run.Style);
        var elem = new XElement(Hp + "run",
            new XAttribute("charPrIDRef", charPrID.ToString()));
        if (run.Text.Length > 0)
            elem.Add(new XElement(Hp + "t", run.Text));
        return elem;
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
                colWidths[i] = (long)Math.Round(table.Columns[i].WidthMm / HwpUnitToMm);
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
                ? (long)Math.Round(table.Rows[i].HeightMm / HwpUnitToMm)
                : 1000;

        // 한컴 기본 마진 (HwpForge DEFAULT_CELL_MARGIN/OUT_MARGIN)
        const long CellMarginLeftRight = 510;
        const long CellMarginTopBottom = 141;
        const long OuterMarginAll      = 283;

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
            new XAttribute("borderFillIDRef", "3"),  // SOLID 4면 (id=3)
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
                    cellW = (long)Math.Round(cell.WidthMm / HwpUnitToMm);
                if (cellW <= 0) cellW = ContentWidth / Math.Max(colCount, 1);

                long cellH = 0;
                for (int k = 0; k < rowSpan && r + k < rowCount; k++) cellH += rowHeights[r + k];
                if (cellH <= 0) cellH = 1000;

                var wcell = new XElement(Hp + "tc",
                    new XAttribute("name",            string.Empty),
                    new XAttribute("header",          row.IsHeader ? "1" : "0"),
                    new XAttribute("hasMargin",       "0"),
                    new XAttribute("protect",         "0"),
                    new XAttribute("editable",        "0"),
                    new XAttribute("dirty",           "0"),
                    new XAttribute("borderFillIDRef", "3")); // SOLID 4면 cells

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
        long w = (long)Math.Round((image.WidthMm  > 0 ? image.WidthMm  : 80) / HwpUnitToMm);
        long h = (long)Math.Round((image.HeightMm > 0 ? image.HeightMm : 60) / HwpUnitToMm);

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
            // 공식 OWPML CPictureType.InitMap 순서:
            // lineShape → imgRect → imgClip → effects → inMargin → imgDim → img
            // (Picture 는 AbstractDrawingObjectType 가 아니므로 shadow/fillBrush/drawText 없음)
            new XElement(Hp + "lineShape",
                new XAttribute("color",        "#000000"),
                new XAttribute("width",        "0"),
                new XAttribute("style",        "NONE"),
                new XAttribute("endCap",       "FLAT"),
                new XAttribute("headStyle",    "NORMAL"),
                new XAttribute("tailStyle",    "NORMAL"),
                new XAttribute("headfill",     "1"),
                new XAttribute("tailfill",     "1"),
                new XAttribute("headSz",       "SMALL_SMALL"),
                new XAttribute("tailSz",       "SMALL_SMALL"),
                new XAttribute("outlineStyle", "NORMAL"),
                new XAttribute("alpha",        "0")),
            new XElement(Hp + "imgRect",
                new XElement(Hc + "pt0", new XAttribute("x", "0"),         new XAttribute("y", "0")),
                new XElement(Hc + "pt1", new XAttribute("x", w.ToString()), new XAttribute("y", "0")),
                new XElement(Hc + "pt2", new XAttribute("x", w.ToString()), new XAttribute("y", h.ToString())),
                new XElement(Hc + "pt3", new XAttribute("x", "0"),         new XAttribute("y", h.ToString()))),
            new XElement(Hp + "imgClip",
                new XAttribute("left",   "0"),         new XAttribute("top",    "0"),
                new XAttribute("right",  w.ToString()), new XAttribute("bottom", h.ToString())),
            new XElement(Hp + "effects"),
            new XElement(Hp + "inMargin",
                new XAttribute("left", "0"), new XAttribute("right",  "0"),
                new XAttribute("top",  "0"), new XAttribute("bottom", "0")),
            new XElement(Hp + "imgDim",
                new XAttribute("dimwidth",  w.ToString()),
                new XAttribute("dimheight", h.ToString())),
            new XElement(Hp + "img",
                new XAttribute("binaryItemIDRef", binId),
                new XAttribute("bright",   "0"),
                new XAttribute("contrast", "0"),
                new XAttribute("effect",   "REAL_PIC"),
                new XAttribute("alpha",    "0")),
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
        // 종류별 매핑: Line → hp:line, Rectangle/RoundedRect → hp:rect,
        // Ellipse → hp:ellipse, 그 외 → hp:rect placeholder.
        string elemName = shape.Kind switch
        {
            ShapeKind.Line     => "line",
            ShapeKind.Ellipse  => "ellipse",
            _                  => "rect", // Rectangle 등 폴백
        };

        long w = (long)Math.Round((shape.WidthMm  > 0 ? shape.WidthMm  : 40) / HwpUnitToMm);
        long h = (long)Math.Round((shape.HeightMm > 0 ? shape.HeightMm : 30) / HwpUnitToMm);
        if (w <= 0) w = 1000;
        if (h <= 0) h = 1;

        // 선·테두리 색·두께.
        string strokeColor = string.IsNullOrEmpty(shape.StrokeColor) ? "#000000" : shape.StrokeColor;
        long strokeWidth   = Math.Max(1, (long)Math.Round(shape.StrokeThicknessPt * 100));
        string strokeStyle = shape.StrokeThicknessPt > 0 ? "SOLID" : "NONE";

        var elem = new XElement(Hp + elemName,
            new XAttribute("id",            ctx.NextObjId().ToString()),
            new XAttribute("zOrder",        ctx.NextZOrder().ToString()),
            new XAttribute("numberingType", "NONE"),
            new XAttribute("textWrap",      "TOP_AND_BOTTOM"),
            new XAttribute("textFlow",      "BOTH_SIDES"),
            new XAttribute("lock",          "0"),
            new XAttribute("dropcapstyle",  "None"),
            new XAttribute("href",          ""),
            new XAttribute("groupLevel",    "0"),
            new XAttribute("instid",        ctx.NextInstId().ToString()),
            new XAttribute("isReverseHV",   "0"),

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
            new XElement(Hp + "lineShape",
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

        // 공식 OWPML CAbstractDrawingObjectType 의 fillBrush 는 hp:line/rect/ellipse 등
        // 모든 drawing object 의 필수 자식. 채움색 없으면 winBrush alpha=255 (투명).
        bool hasFill = !string.IsNullOrEmpty(shape.FillColor);
        elem.Add(new XElement(Hc + "fillBrush",
            new XElement(Hc + "winBrush",
                new XAttribute("faceColor",  hasFill ? shape.FillColor! : "#FFFFFF"),
                new XAttribute("hatchColor", "#000000"),
                new XAttribute("alpha",      hasFill ? "0" : "255"))));

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
            long sxU = (long)Math.Round(sx / HwpUnitToMm);
            long syU = (long)Math.Round(sy / HwpUnitToMm);
            long exU = (long)Math.Round(ex / HwpUnitToMm);
            long eyU = (long)Math.Round(ey / HwpUnitToMm);
            elem.Add(new XElement(Hc + "startPt",
                new XAttribute("x", sxU.ToString()), new XAttribute("y", syU.ToString())));
            elem.Add(new XElement(Hc + "endPt",
                new XAttribute("x", exU.ToString()), new XAttribute("y", eyU.ToString())));
        }
        else
        {
            // Rectangle / RoundedRect / Ellipse 등: 4 코너 좌표 (HWPUNIT, 좌상단 기준).
            // 한컴은 hp:rect/ellipse 등 닫힌 도형에 hc:pt0~pt3 를 요구 — 누락 시 거부.
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

        // Wrap 모드별 위치: InFrontOfText/BehindText = 오버레이 (절대 위치),
        // 그 외 (Inline/WrapLeft/WrapRight) = 글자처럼 인라인 배치.
        bool isOverlay = shape.WrapMode == ImageWrapMode.InFrontOfText
                      || shape.WrapMode == ImageWrapMode.BehindText;
        if (isOverlay)
        {
            long xOff = (long)Math.Round(shape.OverlayXMm / HwpUnitToMm);
            long yOff = (long)Math.Round(shape.OverlayYMm / HwpUnitToMm);
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
        }
        else
        {
            elem.Add(new XElement(Hp + "pos",
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
                new XAttribute("horzOffset",      "0")));
        }
        elem.Add(new XElement(Hp + "outMargin",
            new XAttribute("left", "0"), new XAttribute("right",  "0"),
            new XAttribute("top",  "0"), new XAttribute("bottom", "0")));

        return elem;
    }
}

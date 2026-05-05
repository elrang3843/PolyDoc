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
        private readonly Dictionary<string, string> _imageIdByHash = new(StringComparer.Ordinal);

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
        public int NextParaId() => _nextParaId++;

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
            using var stream = entry.Open();
            stream.Write(data, 0, data.Length);
            _imageIdByHash[hashKey] = id;
            return id;
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
        WriteManifestXml(archive);
        WriteSettingsXml(archive);
        WriteContentHpf(archive, document);
        WriteVersionXml(archive);

        var ctx = new WriteContext(archive);

        // Walk pass: register every used RunStyle / ParagraphStyle / FontFamily.
        foreach (var section in document.Sections)
            foreach (var block in section.Blocks)
                ctx.RegisterFromBlock(block);

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
        var doc = new XDocument(new XDeclaration("1.0", "utf-8", null),
            new XElement(OpfContainer + "container",
                new XAttribute("version", "1.0"),
                new XElement(OpfContainer + "rootfiles",
                    new XElement(OpfContainer + "rootfile",
                        new XAttribute("full-path", HwpxPaths.ContentHpf),
                        new XAttribute("media-type", "application/hwpml-package+xml")))));
        WriteXml(archive, HwpxPaths.ContainerXml, doc);
    }

    private void WriteManifestXml(ZipArchive archive)
    {
        // ODF-style empty manifest — required by Hangul's package validator.
        var ns = XNamespace.Get("urn:oasis:names:tc:opendocument:xmlns:manifest:1.0");
        var doc = new XDocument(new XDeclaration("1.0", "utf-8", null),
            new XElement(ns + "manifest"));
        WriteXml(archive, HwpxPaths.ManifestXml, doc);
    }

    private void WriteSettingsXml(ZipArchive archive)
    {
        var doc = new XDocument(new XDeclaration("1.0", "utf-8", null),
            new XElement(Ha + "HWPApplicationSettings",
                new XAttribute(XNamespace.Xmlns + HwpxNamespaces.PrefixHa, Ha.NamespaceName)));
        WriteXml(archive, HwpxPaths.SettingsXml, doc);
    }

    private void WriteContentHpf(ZipArchive archive, PolyDonkyument document)
    {
        // hrefs are relative to the location of content.hpf (Contents/), so section0.xml not Contents/section0.xml.
        // settings.xml is at root, so go up one level with "../".
        var manifest = new XElement(Opf + "manifest",
            new XElement(Opf + "item",
                new XAttribute("id", "header"),
                new XAttribute("href", "header.xml"),
                new XAttribute("media-type", "application/xml")),
            new XElement(Opf + "item",
                new XAttribute("id", "settings"),
                new XAttribute("href", "../settings.xml"),
                new XAttribute("media-type", "application/xml")));

        // spine: only content sections (not header — header is structural, not a body section)
        var spine = new XElement(Opf + "spine");

        int sectionCount = Math.Max(document.Sections.Count, 1);
        for (int i = 0; i < sectionCount; i++)
        {
            var id = $"section{i}";
            manifest.Add(new XElement(Opf + "item",
                new XAttribute("id", id),
                new XAttribute("href", $"section{i}.xml"),
                new XAttribute("media-type", "application/xml")));
            spine.Add(new XElement(Opf + "itemref", new XAttribute("idref", id)));
        }

        var metadata = new XElement(Opf + "metadata");
        if (!string.IsNullOrEmpty(document.Metadata.Title))
            metadata.Add(new XElement(Dc + "title", document.Metadata.Title));
        if (!string.IsNullOrEmpty(document.Metadata.Author))
            metadata.Add(new XElement(Dc + "creator", document.Metadata.Author));
        metadata.Add(new XElement(Dc + "language", document.Metadata.Language ?? "ko"));
        metadata.Add(new XElement(Opf + "meta",
            new XAttribute("name", "producer"),
            new XAttribute("content", $"{HwpxFormat.ProducedBy}/{HwpxFormat.Version}")));

        var package = new XElement(Opf + "package",
            new XAttribute(XNamespace.Xmlns + "opf", Opf.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "dc",  Dc.NamespaceName),
            new XAttribute("version", HwpxFormat.Version),
            new XAttribute("unique-identifier", "polydoc-id"),
            metadata, manifest, spine);

        WriteXml(archive, HwpxPaths.ContentHpf, new XDocument(new XDeclaration("1.0", "utf-8", null), package));
    }

    private void WriteVersionXml(ZipArchive archive)
    {
        // KS X 6101: "tagetApplication" is the canonical (intentionally misspelled) attribute name.
        var doc = new XDocument(new XDeclaration("1.0", "utf-8", null),
            new XElement(Hv + "HCFVersion",
                new XAttribute(XNamespace.Xmlns + HwpxNamespaces.PrefixHv, Hv.NamespaceName),
                new XAttribute("tagetApplication", "WORDPROCESSOR"),
                new XAttribute("major", "5"),
                new XAttribute("minor", "1"),
                new XAttribute("micro", "1"),
                new XAttribute("buildNumber", "0"),
                new XAttribute("os", "10"),
                new XAttribute("xmlVersion", "1.5"),
                new XAttribute("application", "Hancom Office Hangul"),
                new XAttribute("appVersion", "12.0.0.0")));
        WriteXml(archive, HwpxPaths.VersionXml, doc);
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

    private void WriteHeaderXml(ZipArchive archive, WriteContext ctx, int sectionCount)
    {
        // fontfaces — 7 language groups, each carrying the full font registry.
        var fontFaces = new XElement(Hh + "fontfaces",
            new XAttribute("itemCnt", FontLangs.Length.ToString()));
        foreach (var lang in FontLangs)
        {
            var ff = new XElement(Hh + "fontface",
                new XAttribute("lang", lang),
                new XAttribute("count", ctx.Fonts.Count.ToString()));
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

        // borderFills — id=0 default (no border, no fill).
        var borderFills = new XElement(Hh + "borderFills",
            new XAttribute("itemCnt", "1"),
            new XElement(Hh + "borderFill",
                new XAttribute("id", "0"),
                new XAttribute("threeD", "0"),
                new XAttribute("shadow", "0"),
                new XAttribute("centerLine", "0"),
                new XAttribute("breakCellSeparateLine", "0"),
                new XElement(Hh + "slash",
                    new XAttribute("type", "NONE"),
                    new XAttribute("Crooked", "0"),
                    new XAttribute("isCounter", "0")),
                new XElement(Hh + "backSlash",
                    new XAttribute("type", "NONE"),
                    new XAttribute("Crooked", "0"),
                    new XAttribute("isCounter", "0")),
                new XElement(Hh + "leftBorder",   new XAttribute("type", "NONE"), new XAttribute("width", "0.1 mm"), new XAttribute("color", "#000000")),
                new XElement(Hh + "rightBorder",  new XAttribute("type", "NONE"), new XAttribute("width", "0.1 mm"), new XAttribute("color", "#000000")),
                new XElement(Hh + "topBorder",    new XAttribute("type", "NONE"), new XAttribute("width", "0.1 mm"), new XAttribute("color", "#000000")),
                new XElement(Hh + "bottomBorder", new XAttribute("type", "NONE"), new XAttribute("width", "0.1 mm"), new XAttribute("color", "#000000")),
                new XElement(Hh + "diagonal",     new XAttribute("type", "NONE"), new XAttribute("width", "0.1 mm"), new XAttribute("color", "#000000")),
                new XElement(Hh + "fillInfo", new XElement(Hh + "noFill"))));

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

        // numberings and bullets — empty lists (no ordered/unordered list styles)
        var numberings = new XElement(Hh + "numberings", new XAttribute("itemCnt", "0"));
        var bullets     = new XElement(Hh + "bullets",    new XAttribute("itemCnt", "0"));

        // styles: 0=바탕글(Normal), 1~6=개요 1~6(Heading1~6)
        var styles = new XElement(Hh + "styles", new XAttribute("itemCnt", "7"));
        styles.Add(new XElement(Hh + "style",
            new XAttribute("id", "0"), new XAttribute("type", "PARA"),
            new XAttribute("name", "바탕글"), new XAttribute("engName", "Normal"),
            new XAttribute("paraPrIDRef", "0"), new XAttribute("charPrIDRef", "0"),
            new XAttribute("nextStyleIDRef", "0"), new XAttribute("langIDRef", "1042"),
            new XAttribute("lockForm", "0")));
        for (int level = 1; level <= 6; level++)
        {
            styles.Add(new XElement(Hh + "style",
                new XAttribute("id", level.ToString()), new XAttribute("type", "PARA"),
                new XAttribute("name", $"개요 {level}"), new XAttribute("engName", $"Heading{level}"),
                new XAttribute("paraPrIDRef", "0"), new XAttribute("charPrIDRef", "0"),
                new XAttribute("nextStyleIDRef", "0"), new XAttribute("langIDRef", "1042"),
                new XAttribute("lockForm", "0")));
        }

        // KS X 6101 §5 규정 순서: fontfaces → borderFills → charProperties → tabProperties
        //   → paraProperties → numberings → bullets → styles
        var refList = new XElement(Hh + "refList",
            fontFaces, borderFills, charProps, tabProperties, paraProps,
            numberings, bullets, styles);

        // docInfo — document metadata
        var docInfo = new XElement(Hh + "docInfo",
            new XElement(Hh + "summary",
                new XElement(Hh + "title"),
                new XElement(Hh + "subject"),
                new XElement(Hh + "author"),
                new XElement(Hh + "date", DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss")),
                new XElement(Hh + "keyword"),
                new XElement(Hh + "comment")));

        // compatibleDocument — required for Hangul 2018+ compatibility flag
        var compatDoc = new XElement(Hh + "compatibleDocument",
            new XAttribute("targetProgram", "HWP2018"));

        // page layout is now specified inline via hp:secPr in section0.xml — no masterPageList needed
        var head = new XElement(Hh + "head",
            new XAttribute(XNamespace.Xmlns + HwpxNamespaces.PrefixHh, Hh.NamespaceName),
            new XAttribute("version", HwpxFormat.Version),
            new XAttribute("secCnt", sectionCount.ToString()),
            docInfo,
            new XElement(Hh + "beginNum",
                new XAttribute("page", "1"), new XAttribute("footnote", "1"),
                new XAttribute("endnote", "1"), new XAttribute("pic", "1"),
                new XAttribute("tbl", "1"), new XAttribute("equation", "1")),
            refList,
            compatDoc);

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
            new XAttribute("borderFillIDRef", "0"));

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

        // Required child elements for KS X 6101 charPr completeness
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
            Alignment.Center                        => "CENTER",
            Alignment.Right                         => "RIGHT",
            Alignment.Justify or Alignment.Distributed => "JUSTIFY",
            _                                       => "LEFT",
        };

        long lineSpacingPct = (long)Math.Round(s.LineHeightFactor * 100);
        if (lineSpacingPct <= 0) lineSpacingPct = 160;

        return new XElement(Hh + "paraPr",
            new XAttribute("id", id.ToString()),
            new XAttribute("tabPrIDRef", "0"),
            new XAttribute("condense", "0"),
            new XAttribute("fontLineHeight", "0"),
            new XAttribute("snapToGrid", "0"),
            new XAttribute("suppressLineNumbers", "0"),
            new XAttribute("checked", "0"),
            new XElement(Hh + "align",
                new XAttribute("horizontal", alignStr),
                new XAttribute("vertical", "BASELINE")),
            new XElement(Hh + "heading",
                new XAttribute("type", "NONE"),
                new XAttribute("idRef", "0"),
                new XAttribute("level", "0")),
            new XElement(Hh + "breakSetting",
                new XAttribute("breakLatinWord", "KEEP_WORD"),
                new XAttribute("breakNonLatinWord", "KEEP_WORD"),
                new XAttribute("widowOrphan", "0"),
                new XAttribute("keepWithNext", "0"),
                new XAttribute("keepLines", "0"),
                new XAttribute("pageBreakBefore", "0"),
                new XAttribute("columnBreakBefore", "0")),
            new XElement(Hh + "autoSpacing",
                new XAttribute("eAsianEng", "0"),
                new XAttribute("eAsianNum", "0")),
            new XElement(Hh + "margin",
                new XAttribute("indent", ((long)Math.Round(s.IndentFirstLineMm / HwpUnitToMm)).ToString()),
                new XAttribute("left",   ((long)Math.Round(s.IndentLeftMm      / HwpUnitToMm)).ToString()),
                new XAttribute("right",  ((long)Math.Round(s.IndentRightMm     / HwpUnitToMm)).ToString()),
                new XAttribute("prev",   ((long)Math.Round(s.SpaceBeforePt * 100)).ToString()),
                new XAttribute("next",   ((long)Math.Round(s.SpaceAfterPt  * 100)).ToString())),
            new XElement(Hh + "lineSpacing",
                new XAttribute("type", "PERCENT"),
                new XAttribute("value", lineSpacingPct.ToString())),
            new XElement(Hh + "border",
                new XAttribute("borderFillIDRef", "0"),
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

        var sec = new XElement(Hs + "sec",
            new XAttribute(XNamespace.Xmlns + HwpxNamespaces.PrefixHs, Hs.NamespaceName),
            new XAttribute(XNamespace.Xmlns + HwpxNamespaces.PrefixHp, Hp.NamespaceName),
            new XAttribute(XNamespace.Xmlns + HwpxNamespaces.PrefixHc, Hc.NamespaceName));

        // secPr must be the first run of the first paragraph of the first section (KS X 6101).
        bool injectSecPr = sectionIndex == 0;

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
            Paragraph p    => BuildParagraph(p, ctx),
            Table t        => BuildTableHostingParagraph(t, ctx),
            ImageBlock img => BuildImageHostingParagraph(img, ctx),
            OpaqueBlock op => BuildOpaqueHostingParagraph(op, ctx),
            _              => BuildEmptyParagraph(ctx),
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

    // Injects an hp:secPr control run as the very first child of the paragraph.
    private void PrependSecPrRun(XElement para)
    {
        var secPrRun = new XElement(Hp + "run",
            new XAttribute("charPrIDRef", "0"),
            new XElement(Hp + "secPr",
                new XAttribute("id", ""),
                new XAttribute("textDirection", "HORIZONTAL"),
                new XAttribute("spaceColumns", "1134"),
                new XAttribute("tabStop", "8000"),
                new XAttribute("tabStopVal", "4000"),
                new XAttribute("tabStopUnit", "HWPUNIT"),
                new XAttribute("outlineShapeIDRef", "0"),
                new XAttribute("memoShapeIDRef", "0"),
                new XAttribute("textVerticalWidthHead", "0"),
                new XAttribute("masterPageCnt", "0"),
                new XElement(Hp + "pagePr",
                    new XAttribute("landscape", "PORTRAIT"),
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
                        new XAttribute("bottom", MarginBottom.ToString())))));
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
        long textH  = (long)Math.Round(maxPt * 100);
        long height = (long)Math.Round(textH * lsFactor);
        long descent = (long)Math.Round(textH * 0.25);

        para.Add(new XElement(Hp + "linesegarray",
            new XElement(Hp + "lineseg",
                new XAttribute("textpos",    "0"),
                new XAttribute("vertical",   "0"),
                new XAttribute("height",     height.ToString()),
                new XAttribute("textHeight", textH.ToString()),
                new XAttribute("descent",    descent.ToString()),
                new XAttribute("lineSpacing","0"),
                new XAttribute("horzFmt",    "0"),
                new XAttribute("vertFmt",    "0"),
                new XAttribute("lineBreak",  "0"))));

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
        => new(Hp + "p",
            new XAttribute("id",          ctx.NextParaId().ToString()),
            new XAttribute("paraPrIDRef", "0"),
            new XAttribute("styleIDRef",  "0"),
            new XAttribute("pageBreak",   "0"),
            new XAttribute("columnBreak", "0"),
            new XAttribute("merged",      "0"),
            new XElement(Hp + "run", new XAttribute("charPrIDRef", "0")),
            new XElement(Hp + "linesegarray",
                new XElement(Hp + "lineseg",
                    new XAttribute("textpos", "0"), new XAttribute("vertical", "0"),
                    new XAttribute("height", "1600"), new XAttribute("textHeight", "1000"),
                    new XAttribute("descent", "250"), new XAttribute("lineSpacing", "0"),
                    new XAttribute("horzFmt", "0"), new XAttribute("vertFmt", "0"),
                    new XAttribute("lineBreak", "0"))));

    // ── opaque island ────────────────────────────────────────────────────────

    private XElement BuildOpaqueHostingParagraph(OpaqueBlock op, WriteContext ctx)
    {
        // hwpx 포맷 + XML 있으면 원본 재출력; 그 외 포맷은 placeholder 단락.
        if (op.Format == "hwpx" && !string.IsNullOrEmpty(op.Xml))
        {
            try
            {
                var shapeElem = XElement.Parse(op.Xml);
                var run = new XElement(Hp + "run", new XAttribute("charPrIDRef", "0"), shapeElem);
                return new XElement(Hp + "p",
                    new XAttribute("id",          ctx.NextParaId().ToString()),
                    new XAttribute("paraPrIDRef", "0"),
                    new XAttribute("styleIDRef",  "0"),
                    new XAttribute("pageBreak",   "0"),
                    new XAttribute("columnBreak", "0"),
                    new XAttribute("merged",      "0"),
                    run);
            }
            catch (System.Xml.XmlException)
            {
                // 파싱 실패 시 placeholder 로 fallback.
            }
        }
        return BuildParagraph(Paragraph.Of(op.DisplayLabel), ctx);
    }

    // ── table ────────────────────────────────────────────────────────────────

    private XElement BuildTableHostingParagraph(Table table, WriteContext ctx)
    {
        var run = new XElement(Hp + "run", new XAttribute("charPrIDRef", "0"));
        run.Add(BuildTable(table, ctx));
        return new XElement(Hp + "p",
            new XAttribute("id",          ctx.NextParaId().ToString()),
            new XAttribute("paraPrIDRef", "0"),
            new XAttribute("styleIDRef",  "0"),
            new XAttribute("pageBreak",   "0"),
            new XAttribute("columnBreak", "0"),
            new XAttribute("merged",      "0"),
            run);
    }

    private XElement BuildTable(Table table, WriteContext ctx)
    {
        int rowCount = table.Rows.Count;
        int colCount = table.Columns.Count;
        if (colCount == 0 && rowCount > 0)
            colCount = table.Rows.Max(r => r.Cells.Sum(c => Math.Max(c.ColumnSpan, 1)));

        var wtbl = new XElement(Hp + "tbl",
            new XAttribute("rowCnt", rowCount.ToString()),
            new XAttribute("colCnt", colCount.ToString()),
            new XAttribute("cellSpacing", "0"),
            new XAttribute("borderFillIDRef", "0"));

        wtbl.Add(new XElement(Hp + "sz",
            new XAttribute("width", "0"), new XAttribute("widthRelTo", "ABSOLUTE"),
            new XAttribute("height", "0"), new XAttribute("heightRelTo", "ABSOLUTE"),
            new XAttribute("protect", "0")));
        wtbl.Add(new XElement(Hp + "outMargin",
            new XAttribute("left", "0"), new XAttribute("right", "0"),
            new XAttribute("top", "0"), new XAttribute("bottom", "0")));
        wtbl.Add(new XElement(Hp + "inMargin",
            new XAttribute("left", "0"), new XAttribute("right", "0"),
            new XAttribute("top", "0"), new XAttribute("bottom", "0")));

        for (int r = 0; r < rowCount; r++)
        {
            var row = table.Rows[r];
            var wrow = new XElement(Hp + "tr");
            int c = 0;
            foreach (var cell in row.Cells)
            {
                var wcell = new XElement(Hp + "tc",
                    new XAttribute("name", string.Empty),
                    new XAttribute("header", "0"),
                    new XAttribute("hasMargin", "0"),
                    new XAttribute("protect", "0"),
                    new XAttribute("editable", "0"),
                    new XAttribute("dirty", "0"),
                    new XAttribute("borderFillIDRef", "0"));

                var subList = new XElement(Hp + "subList");
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
                    new XAttribute("colSpan", Math.Max(cell.ColumnSpan, 1).ToString()),
                    new XAttribute("rowSpan", Math.Max(cell.RowSpan, 1).ToString())));
                wcell.Add(new XElement(Hp + "cellSz",
                    new XAttribute("width", ((long)Math.Round(cell.WidthMm / HwpUnitToMm)).ToString()),
                    new XAttribute("height", "0")));
                wcell.Add(new XElement(Hp + "cellMargin",
                    new XAttribute("left", "0"), new XAttribute("right", "0"),
                    new XAttribute("top", "0"), new XAttribute("bottom", "0")));

                wrow.Add(wcell);
                c += Math.Max(cell.ColumnSpan, 1);
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
        return new XElement(Hp + "p",
            new XAttribute("id",          ctx.NextParaId().ToString()),
            new XAttribute("paraPrIDRef", "0"),
            new XAttribute("styleIDRef",  "0"),
            new XAttribute("pageBreak",   "0"),
            new XAttribute("columnBreak", "0"),
            new XAttribute("merged",      "0"),
            run);
    }

    private XElement BuildPicture(ImageBlock image, WriteContext ctx)
    {
        var binId = ctx.AddImage(image.Data, image.MediaType);
        long w = (long)Math.Round((image.WidthMm  > 0 ? image.WidthMm  : 80) / HwpUnitToMm);
        long h = (long)Math.Round((image.HeightMm > 0 ? image.HeightMm : 60) / HwpUnitToMm);

        return new XElement(Hp + "pic",
            new XElement(Hp + "offset",       new XAttribute("x", "0"), new XAttribute("y", "0")),
            new XElement(Hp + "orgSz",         new XAttribute("width", w.ToString()), new XAttribute("height", h.ToString())),
            new XElement(Hp + "curSz",         new XAttribute("width", w.ToString()), new XAttribute("height", h.ToString())),
            new XElement(Hp + "flip",          new XAttribute("horizontal", "0"), new XAttribute("vertical", "0")),
            new XElement(Hp + "rotationInfo",  new XAttribute("angle", "0"), new XAttribute("centerX", "0"), new XAttribute("centerY", "0")),
            new XElement(Hp + "imgClip",
                new XAttribute("left", "0"), new XAttribute("top", "0"),
                new XAttribute("right", "0"), new XAttribute("bottom", "0")),
            new XElement(Hp + "inMargin",
                new XAttribute("left", "0"), new XAttribute("right", "0"),
                new XAttribute("top", "0"), new XAttribute("bottom", "0")),
            new XElement(Hp + "img",
                new XAttribute("binaryItemIDRef", binId),
                new XAttribute("bright", "0"),
                new XAttribute("contrast", "0"),
                new XAttribute("effect", "REAL_PIC"),
                new XAttribute("alpha", "0")));
    }
}

using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using PolyDoc.Core;

namespace PolyDoc.Codecs.Hwpx;

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

    public void Write(PolyDocument document, Stream output)
    {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(output);

        using var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true);
        WriteRawText(archive, HwpxPaths.Mimetype, HwpxPaths.MimetypeContent, CompressionLevel.NoCompression);
        WriteContainerXml(archive);
        WriteContentHpf(archive, document);
        WriteVersionXml(archive);

        var ctx = new WriteContext(archive);

        // Walk pass: register every used RunStyle / ParagraphStyle / FontFamily.
        foreach (var section in document.Sections)
            foreach (var block in section.Blocks)
                ctx.RegisterFromBlock(block);

        WriteHeaderXml(archive, ctx);

        int cnt = document.Sections.Count;
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

    private void WriteContentHpf(ZipArchive archive, PolyDocument document)
    {
        var manifest = new XElement(Opf + "manifest",
            new XElement(Opf + "item",
                new XAttribute("id", "header"),
                new XAttribute("href", "header.xml"),
                new XAttribute("media-type", "application/xml")));

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

        var metadata = new XElement(Opf + "metadata",
            new XAttribute(XNamespace.Xmlns + "dc", Dc.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "opf", Opf.NamespaceName));
        if (!string.IsNullOrEmpty(document.Metadata.Title))
            metadata.Add(new XElement(Dc + "title", document.Metadata.Title));
        if (!string.IsNullOrEmpty(document.Metadata.Author))
            metadata.Add(new XElement(Dc + "creator", document.Metadata.Author));
        metadata.Add(new XElement(Dc + "language", document.Metadata.Language ?? "ko"));
        metadata.Add(new XElement(Opf + "meta",
            new XAttribute("name", "producer"),
            new XAttribute("content", $"{HwpxFormat.ProducedBy}/{HwpxFormat.Version}")));

        var package = new XElement(Opf + "package",
            new XAttribute("version", HwpxFormat.Version),
            new XAttribute("unique-identifier", "polydoc-id"),
            metadata, manifest, spine);

        WriteXml(archive, HwpxPaths.ContentHpf, new XDocument(new XDeclaration("1.0", "utf-8", null), package));
    }

    private void WriteVersionXml(ZipArchive archive)
    {
        var doc = new XDocument(new XDeclaration("1.0", "utf-8", null),
            new XElement(XName.Get("HCFVersion", HwpxNamespaces.Hancom + "version"),
                new XAttribute("targetApplication", "WORDPROCESSOR"),
                new XAttribute("major", "5"),
                new XAttribute("minor", "1"),
                new XAttribute("micro", "1"),
                new XAttribute("buildNumber", "0"),
                new XAttribute("os", "0"),
                new XAttribute("xmlVersion", HwpxFormat.Version)));
        WriteXml(archive, HwpxPaths.VersionXml, doc);
    }

    // ── header.xml (dynamic registry) ────────────────────────────────────────

    private void WriteHeaderXml(ZipArchive archive, WriteContext ctx)
    {
        // fontfaces
        var fonts = new XElement(Hh + "fontfaces",
            new XAttribute("itemCnt", ctx.Fonts.Count.ToString()));
        for (int i = 0; i < ctx.Fonts.Count; i++)
        {
            fonts.Add(new XElement(Hh + "fontface",
                new XAttribute("lang", "HANGUL"),
                new XAttribute("count", "1"),
                new XElement(Hh + "font",
                    new XAttribute("id", i.ToString()),
                    new XAttribute("face", ctx.Fonts[i]),
                    new XAttribute("type", "ttf"),
                    new XAttribute("isEmbedded", "0"))));
        }

        // charProperties — one entry per unique RunStyle
        var charProps = new XElement(Hh + "charProperties",
            new XAttribute("itemCnt", ctx.RunStyles.Count.ToString()));
        for (int id = 0; id < ctx.RunStyles.Count; id++)
            charProps.Add(BuildDynamicCharPr(id, ctx.RunStyles[id], ctx));

        // paraProperties — one entry per unique ParagraphStyle
        var paraProps = new XElement(Hh + "paraProperties",
            new XAttribute("itemCnt", ctx.ParaStyles.Count.ToString()));
        for (int id = 0; id < ctx.ParaStyles.Count; id++)
            paraProps.Add(BuildDynamicParaPr(id, ctx.ParaStyles[id]));

        // styles: 0=Normal, 1~6=Heading1~6 (outline recovery via engName)
        var styles = new XElement(Hh + "styles", new XAttribute("itemCnt", "7"));
        styles.Add(new XElement(Hh + "style",
            new XAttribute("id", "0"), new XAttribute("type", "PARA"),
            new XAttribute("name", "Normal"), new XAttribute("engName", "Normal"),
            new XAttribute("paraPrIDRef", "0"), new XAttribute("charPrIDRef", "0"),
            new XAttribute("nextStyleIDRef", "0"), new XAttribute("langID", "1042"),
            new XAttribute("lockForm", "0")));
        for (int level = 1; level <= 6; level++)
        {
            styles.Add(new XElement(Hh + "style",
                new XAttribute("id", level.ToString()), new XAttribute("type", "PARA"),
                new XAttribute("name", $"Heading{level}"), new XAttribute("engName", $"Heading{level}"),
                new XAttribute("paraPrIDRef", "0"), new XAttribute("charPrIDRef", "0"),
                new XAttribute("nextStyleIDRef", "0"), new XAttribute("langID", "1042"),
                new XAttribute("lockForm", "0")));
        }

        var refList = new XElement(Hh + "refList", fonts, charProps, paraProps, styles);
        var head = new XElement(Hh + "head",
            new XAttribute(XNamespace.Xmlns + HwpxNamespaces.PrefixHh, Hh.NamespaceName),
            new XAttribute(XNamespace.Xmlns + HwpxNamespaces.PrefixHc, Hc.NamespaceName),
            new XAttribute("version", HwpxFormat.Version),
            new XAttribute("secCnt", "1"),
            new XElement(Hh + "beginNum",
                new XAttribute("page", "1"), new XAttribute("footnote", "1"),
                new XAttribute("endnote", "1"), new XAttribute("pic", "1"),
                new XAttribute("tbl", "1"), new XAttribute("equation", "1")),
            refList);

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
            new XAttribute("shadeColor", s.Background is { } bg ? bg.ToHex() : "none"),
            new XAttribute("useFontSpace", "0"),
            new XAttribute("useKerning", "0"),
            new XAttribute("symMark", "NONE"),
            new XAttribute("borderFillIDRef", "0"));

        if (s.Bold)         el.Add(new XElement(Hh + "bold"));
        if (s.Italic)       el.Add(new XElement(Hh + "italic"));
        if (s.Underline)    el.Add(new XElement(Hh + "underline",
                                new XAttribute("type", "BOTTOM"),
                                new XAttribute("shape", "SOLID"),
                                new XAttribute("color", "#000000")));
        if (s.Strikethrough) el.Add(new XElement(Hh + "strikeout",
                                new XAttribute("shape", "SOLID"),
                                new XAttribute("color", "#000000")));
        if (s.Superscript)  el.Add(new XElement(Hh + "supScript"));
        if (s.Subscript)    el.Add(new XElement(Hh + "subScript"));

        var fontId = ctx.FontId(s.FontFamily).ToString();
        el.Add(new XElement(Hh + "fontRef",
            new XAttribute("hangul",   fontId),
            new XAttribute("latin",    fontId),
            new XAttribute("hanja",    fontId),
            new XAttribute("japanese", fontId),
            new XAttribute("other",    fontId),
            new XAttribute("symbol",   fontId),
            new XAttribute("user",     fontId)));

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
            new XElement(Hh + "margin",
                new XAttribute("left",   ((long)Math.Round(s.IndentLeftMm      / HwpUnitToMm)).ToString()),
                new XAttribute("right",  ((long)Math.Round(s.IndentRightMm     / HwpUnitToMm)).ToString()),
                new XAttribute("intent", ((long)Math.Round(s.IndentFirstLineMm / HwpUnitToMm)).ToString()),
                new XAttribute("prev",   ((long)Math.Round(s.SpaceBeforePt * 100)).ToString()),
                new XAttribute("next",   ((long)Math.Round(s.SpaceAfterPt  * 100)).ToString())),
            new XElement(Hh + "lineSpacing",
                new XAttribute("type", "PERCENT"),
                new XAttribute("value", lineSpacingPct.ToString())));
    }

    // ── section XML ──────────────────────────────────────────────────────────

    private void WriteSectionXml(ZipArchive archive, int sectionIndex, Section section, WriteContext ctx)
    {
        var sec = new XElement(Hs + "sec",
            new XAttribute(XNamespace.Xmlns + HwpxNamespaces.PrefixHs, Hs.NamespaceName),
            new XAttribute(XNamespace.Xmlns + HwpxNamespaces.PrefixHp, Hp.NamespaceName),
            new XAttribute(XNamespace.Xmlns + HwpxNamespaces.PrefixHc, Hc.NamespaceName));

        if (section.Blocks.Count == 0)
        {
            sec.Add(BuildEmptyParagraph());
        }
        else
        {
            foreach (var block in section.Blocks)
                AppendBlock(sec, block, ctx);
        }

        WriteXml(archive, HwpxPaths.SectionXml(sectionIndex),
            new XDocument(new XDeclaration("1.0", "utf-8", null), sec));
    }

    private void AppendBlock(XElement target, Block block, WriteContext ctx)
    {
        switch (block)
        {
            case Paragraph p:
                target.Add(BuildParagraph(p, ctx));
                break;
            case Table t:
                target.Add(BuildTableHostingParagraph(t, ctx));
                break;
            case ImageBlock img:
                target.Add(BuildImageHostingParagraph(img, ctx));
                break;
            case OpaqueBlock op:
                target.Add(BuildParagraph(Paragraph.Of(op.DisplayLabel), ctx));
                break;
        }
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
            new XAttribute("paraPrIDRef", paraPrID.ToString()),
            new XAttribute("styleIDRef",  styleID.ToString()));

        foreach (var run in p.Runs)
            para.Add(BuildRun(run, ctx));
        if (p.Runs.Count == 0)
            para.Add(new XElement(Hp + "run", new XAttribute("charPrIDRef", "0")));

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

    private XElement BuildEmptyParagraph()
        => new(Hp + "p",
            new XAttribute("paraPrIDRef", "0"),
            new XAttribute("styleIDRef", "0"),
            new XElement(Hp + "run", new XAttribute("charPrIDRef", "0")));

    // ── table ────────────────────────────────────────────────────────────────

    private XElement BuildTableHostingParagraph(Table table, WriteContext ctx)
    {
        var run = new XElement(Hp + "run", new XAttribute("charPrIDRef", "0"));
        run.Add(BuildTable(table, ctx));
        return new XElement(Hp + "p",
            new XAttribute("paraPrIDRef", "0"),
            new XAttribute("styleIDRef", "0"),
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
                    subList.Add(BuildEmptyParagraph());
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
            new XAttribute("paraPrIDRef", "0"),
            new XAttribute("styleIDRef", "0"),
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

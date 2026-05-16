using System.IO.Compression;
using System.Text;
using OpenMcdf;
using PolyDonky.Core;

namespace PolyDonky.Codecs.Hwp;

/// <summary>
/// HWP (Hangul Word Processor 5.x, KS X 5700) → PolyDonkyument 리더.
/// OLE2 Compound File Binary 기반, zlib 압축, UTF-16 LE 인코딩.
/// </summary>
public sealed class HwpReader : IDocumentReader
{
    public string FormatId => "hwp";

    private readonly List<string> _fonts = new();
    private readonly Dictionary<int, CharShape> _charShapes = new();
    private readonly Dictionary<int, ParaShape> _paraShapes = new();

    public PolyDonkyument Read(Stream input)
    {
        ArgumentNullException.ThrowIfNull(input);

        using var cf = new CompoundFile(input);
        ValidateHwpFile(cf);

        var fileHeader = ReadFileHeader(cf);
        var docInfo = ReadDocInfo(cf, fileHeader.IsCompressed);
        var body = ReadBody(cf, fileHeader.IsCompressed);

        var doc = BuildDocument(docInfo, body);
        return doc;
    }

    private FileHeader ReadFileHeader(CompoundFile cf)
    {
        try
        {
            var fileHeaderEntry = cf.RootEntry.FindChild("FileHeader")
                ?? throw new InvalidOperationException("FileHeader not found");

            var data = new byte[256];
            using (var stream = cf.OpenStream((CFItem)fileHeaderEntry))
                stream.Read(data, 0, 256);

        // Signature (32 bytes)
        var sig = Encoding.ASCII.GetString(data, 0, 18).TrimEnd('\0');
        if (sig != "HWP Document File")
            throw new InvalidOperationException($"Invalid HWP signature: {sig}");

        // Version (bytes 32-35)
        uint version = BitConverter.ToUInt32(data, 32);
        var vMajor = (version >> 24) & 0xFF;
        var vMinor = (version >> 16) & 0xFF;
        var vMicro = (version >> 8) & 0xFF;
        var vBuild = version & 0xFF;

        // Flags (bytes 36-39)
        uint flags = BitConverter.ToUInt32(data, 36);
        bool compressed = (flags & 0x01) != 0;
        bool encrypted = (flags & 0x02) != 0;

        if (encrypted)
            throw new InvalidOperationException("Encrypted HWP files are not supported");

        return new FileHeader
        {
            Version = $"{vMajor}.{vMinor}.{vMicro}.{vBuild}",
            IsCompressed = compressed,
        };
    }

    private DocInfo ReadDocInfo(OleFile ole, bool compressed)
    {
        if (!ole.Exists("DocInfo"))
            throw new InvalidOperationException("DocInfo not found");

        var stream = ole.OpenStream("DocInfo");
        var data = ReadStream(stream, compressed);

        var docInfo = new DocInfo();
        var offset = 0;

        while (offset < data.Length)
        {
            if (!ParseRecord(data, ref offset, out var tag, out var level, out var payload))
                break;

            switch (tag)
            {
                case 0x010: // DOCUMENT_PROPERTIES
                    ParseDocumentProperties(payload, docInfo);
                    break;
                case 0x013: // FACE_NAME
                    ParseFaceName(payload, docInfo);
                    break;
                case 0x015: // CHAR_SHAPE
                    ParseCharShape(payload, docInfo);
                    break;
                case 0x019: // PARA_SHAPE
                    ParseParaShape(payload, docInfo);
                    break;
            }
        }

        return docInfo;
    }

    private Body ReadBody(OleFile ole, bool compressed)
    {
        var body = new Body();

        // Read all sections (Section0, Section1, ...)
        for (int i = 0; i < 100; i++)
        {
            var sectionPath = $"BodyText/Section{i}";
            if (!ole.Exists(sectionPath))
                break;

            var stream = ole.OpenStream(sectionPath);
            var data = ReadStream(stream, compressed);
            ParseSection(data, body);
        }

        return body;
    }

    private void ParseSection(byte[] data, Body body)
    {
        var offset = 0;

        Paragraph? currentPara = null;

        while (offset < data.Length)
        {
            if (!ParseRecord(data, ref offset, out var tag, out var level, out var payload))
                break;

            switch (tag)
            {
                case 0x034: // PARA_HEADER
                    if (currentPara != null)
                        body.Paragraphs.Add(currentPara);
                    currentPara = new Paragraph();
                    break;

                case 0x035: // PARA_TEXT
                    if (currentPara == null)
                        currentPara = new Paragraph();
                    try
                    {
                        var text = Encoding.Unicode.GetString(payload).TrimEnd('\0');
                        currentPara.Text = text;
                    }
                    catch { }
                    break;
            }
        }

        if (currentPara != null)
            body.Paragraphs.Add(currentPara);
    }

    private static bool ParseRecord(byte[] data, ref int offset, out ushort tag, out byte level, out byte[] payload)
    {
        tag = 0;
        level = 0;
        payload = Array.Empty<byte>();

        if (offset + 4 > data.Length)
            return false;

        // Tag ID (bytes 0-1, 10-bit) + Reserved (6-bit)
        tag = BitConverter.ToUInt16(data, offset);

        // Level (4-bit) + Size (12-bit)
        ushort info = BitConverter.ToUInt16(data, offset + 2);
        level = (byte)((info >> 12) & 0xF);
        int size = info & 0xFFF;

        offset += 4;

        // Extended size if needed
        if (size == 0xFFF)
        {
            if (offset + 4 > data.Length)
                return false;
            size = BitConverter.ToInt32(data, offset);
            offset += 4;
        }

        if (offset + size > data.Length)
            return false;

        payload = new byte[size];
        Array.Copy(data, offset, payload, 0, size);
        offset += size;

        return true;
    }

    private static byte[] ReadStream(OleStream stream, bool compressed)
    {
        using var ms = new MemoryStream();
        stream.CopyTo(ms);
        var data = ms.ToArray();

        if (compressed)
        {
            using var input = new MemoryStream(data);
            using var output = new MemoryStream();
            using (var deflate = new DeflateStream(input, CompressionMode.Decompress, leaveOpen: false))
                deflate.CopyTo(output);
            return output.ToArray();
        }

        return data;
    }

    private static void ParseDocumentProperties(byte[] payload, DocInfo docInfo)
    {
        if (payload.Length < 4)
            return;

        uint sections = BitConverter.ToUInt32(payload, 0);
        docInfo.SectionCount = (int)sections;
    }

    private void ParseFaceName(byte[] payload, DocInfo docInfo)
    {
        try
        {
            var name = Encoding.Unicode.GetString(payload).TrimEnd('\0');
            _fonts.Add(name);
        }
        catch { }
    }

    private void ParseCharShape(byte[] payload, DocInfo docInfo)
    {
        // Store for later reference
        var shape = new CharShape { RawData = payload };
        _charShapes[_charShapes.Count] = shape;
    }

    private void ParseParaShape(byte[] payload, DocInfo docInfo)
    {
        var shape = new ParaShape { RawData = payload };
        _paraShapes[_paraShapes.Count] = shape;
    }

    private PolyDonkyument BuildDocument(DocInfo docInfo, Body body)
    {
        var doc = new PolyDonkyument();
        var section = new Section();
        doc.Sections.Add(section);

        foreach (var para in body.Paragraphs)
        {
            var paragraph = new Paragraph();
            if (!string.IsNullOrEmpty(para.Text))
            {
                var run = new Run { Text = para.Text };
                paragraph.Runs.Add(run);
            }
            section.Blocks.Add(paragraph);
        }

        return doc;
    }

    private static void ValidateHwpFile(OleFile ole)
    {
        if (!ole.Exists("FileHeader"))
            throw new InvalidOperationException("Not a valid HWP file: FileHeader not found");
    }

    private sealed class FileHeader
    {
        public string Version { get; set; } = "";
        public bool IsCompressed { get; set; }
    }

    private sealed class DocInfo
    {
        public int SectionCount { get; set; }
    }

    private sealed class Body
    {
        public List<Paragraph> Paragraphs { get; } = new();
    }

    private sealed class Paragraph
    {
        public string Text { get; set; } = "";
    }

    private sealed class CharShape
    {
        public byte[] RawData { get; set; } = Array.Empty<byte>();
    }

    private sealed class ParaShape
    {
        public byte[] RawData { get; set; } = Array.Empty<byte>();
    }
}

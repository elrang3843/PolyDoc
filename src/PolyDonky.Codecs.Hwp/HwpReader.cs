using System.IO.Compression;
using System.Text;
using OpenMcdf;
using PolyDonky.Core;

namespace PolyDonky.Codecs.Hwp;

/// <summary>
/// HWP 5.x (KS X 5700) → PolyDonkyument 리더.
///
/// OLE2 Compound File Binary 컨테이너 구조:
///   FileHeader  — HWP 서명·버전·압축 플래그
///   DocInfo     — 폰트·글자서식·단락서식·스타일 (레코드 스트림, zlib 압축 가능)
///   BodyText/SectionN — 본문 (레코드 스트림, zlib 압축 가능)
///   PrvText     — 미리보기 텍스트 (UTF-16 LE)
///
/// 레코드 헤더 (4바이트 DWORD, 공식 KS X 5700 스펙):
///   bit  9~ 0: Tag ID  (10비트)
///   bit 19~10: Level   (10비트)
///   bit 31~20: Size    (12비트)
///   Size == 0xFFF → 다음 4바이트 uint32 가 실제 크기(확장).
///
/// Tag ID 베이스:
///   HWPTAG_BEGIN (=0x010) → DocInfo 태그 시작
///   BodyText 태그: PARA_HEADER=0x034, PARA_TEXT=0x035, CTRL_HEADER=0x03A, LIST_HEADER=0x03B, …
///
/// 지원 범위(1단계 ingest):
///   텍스트(PARA_TEXT), 단락 구분(PARA_HEADER), 글상자·표 내부 텍스트(LIST_HEADER 재귀).
/// 미지원: 암호화, 도형 좌표, 이미지, 변경추적.
/// </summary>
public sealed class HwpReader : IDocumentReader
{
    public string FormatId => "hwp";

    // ── DocInfo Tag ID (HWPTAG_BEGIN = 0x010) ──────────────────────────────
    private const uint TAG_DOCUMENT_PROPERTIES = 0x010; // HWPTAG_BEGIN
    private const uint TAG_ID_MAPPINGS         = 0x011; // HWPTAG_BEGIN + 1
    private const uint TAG_BIN_DATA            = 0x012; // HWPTAG_BEGIN + 2
    private const uint TAG_FACE_NAME           = 0x013; // HWPTAG_BEGIN + 3
    private const uint TAG_BORDER_FILL         = 0x014;
    private const uint TAG_CHAR_SHAPE          = 0x015;
    private const uint TAG_PARA_SHAPE          = 0x019;
    private const uint TAG_STYLE               = 0x01A;

    // ── BodyText Tag ID ────────────────────────────────────────────────────
    private const uint TAG_PARA_HEADER     = 0x034;
    private const uint TAG_PARA_TEXT       = 0x035;
    private const uint TAG_PARA_CHAR_SHAPE = 0x036;
    private const uint TAG_CTRL_HEADER     = 0x03A;
    private const uint TAG_LIST_HEADER     = 0x03B;
    private const uint TAG_PAGE_DEF        = 0x03C;
    private const uint TAG_TABLE           = 0x040;

    // ──────────────────────────────────────────────────────────────────────
    public PolyDonkyument Read(Stream input)
    {
        ArgumentNullException.ThrowIfNull(input);

        // OpenMcdf v3 RootStorage.OpenRead 는 파일 경로 전용 → 임시 파일 경유
        var tmpPath = Path.GetTempFileName();
        try
        {
            using (var fs = File.Create(tmpPath))
                input.CopyTo(fs);

            using var root = RootStorage.OpenRead(tmpPath);

            var header  = ParseFileHeader(root);
            var docInfo = ParseDocInfo(root, header.IsCompressed);
            var body    = ParseBodyText(root, header.IsCompressed);
            return BuildDocument(docInfo, body);
        }
        finally
        {
            try { File.Delete(tmpPath); } catch { }
        }
    }

    // ── FileHeader ─────────────────────────────────────────────────────────

    private static HwpFileHeader ParseFileHeader(RootStorage root)
    {
        using var stream = root.OpenStream("FileHeader");
        Span<byte> buf = stackalloc byte[256];
        int read = stream.Read(buf);
        if (read < 40)
            throw new InvalidOperationException("FileHeader too short");

        var sig = Encoding.ASCII.GetString(buf[..18]).TrimEnd('\0');
        if (sig != "HWP Document File")
            throw new InvalidOperationException($"Invalid HWP signature: {sig}");

        // 버전 DWORD: byte[32..35], LE → Major=byte[35], Minor=byte[34], Micro=byte[33], Build=byte[32]
        uint flags = BitConverter.ToUInt32(buf[36..]);
        if ((flags & 0x02) != 0)
            throw new InvalidOperationException("Encrypted HWP files are not supported");

        return new HwpFileHeader { IsCompressed = (flags & 0x01) != 0 };
    }

    // ── DocInfo ────────────────────────────────────────────────────────────

    private static HwpDocInfo ParseDocInfo(RootStorage root, bool isCompressed)
    {
        using var stream = root.OpenStream("DocInfo");
        var data = ReadAllBytes(stream);
        if (isCompressed) data = Decompress(data);

        var info = new HwpDocInfo();
        ForEachRecord(data, (tagId, level, payload) =>
        {
            switch (tagId)
            {
                case TAG_DOCUMENT_PROPERTIES when payload.Length >= 2:
                    info.SectionCount = BitConverter.ToUInt16(payload, 0);
                    break;
                case TAG_FACE_NAME:
                    try { info.FontNames.Add(Encoding.Unicode.GetString(payload).TrimEnd('\0')); }
                    catch { }
                    break;
            }
        });
        return info;
    }

    // ── BodyText ───────────────────────────────────────────────────────────

    private static HwpBodyText ParseBodyText(RootStorage root, bool isCompressed)
    {
        var body = new HwpBodyText();

        // TryOpenStorage: Storage 는 IDisposable 이 아니므로 using 불가
        if (!root.TryOpenStorage("BodyText", out var bodyDir))
            return body;

        for (int i = 0; i < 512; i++)
        {
            if (!bodyDir.TryOpenStream($"Section{i}", out var sectionStream))
                break;

            using (sectionStream)
            {
                var data = ReadAllBytes(sectionStream);
                if (isCompressed) data = Decompress(data);
                ParseSectionRecords(data, body);
            }
        }

        return body;
    }

    private static void ParseSectionRecords(byte[] data, HwpBodyText body)
    {
        HwpParagraph? current = null;

        ForEachRecord(data, (tagId, level, payload) =>
        {
            switch (tagId)
            {
                case TAG_PARA_HEADER:
                    if (current != null) body.Paragraphs.Add(current);
                    current = new HwpParagraph();
                    break;

                case TAG_PARA_TEXT:
                    if (current == null) current = new HwpParagraph();
                    try
                    {
                        // UTF-16 LE, 제어코드(0x00~0x1F 일부)가 섞임 — null만 제거
                        var text = Encoding.Unicode.GetString(payload).Replace("\0", "");
                        current.Text += text;
                    }
                    catch { }
                    break;

                case TAG_LIST_HEADER:
                    // 글상자·표 셀 등 내부 목록 시작 — 현재 단락 확정 후 새 컨텍스트
                    if (current != null)
                    {
                        body.Paragraphs.Add(current);
                        current = null;
                    }
                    break;
            }
        });

        if (current != null) body.Paragraphs.Add(current);
    }

    // ── 레코드 순회 ─────────────────────────────────────────────────────────

    /// <summary>
    /// HWP 레코드 스트림 순회.
    /// 헤더 DWORD:  bit 9-0 = Tag ID,  bit 19-10 = Level,  bit 31-20 = Size.
    /// Size == 0xFFF → 다음 uint32 가 실제 크기.
    /// </summary>
    private static void ForEachRecord(byte[] data, Action<uint, uint, byte[]> callback)
    {
        int offset = 0;
        while (offset + 4 <= data.Length)
        {
            uint dword = BitConverter.ToUInt32(data, offset);
            offset += 4;

            uint tagId = dword & 0x3FFu;           // bit  9~ 0
            uint level = (dword >> 10) & 0x3FFu;  // bit 19~10
            uint size  = dword >> 20;              // bit 31~20

            if (size == 0xFFFu)
            {
                if (offset + 4 > data.Length) break;
                size = BitConverter.ToUInt32(data, offset);
                offset += 4;
            }

            if (offset + (int)size > data.Length) break;

            var payload = new byte[size];
            Array.Copy(data, offset, payload, 0, (int)size);
            offset += (int)size;

            callback(tagId, level, payload);
        }
    }

    // ── 문서 모델 구성 ──────────────────────────────────────────────────────

    private static PolyDonkyument BuildDocument(HwpDocInfo docInfo, HwpBodyText body)
    {
        var doc     = new PolyDonkyument();
        var section = new Section();
        doc.Sections.Add(section);

        foreach (var hwpPara in body.Paragraphs)
        {
            if (string.IsNullOrWhiteSpace(hwpPara.Text)) continue;

            var paragraph = new Core.Paragraph();
            // HWP 단락 내 0x0D(CR)는 줄바꿈 구분자
            var lines = hwpPara.Text.Split('\r', StringSplitOptions.RemoveEmptyEntries);
            foreach (var line in lines)
            {
                var trimmed = line.Trim('\n');
                if (!string.IsNullOrWhiteSpace(trimmed))
                    paragraph.Runs.Add(new Run { Text = trimmed });
            }
            if (paragraph.Runs.Count > 0)
                section.Blocks.Add(paragraph);
        }

        return doc;
    }

    // ── 유틸리티 ────────────────────────────────────────────────────────────

    private static byte[] ReadAllBytes(CfbStream stream)
    {
        using var ms = new MemoryStream();
        stream.CopyTo(ms);
        return ms.ToArray();
    }

    private static byte[] Decompress(byte[] data)
    {
        using var input  = new MemoryStream(data);
        using var output = new MemoryStream();
        using (var deflate = new DeflateStream(input, CompressionMode.Decompress, leaveOpen: false))
            deflate.CopyTo(output);
        return output.ToArray();
    }

    // ── 내부 전용 모델 (Core 모델과 이름 충돌 방지를 위해 Hwp 접두사) ──────

    private sealed class HwpFileHeader
    {
        public bool IsCompressed { get; set; }
    }

    private sealed class HwpDocInfo
    {
        public int SectionCount { get; set; }
        public List<string> FontNames { get; } = new();
    }

    private sealed class HwpBodyText
    {
        public List<HwpParagraph> Paragraphs { get; } = new();
    }

    private sealed class HwpParagraph
    {
        public string Text { get; set; } = "";
    }
}

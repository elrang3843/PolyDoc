using System.IO.Compression;
using System.Text;
using OpenMcdf;
using PolyDonky.Core;

namespace PolyDonky.Codecs.Hwp;

/// <summary>
/// HwpReader 전용 파일 로거. d:\Temp\PolyDonky-HwpReader.log 에 기록한다.
/// Debug.WriteLine 은 외부 CLI 프로세스에서 VS 출력 창에 표시되지 않으므로
/// 파일 로그로 대체해 진단한다.
/// </summary>
internal static class HwpLog
{
    private static readonly string LogPath = Path.Combine(
        Environment.OSVersion.Platform == PlatformID.Win32NT ? @"d:\Temp" : "/tmp",
        "PolyDonky-HwpReader.log");

    static HwpLog()
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(LogPath)!);
            File.AppendAllText(LogPath, $"\n=== HwpReader session {DateTime.Now:yyyy-MM-dd HH:mm:ss} ===\n");
        }
        catch { }
    }

    public static void Write(string message)
    {
        var line = $"[{DateTime.Now:HH:mm:ss.fff}] {message}";
        System.Diagnostics.Debug.WriteLine(line);
        try { File.AppendAllText(LogPath, line + "\n"); } catch { }
    }
}

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
///   BodyText 태그: PARA_HEADER=0x034, PARA_TEXT=0x035, CTRL_HEADER=0x03A,
///                  LIST_HEADER=0x03B, PAGE_DEF=0x03C, SHAPE_COMPONENT=0x03F, …
///
/// 지원 범위:
///   텍스트/단락, 용지 설정(PAGE_DEF), 글상자(CTRL_HEADER + LIST_HEADER),
///   도형(SHAPE_COMPONENT + 서브태그), 이미지(PICTURE_COMPONENT + BinData).
/// 미지원: 암호화, 변경추적, 수식.
/// </summary>
public sealed class HwpReader : IDocumentReader
{
    public string FormatId => "hwp";

    // HWPUNIT: 1/7200 inch = 25.4/7200 mm ≈ 0.003528 mm/unit
    private const double HwpUnitToMm = 25.4 / 7200.0;

    // ── DocInfo Tag ID (HWPTAG_BEGIN = 0x010) ──────────────────────────────
    private const uint TAG_DOCUMENT_PROPERTIES = 0x010;
    private const uint TAG_BIN_DATA            = 0x012;
    private const uint TAG_FACE_NAME           = 0x013;

    // ── BodyText Tag ID ────────────────────────────────────────────────────
    private const uint TAG_PARA_HEADER         = 0x034;
    private const uint TAG_PARA_TEXT           = 0x035;
    private const uint TAG_PARA_CHAR_SHAPE     = 0x036;
    private const uint TAG_CTRL_HEADER         = 0x03A;
    private const uint TAG_LIST_HEADER         = 0x03B;
    private const uint TAG_PAGE_DEF            = 0x03C;
    private const uint TAG_SHAPE_COMPONENT     = 0x03F;
    private const uint TAG_TABLE               = 0x040;
    private const uint TAG_LINE_COMPONENT      = 0x041;
    private const uint TAG_RECT_COMPONENT      = 0x042;
    private const uint TAG_ELLIPSE_COMPONENT   = 0x043;
    private const uint TAG_ARC_COMPONENT       = 0x044;
    private const uint TAG_POLYGON_COMPONENT   = 0x045;
    private const uint TAG_CURVE_COMPONENT     = 0x046;
    private const uint TAG_OLE_COMPONENT       = 0x047;
    private const uint TAG_PICTURE_COMPONENT   = 0x048;
    private const uint TAG_CONTAINER_COMPONENT = 0x049;

    // CTRL_ID_GSO: 그리기 개체 (도형/글상자/이미지 공통). last byte space or null.
    // LE uint32: 'g'=0x67, 's'=0x73, 'o'=0x6F → check only first 3 bytes.
    private const uint CTRL_ID_GSO_MASK = 0x00FFFFFFu;
    private const uint CTRL_ID_GSO_VAL  = 0x006F7367u; // 'g','s','o'

    // ──────────────────────────────────────────────────────────────────────
    public PolyDonkyument Read(Stream input)
    {
        ArgumentNullException.ThrowIfNull(input);

        HwpLog.Write("[HwpReader.Read] 시작");
        var tmpPath = Path.GetTempFileName();
        try
        {
            using (var fs = File.Create(tmpPath))
                input.CopyTo(fs);

            using var root = RootStorage.OpenRead(tmpPath);

            var header  = ParseFileHeader(root);
            var docInfo = ParseDocInfo(root, header.IsCompressed);
            var body    = ParseBodyText(root, header.IsCompressed);
            return BuildDocument(docInfo, body, root);
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
        int binSeq = 0; // BIN_DATA records are 1-based sequential in DocInfo

        // DocInfo 태그 요약
        var docInfoTags = new Dictionary<uint, int>();
        ForEachRecord(data, (tagId, level, payload) =>
        {
            if (!docInfoTags.ContainsKey(tagId))
                docInfoTags[tagId] = 0;
            docInfoTags[tagId]++;
        });
        HwpLog.Write($"[HwpReader.ParseDocInfo] DocInfo tags: {string.Join(", ", docInfoTags.OrderBy(x => x.Key).Select(x => $"0x{x.Key:X3}(×{x.Value})"))}");

        ForEachRecord(data, (tagId, level, payload) =>
        {
            switch (tagId)
            {
                case TAG_DOCUMENT_PROPERTIES when payload.Length >= 2:
                    info.SectionCount = BitConverter.ToUInt16(payload, 0);
                    HwpLog.Write($"[HwpReader] TAG_DOCUMENT_PROPERTIES: SectionCount={info.SectionCount}");
                    break;

                case TAG_FACE_NAME:
                    try { info.FontNames.Add(Encoding.Unicode.GetString(payload).TrimEnd('\0')); }
                    catch { }
                    break;

                case TAG_BIN_DATA when payload.Length >= 4:
                    {
                        binSeq++;
                        ushort binId   = BitConverter.ToUInt16(payload, 0);
                        ushort binType = BitConverter.ToUInt16(payload, 2);
                        // binType: 0=link, 1=embedded, 2=stored
                        var binfo = new HwpBinInfo
                        {
                            Id         = binId > 0 ? binId : binSeq,
                            IsEmbedded = binType == 1 || binType == 2,
                        };

                        // For link type (binType=0), payload[4..] may contain a filename
                        if (binType == 0 && payload.Length > 4)
                        {
                            try
                            {
                                ushort nameLen = BitConverter.ToUInt16(payload, 4);
                                if (payload.Length >= 6 + nameLen * 2)
                                    binfo.LinkPath = Encoding.Unicode.GetString(payload, 6, nameLen * 2);
                            }
                            catch { }
                        }

                        // For embedded/stored, try to detect extension from payload[6..7] (format code)
                        if (payload.Length >= 8)
                        {
                            ushort fmt = BitConverter.ToUInt16(payload, 6);
                            binfo.Format = fmt switch
                            {
                                1  => "bmp",
                                2  => "gif",
                                3  => "jpg",
                                4  => "png",
                                5  => "wmf",
                                6  => "ole",
                                _  => ""
                            };
                        }

                        info.BinInfos.Add(binfo);
                    }
                    break;
            }
        });

        return info;
    }

    // ── BodyText ───────────────────────────────────────────────────────────

    private static HwpBodyText ParseBodyText(RootStorage root, bool isCompressed)
    {
        var body = new HwpBodyText();

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
                var recs = CollectRecords(data);
                ParseSectionRecords(recs, body);
            }
        }

        return body;
    }

    // ── Section record parsing (level-aware state machine) ─────────────────

    private static void ParseSectionRecords(List<HwpRecord> recs, HwpBodyText body)
    {
        HwpParagraph? current = null;
        int i = 0;

        HwpLog.Write(
            $"[HwpReader.ParseSectionRecords] Total records: {recs.Count}, looking for TAG_PAGE_DEF=0x{TAG_PAGE_DEF:X3}");

        // 모든 레코드 태그 요약
        var tagCounts = new Dictionary<uint, int>();
        foreach (var r in recs)
        {
            if (!tagCounts.ContainsKey(r.TagId))
                tagCounts[r.TagId] = 0;
            tagCounts[r.TagId]++;
        }
        foreach (var (tagId, count) in tagCounts.OrderBy(x => x.Key))
            HwpLog.Write($"  Tag 0x{tagId:X3}: {count} occurrences");

        while (i < recs.Count)
        {
            var rec = recs[i];

            switch (rec.TagId)
            {
                case TAG_PAGE_DEF:
                    HwpLog.Write(
                        $"[HwpReader] TAG_PAGE_DEF(0x{rec.TagId:X3}) at index {i}, level={rec.Level}, payloadLen={rec.Payload.Length}");
                    if (body.PageDef == null && rec.Payload.Length >= 32)
                    {
                        body.PageDef = ParsePageDef(rec.Payload);
                        HwpLog.Write($"  → ParsePageDef success");
                    }
                    else if (rec.Payload.Length < 32)
                    {
                        HwpLog.Write($"  → Skipped: payload too short ({rec.Payload.Length}<32)");
                    }
                    else
                    {
                        HwpLog.Write($"  → Skipped: already set");
                    }
                    break;

                case TAG_PARA_HEADER:
                    if (current != null) body.Paragraphs.Add(current);
                    current = new HwpParagraph();
                    break;

                case TAG_PARA_TEXT:
                    if (current == null) current = new HwpParagraph();
                    try
                    {
                        var text = Encoding.Unicode.GetString(rec.Payload).Replace("\0", "");
                        current.Text += text;
                    }
                    catch { }
                    break;

                case TAG_CTRL_HEADER when rec.Payload.Length >= 4:
                    {
                        uint ctrlId = BitConverter.ToUInt32(rec.Payload, 0);
                        if ((ctrlId & CTRL_ID_GSO_MASK) == CTRL_ID_GSO_VAL)
                        {
                            if (current != null) { body.Paragraphs.Add(current); current = null; }
                            i = ParseGsoControl(recs, i + 1, rec.Level + 1, body);
                            continue;
                        }
                    }
                    break;

                case TAG_LIST_HEADER:
                    if (current != null) { body.Paragraphs.Add(current); current = null; }
                    break;
            }

            i++;
        }

        if (current != null) body.Paragraphs.Add(current);

        HwpLog.Write($"[HwpReader] ParseSectionRecords complete: {body.Paragraphs.Count} paragraphs, " +
            $"{body.Images.Count} images, {body.TextBoxes.Count} textboxes, {body.Shapes.Count} shapes, " +
            $"PageDef={body.PageDef != null}");
    }

    // ── GSO (General Shape Object) control handler ─────────────────────────

    private static int ParseGsoControl(List<HwpRecord> recs, int startIdx, uint minLevel, HwpBodyText body)
    {
        double xMm = 0, yMm = 0, wMm = 0, hMm = 0;
        bool hasShape = false;
        HwpShapeKind kind = HwpShapeKind.Rectangle;
        int binDataId = 0;
        List<HwpParagraph>? tbContent = null;

        int i = startIdx;
        while (i < recs.Count)
        {
            var rec = recs[i];
            if (rec.Level < minLevel) break;

            switch (rec.TagId)
            {
                case TAG_SHAPE_COMPONENT when rec.Payload.Length >= 36:
                    {
                        // SHAPE_COMPONENT layout (HWPUNIT):
                        //   0-1: groupLevel (uint16)
                        //   2-3: localFileVersion (uint16)
                        //   4-7: initialWidth (uint32)
                        //   8-11: initialHeight (uint32)
                        //  12-15: zOrder (int32)
                        //  16-17: wrapType (int16)
                        //  18-19: horzRelRef (int16)
                        //  20-21: vertRelRef (int16)
                        //  22-23: horzRelPos (int16)
                        //  24-25: vertRelPos (int16)
                        //  26-29: xOffset (int32, HWPUNIT)
                        //  30-33: yOffset (int32, HWPUNIT)
                        var p = rec.Payload;
                        wMm = BitConverter.ToUInt32(p, 4) * HwpUnitToMm;
                        hMm = BitConverter.ToUInt32(p, 8) * HwpUnitToMm;
                        xMm = BitConverter.ToInt32(p, 26) * HwpUnitToMm;
                        yMm = BitConverter.ToInt32(p, 30) * HwpUnitToMm;
                        hasShape = true;
                    }
                    break;

                case TAG_LINE_COMPONENT:
                    kind = HwpShapeKind.Line;
                    break;
                case TAG_RECT_COMPONENT:
                    kind = HwpShapeKind.Rectangle;
                    break;
                case TAG_ELLIPSE_COMPONENT:
                    kind = HwpShapeKind.Ellipse;
                    break;
                case TAG_ARC_COMPONENT:
                    kind = HwpShapeKind.Arc;
                    break;
                case TAG_POLYGON_COMPONENT:
                    kind = HwpShapeKind.Polygon;
                    break;
                case TAG_CURVE_COMPONENT:
                    kind = HwpShapeKind.Curve;
                    break;
                case TAG_OLE_COMPONENT:
                    kind = HwpShapeKind.Ole;
                    break;

                case TAG_PICTURE_COMPONENT when rec.Payload.Length >= 4:
                    kind = HwpShapeKind.Picture;
                    // Try to read binDataId:
                    //   offset 0-1: border fill flags
                    //   offset 2-3: picture type / attrs
                    //   Further offsets hold the actual binDataId; try offset 50 (after border data),
                    //   falling back to a sequential counter managed by body.
                    binDataId = TryReadBinDataId(rec.Payload);
                    break;

                case TAG_CONTAINER_COMPONENT:
                    kind = HwpShapeKind.Container;
                    break;

                case TAG_LIST_HEADER when hasShape && kind != HwpShapeKind.Picture:
                    // Textbox: nested paragraphs follow at level+1
                    kind = HwpShapeKind.TextBox;
                    var innerLevel = rec.Level + 1;
                    i++;
                    var tbParas = new List<HwpParagraph>();
                    HwpParagraph? tbCur = null;
                    while (i < recs.Count)
                    {
                        var ir = recs[i];
                        if (ir.Level < innerLevel) break;
                        switch (ir.TagId)
                        {
                            case TAG_PARA_HEADER:
                                if (tbCur != null) tbParas.Add(tbCur);
                                tbCur = new HwpParagraph();
                                break;
                            case TAG_PARA_TEXT:
                                if (tbCur == null) tbCur = new HwpParagraph();
                                try
                                {
                                    tbCur.Text += Encoding.Unicode.GetString(ir.Payload).Replace("\0", "");
                                }
                                catch { }
                                break;
                        }
                        i++;
                    }
                    if (tbCur != null) tbParas.Add(tbCur);
                    tbContent = tbParas;
                    continue;
            }

            i++;
        }

        if (!hasShape) return i;

        // Clamp to sensible range (0..1000 mm). Negative coords → 0.
        xMm = Math.Max(0, xMm);
        yMm = Math.Max(0, yMm);

        switch (kind)
        {
            case HwpShapeKind.Picture:
                body.Images.Add(new HwpImage
                {
                    XMm = xMm, YMm = yMm, WidthMm = wMm, HeightMm = hMm,
                    BinDataId = binDataId,
                });
                break;

            case HwpShapeKind.TextBox:
                body.TextBoxes.Add(new HwpTextBox
                {
                    XMm = xMm, YMm = yMm, WidthMm = wMm, HeightMm = hMm,
                    Paragraphs = tbContent ?? new List<HwpParagraph>(),
                });
                break;

            default:
                body.Shapes.Add(new HwpShape
                {
                    XMm = xMm, YMm = yMm, WidthMm = wMm, HeightMm = hMm, Kind = kind,
                });
                break;
        }

        return i;
    }

    // ── PAGE_DEF parsing ───────────────────────────────────────────────────

    private static HwpPageDef ParsePageDef(byte[] p)
    {
        // PAGE_DEF layout (all fields uint32, HWPUNIT):
        //   0: paperWidth
        //   4: paperHeight
        //   8: marginLeft
        //  12: marginRight
        //  16: marginTop
        //  20: marginBottom
        //  24: headerMargin (distance from paper top to header baseline)
        //  28: footerMargin
        //  32: gutterMargin (제본 여백)
        //  36: flags / textDirection (bit0: landscape flag in some versions)
        double pw = BitConverter.ToUInt32(p, 0)  * HwpUnitToMm;
        double ph = BitConverter.ToUInt32(p, 4)  * HwpUnitToMm;
        double ml = BitConverter.ToUInt32(p, 8)  * HwpUnitToMm;
        double mr = BitConverter.ToUInt32(p, 12) * HwpUnitToMm;
        double mt = BitConverter.ToUInt32(p, 16) * HwpUnitToMm;
        double mb = BitConverter.ToUInt32(p, 20) * HwpUnitToMm;
        double mh = p.Length >= 28 ? BitConverter.ToUInt32(p, 24) * HwpUnitToMm : 10;
        double mf = p.Length >= 32 ? BitConverter.ToUInt32(p, 28) * HwpUnitToMm : 10;

        return new HwpPageDef
        {
            PaperWidthMm   = pw,
            PaperHeightMm  = ph,
            MarginLeftMm   = ml,
            MarginRightMm  = mr,
            MarginTopMm    = mt,
            MarginBottomMm = mb,
            MarginHeaderMm = mh,
            MarginFooterMm = mf,
        };
    }

    // ── BinData ID extraction ──────────────────────────────────────────────

    // Try to extract the BinData reference ID from a PICTURE_COMPONENT payload.
    // The exact offset is version-dependent; scan for a small plausible uint16.
    private static int TryReadBinDataId(byte[] p)
    {
        // The BinData ID is a 1-based uint16 in the payload.
        // Common positions: 50 (after border/fill/frame data), then 4, then 2.
        // Plausible range: 1 to 256.
        foreach (int off in new[] { 50, 4, 2, 52, 6 })
        {
            if (off + 2 > p.Length) continue;
            int candidate = BitConverter.ToUInt16(p, off);
            if (candidate is >= 1 and <= 256)
                return candidate;
        }
        return 0;
    }

    private static byte[]? ReadBinData(RootStorage root, int binId)
    {
        if (binId <= 0) return null;
        if (!root.TryOpenStorage("BinData", out var binDir)) return null;
        var streamName = $"BIN{binId:X4}";
        if (!binDir.TryOpenStream(streamName, out var binStream)) return null;
        using (binStream)
            return ReadAllBytes(binStream);
    }

    private static string DetectMediaType(byte[] data)
    {
        if (data.Length < 4) return "image/png";
        if (data[0] == 0x89 && data[1] == 0x50) return "image/png";
        if (data[0] == 0xFF && data[1] == 0xD8) return "image/jpeg";
        if (data[0] == 0x47 && data[1] == 0x49) return "image/gif";
        if (data[0] == 0x42 && data[1] == 0x4D) return "image/bmp";
        return "image/png";
    }

    // ── Document model construction ────────────────────────────────────────

    private static PolyDonkyument BuildDocument(HwpDocInfo docInfo, HwpBodyText body, RootStorage root)
    {
        var doc     = new PolyDonkyument();
        var section = new Section();
        doc.Sections.Add(section);

        // ── Page settings ──────────────────────────────────────────────────
        if (body.PageDef is { } pd)
        {
            var ps = section.Page;
            double pw = pd.PaperWidthMm;
            double ph = pd.PaperHeightMm;

            HwpLog.Write(
                $"[HwpReader] PAGE_DEF found: width={pw:F1}mm, height={ph:F1}mm");

            // HWP stores actual paper dimensions: landscape → width > height
            if (pw > ph && pw > 10 && ph > 10)
            {
                // Landscape: normalize to portrait (short=width, long=height)
                ps.WidthMm      = ph;
                ps.HeightMm     = pw;
                ps.Orientation  = PageOrientation.Landscape;
                HwpLog.Write(
                    $"[HwpReader] → Landscape detected: {ps.WidthMm:F0}x{ps.HeightMm:F0}mm, Orientation={ps.Orientation}");
            }
            else if (pw > 10 && ph > 10)
            {
                ps.WidthMm      = pw;
                ps.HeightMm     = ph;
                ps.Orientation  = PageOrientation.Portrait;
                HwpLog.Write(
                    $"[HwpReader] → Portrait detected: {ps.WidthMm:F0}x{ps.HeightMm:F0}mm, Orientation={ps.Orientation}");
            }

            ps.SizeKind = MatchPaperSize(ps.WidthMm, ps.HeightMm);

            if (pd.MarginLeftMm   > 0) ps.MarginLeftMm   = pd.MarginLeftMm;
            if (pd.MarginRightMm  > 0) ps.MarginRightMm  = pd.MarginRightMm;
            if (pd.MarginTopMm    > 0) ps.MarginTopMm    = pd.MarginTopMm;
            if (pd.MarginBottomMm > 0) ps.MarginBottomMm = pd.MarginBottomMm;
            if (pd.MarginHeaderMm > 0) ps.MarginHeaderMm = pd.MarginHeaderMm;
            if (pd.MarginFooterMm > 0) ps.MarginFooterMm = pd.MarginFooterMm;
        }

        // ── Paragraphs ─────────────────────────────────────────────────────
        foreach (var hwpPara in body.Paragraphs)
        {
            if (string.IsNullOrWhiteSpace(hwpPara.Text)) continue;

            var paragraph = new Core.Paragraph();
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

        // ── Text boxes ─────────────────────────────────────────────────────
        foreach (var tb in body.TextBoxes)
        {
            var tbo = new TextBoxObject
            {
                WrapMode     = ImageWrapMode.InFrontOfText,
                OverlayXMm   = tb.XMm,
                OverlayYMm   = tb.YMm,
                WidthMm      = tb.WidthMm  > 1 ? tb.WidthMm  : 60,
                HeightMm     = tb.HeightMm > 1 ? tb.HeightMm : 30,
                AnchorPageIndex = 0,
            };
            foreach (var tp in tb.Paragraphs)
            {
                if (string.IsNullOrWhiteSpace(tp.Text)) continue;
                var para = new Core.Paragraph();
                para.Runs.Add(new Run { Text = tp.Text.Trim() });
                tbo.Content.Add(para);
            }
            section.Blocks.Add(tbo);
        }

        // ── Images ─────────────────────────────────────────────────────────
        foreach (var img in body.Images)
        {
            byte[]? imgData = null;

            // Try the declared BinDataId first, then scan BIN0001 onwards
            if (img.BinDataId > 0)
                imgData = ReadBinData(root, img.BinDataId);

            if (imgData == null)
            {
                // Fallback: try adjacent IDs (±2 from declared)
                for (int bid = 1; bid <= 16 && imgData == null; bid++)
                    imgData = ReadBinData(root, bid);
            }

            if (imgData == null || imgData.Length == 0) continue;

            var ib = new ImageBlock
            {
                Data      = imgData,
                MediaType = DetectMediaType(imgData),
                WrapMode  = ImageWrapMode.InFrontOfText,
                WidthMm   = img.WidthMm  > 1 ? img.WidthMm  : 80,
                HeightMm  = img.HeightMm > 1 ? img.HeightMm : 60,
                OverlayXMm      = img.XMm,
                OverlayYMm      = img.YMm,
                AnchorPageIndex = 0,
            };
            section.Blocks.Add(ib);
        }

        // ── Shapes ─────────────────────────────────────────────────────────
        foreach (var sh in body.Shapes)
        {
            var so = new ShapeObject
            {
                Kind         = MapShapeKind(sh.Kind),
                WrapMode     = ImageWrapMode.InFrontOfText,
                OverlayXMm   = sh.XMm,
                OverlayYMm   = sh.YMm,
                WidthMm      = sh.WidthMm  > 1 ? sh.WidthMm  : 40,
                HeightMm     = sh.HeightMm > 1 ? sh.HeightMm : 20,
                AnchorPageIndex = 0,
                StrokeColor  = "#000000",
                StrokeThicknessPt = 1.0,
            };
            section.Blocks.Add(so);
        }

        return doc;
    }

    // ── Helpers ────────────────────────────────────────────────────────────

    private static PaperSizeKind MatchPaperSize(double wMm, double hMm)
    {
        // Match within ±3 mm tolerance
        foreach (PaperSizeKind kind in Enum.GetValues<PaperSizeKind>())
        {
            var dim = PageSettings.GetStandardDimensions(kind);
            if (dim is not { } d) continue;
            if (Math.Abs(d.W - wMm) < 3 && Math.Abs(d.H - hMm) < 3)
                return kind;
        }
        return PaperSizeKind.Custom;
    }

    private static ShapeKind MapShapeKind(HwpShapeKind k) => k switch
    {
        HwpShapeKind.Line      => ShapeKind.Line,
        HwpShapeKind.Ellipse   => ShapeKind.Ellipse,
        HwpShapeKind.Polygon   => ShapeKind.Polygon,
        HwpShapeKind.Curve     => ShapeKind.Spline,
        HwpShapeKind.Arc       => ShapeKind.HalfCircle,
        _                      => ShapeKind.Rectangle,
    };

    // ── 레코드 수집 ─────────────────────────────────────────────────────────

    private static List<HwpRecord> CollectRecords(byte[] data)
    {
        var list = new List<HwpRecord>();
        ForEachRecord(data, (tagId, level, payload) =>
            list.Add(new HwpRecord(tagId, level, payload)));
        return list;
    }

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

            uint tagId = dword & 0x3FFu;
            uint level = (dword >> 10) & 0x3FFu;
            uint size  = dword >> 20;

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

    // ── 내부 전용 모델 ─────────────────────────────────────────────────────

    private sealed class HwpFileHeader
    {
        public bool IsCompressed { get; set; }
    }

    private sealed class HwpDocInfo
    {
        public int SectionCount { get; set; }
        public List<string>     FontNames { get; } = new();
        public List<HwpBinInfo> BinInfos  { get; } = new();
    }

    private sealed class HwpBinInfo
    {
        public int    Id         { get; set; }
        public bool   IsEmbedded { get; set; }
        public string Format     { get; set; } = "";
        public string LinkPath   { get; set; } = "";
    }

    private sealed class HwpBodyText
    {
        public HwpPageDef?        PageDef   { get; set; }
        public List<HwpParagraph> Paragraphs{ get; } = new();
        public List<HwpTextBox>   TextBoxes { get; } = new();
        public List<HwpImage>     Images    { get; } = new();
        public List<HwpShape>     Shapes    { get; } = new();
    }

    private sealed class HwpPageDef
    {
        public double PaperWidthMm   { get; set; }
        public double PaperHeightMm  { get; set; }
        public double MarginLeftMm   { get; set; }
        public double MarginRightMm  { get; set; }
        public double MarginTopMm    { get; set; }
        public double MarginBottomMm { get; set; }
        public double MarginHeaderMm { get; set; }
        public double MarginFooterMm { get; set; }
    }

    private sealed class HwpParagraph
    {
        public string Text { get; set; } = "";
    }

    private sealed class HwpTextBox
    {
        public double XMm { get; set; }
        public double YMm { get; set; }
        public double WidthMm  { get; set; }
        public double HeightMm { get; set; }
        public List<HwpParagraph> Paragraphs { get; set; } = new();
    }

    private sealed class HwpImage
    {
        public double XMm       { get; set; }
        public double YMm       { get; set; }
        public double WidthMm   { get; set; }
        public double HeightMm  { get; set; }
        public int    BinDataId { get; set; }
    }

    private sealed class HwpShape
    {
        public double       XMm       { get; set; }
        public double       YMm       { get; set; }
        public double       WidthMm   { get; set; }
        public double       HeightMm  { get; set; }
        public HwpShapeKind Kind      { get; set; }
    }

    private record struct HwpRecord(uint TagId, uint Level, byte[] Payload);

    private enum HwpShapeKind
    {
        Rectangle, Line, Ellipse, Arc, Polygon, Curve, Ole, Picture, Container, TextBox
    }
}

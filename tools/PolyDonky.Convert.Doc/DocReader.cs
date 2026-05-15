using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using PolyDonky.Core;

namespace PolyDonky.Convert.Doc;

/// <summary>
/// RTF (Rich Text Format) → IWPF 변환기.
/// 지원: 글자 서식·단락 서식·위첨자/아래첨자·들여쓰기·리스트·이미지·표·메타데이터·
///       도형(\shp, 위치·크기·종류·색상 아웃라인)·OLE 개체(\object, OpaqueBlock 보존).
/// v1.0.0 이후 계획: \shp 전체 속성(그림자·3D·곡선 경로 등) + \object 데이터 복원.
/// </summary>
public class DocReader
{
    private readonly List<string> _fontTable  = new();
    private readonly List<Color>  _colorTable = new();

    private const double TwipsToMm = 1.0 / 56.692;
    private const double TwipsToPt = 1.0 / 20.0;

    // ─────────────────────────────── public entry ───────────────────────────────

    public PolyDonkyument Read(Stream input)
    {
        _fontTable.Clear();
        _colorTable.Clear();

        string rtf;
        using (var sr = new StreamReader(input, Encoding.Default,
                   detectEncodingFromByteOrderMarks: true, leaveOpen: true))
            rtf = sr.ReadToEnd();

        ParseTables(rtf);
        var meta = ParseInfo(rtf);
        var doc  = ParseContent(rtf);
        doc.Metadata.Title    = meta.Title    ?? doc.Metadata.Title;
        doc.Metadata.Author   = meta.Author   ?? doc.Metadata.Author;
        doc.Metadata.Created  = meta.Created  ?? doc.Metadata.Created;
        doc.Metadata.Modified = meta.Modified ?? doc.Metadata.Modified;
        return doc;
    }

    // ──────────────────────── table / meta pre-pass ─────────────────────────────

    private void ParseTables(string rtf)
    {
        ParseFontTable(rtf);
        ParseColorTable(rtf);
        if (_colorTable.Count == 0) _colorTable.Add(Color.Black);
    }

    private void ParseFontTable(string rtf)
    {
        int idx = rtf.IndexOf(@"\fonttbl", StringComparison.Ordinal);
        if (idx < 0) return;
        var (gs, ge) = FindGroupBounds(rtf, idx);
        if (gs < 0) return;

        var dict = new Dictionary<int, string>();
        int pos  = gs + 1;
        while (pos < ge)
        {
            if (rtf[pos] != '{') { pos++; continue; }
            var (is_, ie) = FindGroupBounds(rtf, pos);
            if (ie < 0) break;
            var inner  = rtf.Substring(is_ + 1, ie - is_ - 1);
            var fmatch = Regex.Match(inner, @"\\f(\d+)");
            if (fmatch.Success && int.TryParse(fmatch.Groups[1].Value, out int fnum))
            {
                var nm = Regex.Match(inner, @"\s([A-Za-z][^\\{};]*?);?\s*$");
                if (!nm.Success) nm = Regex.Match(inner, @"\s([^\\\{\};]+?);?\s*$");
                dict[fnum] = nm.Success ? nm.Groups[1].Value.Trim().TrimEnd(';') : "Arial";
            }
            pos = ie + 1;
        }
        if (dict.Count > 0)
        {
            int max = 0; foreach (var k in dict.Keys) if (k > max) max = k;
            for (int i = 0; i <= max; i++)
                _fontTable.Add(dict.TryGetValue(i, out var fn) ? fn : "Arial");
        }
    }

    private void ParseColorTable(string rtf)
    {
        int idx = rtf.IndexOf(@"\colortbl", StringComparison.Ordinal);
        if (idx < 0) return;
        var (gs, ge) = FindGroupBounds(rtf, idx);
        if (gs < 0) return;

        var seg = rtf.Substring(gs, ge - gs + 1);
        int start = 0;
        while (start < seg.Length)
        {
            int semi = seg.IndexOf(';', start);
            if (semi < 0) break;
            var part = seg.Substring(start, semi - start);
            var rm = Regex.Match(part, @"\\red(\d+)");
            var gm = Regex.Match(part, @"\\green(\d+)");
            var bm = Regex.Match(part, @"\\blue(\d+)");
            _colorTable.Add(rm.Success && gm.Success && bm.Success
                ? new Color((byte)int.Parse(rm.Groups[1].Value),
                            (byte)int.Parse(gm.Groups[1].Value),
                            (byte)int.Parse(bm.Groups[1].Value))
                : Color.Black);
            start = semi + 1;
        }
    }

    private sealed class MetaInfo
    {
        public string?          Title    { get; set; }
        public string?          Author   { get; set; }
        public DateTimeOffset?  Created  { get; set; }
        public DateTimeOffset?  Modified { get; set; }
    }

    private static MetaInfo ParseInfo(string rtf)
    {
        var meta = new MetaInfo();
        int idx  = rtf.IndexOf(@"\info", StringComparison.Ordinal);
        if (idx < 0) return meta;
        var (gs, ge) = FindGroupBounds(rtf, idx);
        if (gs < 0) return meta;

        string seg = rtf.Substring(gs, ge - gs + 1);
        meta.Title  = ExtractInfoField(seg, "title");
        meta.Author = ExtractInfoField(seg, "author");

        var ctm = Regex.Match(seg, @"\\creatim\\yr(\d+)\\mo(\d+)\\dy(\d+)\\hr(\d+)\\min(\d+)");
        if (ctm.Success) meta.Created = ParseRtfDate(ctm);
        var rtm = Regex.Match(seg, @"\\revtim\\yr(\d+)\\mo(\d+)\\dy(\d+)\\hr(\d+)\\min(\d+)");
        if (rtm.Success) meta.Modified = ParseRtfDate(rtm);

        return meta;
    }

    private static string? ExtractInfoField(string seg, string field)
    {
        var m = Regex.Match(seg, $@"\{{\\{field}\s(.+?)}}");
        return m.Success ? m.Groups[1].Value.Trim() : null;
    }

    private static DateTimeOffset ParseRtfDate(Match m)
    {
        int yr = int.Parse(m.Groups[1].Value), mo = int.Parse(m.Groups[2].Value),
            dy = int.Parse(m.Groups[3].Value), hr = int.Parse(m.Groups[4].Value),
            mn = int.Parse(m.Groups[5].Value);
        return new DateTimeOffset(yr, Math.Clamp(mo, 1, 12), Math.Clamp(dy, 1, 28),
                                  hr, mn, 0, TimeSpan.Zero);
    }

    // ────────────────────────── main content pass ───────────────────────────────

    private PolyDonkyument ParseContent(string rtf)
    {
        var doc     = new PolyDonkyument();
        var section = new Section();
        doc.Sections.Add(section);

        var stateStack = new Stack<RtfState>();
        var cur        = new RtfState();

        // 표 누적
        TableAccum?  tableAcc   = null;
        RowAccum?    rowAcc     = null;
        CellAccum?   cellAcc    = null;
        List<int>    cellWidths = new();  // \cellx 값 목록 (현재 행)

        // 단락 누적
        ParaAccum    para = new();
        var          text = new StringBuilder();

        // 그룹 스킵
        int  skipDepth        = 0;
        int  headerSkipDepth  = 0;
        bool nextGroupIgnore  = false;
        bool inPict           = false;
        int  pictDepth        = 0;
        var  pictData         = new StringBuilder();

        int pos = 0, n = rtf.Length;

        while (pos < n)
        {
            char c = rtf[pos];

            // ── { ─────────────────────────────────────────────────────────────
            if (c == '{')
            {
                if (nextGroupIgnore || skipDepth > 0)
                { skipDepth++; nextGroupIgnore = false; pos++; continue; }
                if (headerSkipDepth > 0) { headerSkipDepth++; pos++; continue; }
                if (inPict) { pictData.Append(c); pos++; continue; }

                FlushText(text, cur, para);
                stateStack.Push(cur.Clone());
                pos++;
            }
            // ── } ─────────────────────────────────────────────────────────────
            else if (c == '}')
            {
                if (skipDepth > 0)       { skipDepth--;      pos++; continue; }
                if (headerSkipDepth > 0) { headerSkipDepth--; pos++; continue; }
                if (inPict)
                {
                    pictDepth--;
                    if (pictDepth <= 0) { inPict = false; }
                    else pictData.Append(c);
                    pos++; continue;
                }
                FlushText(text, cur, para);
                if (stateStack.Count > 0) cur = stateStack.Pop();
                pos++;
            }
            // ── \ ─────────────────────────────────────────────────────────────
            else if (c == '\\')
            {
                pos++;
                if (pos >= n) break;
                char nc = rtf[pos];

                if (skipDepth > 0 || headerSkipDepth > 0)
                { SkipControlWord(rtf, ref pos); continue; }

                if (inPict)
                { SkipControlWord(rtf, ref pos); continue; }  // pict 안 제어어 스킵

                if (nc == '*') { nextGroupIgnore = true; pos++; }
                else if (nc == '\n' || nc == '\r')
                {
                    FlushText(text, cur, para);
                    CommitPara(ref para, cur, ref tableAcc, ref rowAcc, ref cellAcc, section);
                    para = NewPara(cur);
                    pos++;
                }
                else if (nc == '\'')
                {
                    pos++;
                    if (pos + 1 < n && int.TryParse(rtf.Substring(pos, 2),
                            System.Globalization.NumberStyles.HexNumber, null, out int code))
                    {
                        try { text.Append(Encoding.GetEncoding(1252).GetString(new[] { (byte)code })); }
                        catch { text.Append((char)code); }
                    }
                    pos += 2;
                }
                else if (nc == 'u' && pos + 1 < n && (char.IsDigit(rtf[pos + 1]) || rtf[pos + 1] == '-'))
                {
                    pos++;
                    bool neg = pos < n && rtf[pos] == '-'; if (neg) pos++;
                    var nb = new StringBuilder();
                    while (pos < n && char.IsDigit(rtf[pos])) { nb.Append(rtf[pos]); pos++; }
                    if (pos < n && rtf[pos] == '?') pos++;
                    if (nb.Length > 0 && short.TryParse(nb.ToString(), out short uc))
                        text.Append((char)(ushort)(neg ? -uc : uc));
                }
                else if (char.IsLetter(nc))
                {
                    var wb = new StringBuilder();
                    while (pos < n && char.IsLetter(rtf[pos])) { wb.Append(rtf[pos]); pos++; }
                    var word = wb.ToString();

                    bool hasN = false; bool negN = false; int num = 0;
                    if (pos < n && rtf[pos] == '-') { negN = true; pos++; }
                    if (pos < n && char.IsDigit(rtf[pos]))
                    {
                        hasN = true;
                        while (pos < n && char.IsDigit(rtf[pos]))
                        { num = num * 10 + (rtf[pos] - '0'); pos++; }
                        if (negN) num = -num;
                    }
                    else if (negN) pos--;
                    if (pos < n && rtf[pos] == ' ') pos++;

                    // 헤더 그룹
                    if (word is "fonttbl" or "colortbl" or "stylesheet" or "info"
                             or "listable" or "listtable" or "listoverridetable"
                             or "rsidtbl" or "generator" or "themedata"
                             or "colorschememapping" or "latentstyles" or "xmlnstbl")
                    { headerSkipDepth = 1; continue; }

                    // 이미지 그룹 시작
                    if (word == "pict")
                    {
                        inPict     = true;
                        pictDepth  = 1;
                        pictData.Clear();
                        continue;
                    }

                    // 도형 그룹 (\shp) — 위치·크기·종류·색상 아웃라인 수준 파싱
                    if (word == "shp")
                    {
                        int bracePos = rtf.LastIndexOf('{', pos - word.Length - 2);
                        if (bracePos >= 0)
                        {
                            var (_, ge) = FindGroupBounds(rtf, bracePos);
                            if (ge > 0)
                            {
                                var shape = ParseShapeGroup(rtf.Substring(bracePos, ge - bracePos + 1));
                                if (shape is not null)
                                {
                                    FlushText(text, cur, para);
                                    CommitPara(ref para, cur, ref tableAcc, ref rowAcc, ref cellAcc, section);
                                    section.Blocks.Add(shape);
                                    para = NewPara(cur);
                                }
                                pos = ge + 1;
                                if (stateStack.Count > 0) cur = stateStack.Pop();
                                continue;
                            }
                        }
                        // 그룹 경계를 못 찾으면 일반 skip
                        skipDepth = 1; continue;
                    }

                    // OLE 개체 그룹 (\object) — OpaqueBlock으로 보존
                    if (word == "object")
                    {
                        int bracePos = rtf.LastIndexOf('{', pos - word.Length - 2);
                        if (bracePos >= 0)
                        {
                            var (_, ge) = FindGroupBounds(rtf, bracePos);
                            if (ge > 0)
                            {
                                var grpText = rtf.Substring(bracePos, ge - bracePos + 1);
                                string oleKind = grpText.Contains(@"\objemb") ? "ole-embedded"
                                              : grpText.Contains(@"\objhtml") ? "ole-html"
                                              : grpText.Contains(@"\objocx")  ? "ole-ocx"
                                              : "ole-unknown";
                                var opaque = new OpaqueBlock
                                {
                                    Format       = "rtf",
                                    Kind         = oleKind,
                                    Xml          = grpText,
                                    DisplayLabel = "[OLE 개체]",
                                };
                                FlushText(text, cur, para);
                                CommitPara(ref para, cur, ref tableAcc, ref rowAcc, ref cellAcc, section);
                                section.Blocks.Add(opaque);
                                para = NewPara(cur);
                                pos = ge + 1;
                                if (stateStack.Count > 0) cur = stateStack.Pop();
                                continue;
                            }
                        }
                        skipDepth = 1; continue;
                    }

                    FlushText(text, cur, para);
                    ApplyWord(word, hasN, num, ref cur, ref para,
                              ref tableAcc, ref rowAcc, ref cellAcc,
                              ref cellWidths, section, text);
                }
                else
                {
                    switch (nc)
                    {
                        case '-': text.Append('­'); break;
                        case '~': text.Append(' '); break;
                        case '_': text.Append('‑'); break;
                        case '{': case '}': case '\\': text.Append(nc); break;
                    }
                    pos++;
                }
            }
            // ── 텍스트 ─────────────────────────────────────────────────────────
            else
            {
                if (skipDepth == 0 && headerSkipDepth == 0)
                {
                    if (inPict)
                    {
                        // pict 내 hex 데이터만 수집
                        if ((c >= '0' && c <= '9') || (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F'))
                            pictData.Append(c);
                    }
                    else text.Append(c);
                }
                pos++;
            }
        }

        FlushText(text, cur, para);
        CommitPara(ref para, cur, ref tableAcc, ref rowAcc, ref cellAcc, section);

        if (section.Blocks.Count == 0)
            section.Blocks.Add(new Paragraph());

        return doc;
    }

    // ── 제어 단어 적용 ────────────────────────────────────────────────────────

    private void ApplyWord(string w, bool hasN, int n,
        ref RtfState s, ref ParaAccum p,
        ref TableAccum? tbl, ref RowAccum? row, ref CellAccum? cell,
        ref List<int> cellWidths, Section section, StringBuilder textBuf)
    {
        switch (w)
        {
            // ─ 단락 제어 ─
            case "par":
            case "line":
                CommitPara(ref p, s, ref tbl, ref row, ref cell, section);
                p = NewPara(s);
                break;

            case "pard":
                s.Alignment = Alignment.Left;
                s.IndentLeft = 0; s.IndentRight = 0; s.IndentFirst = 0;
                s.SpaceBefore = 0; s.SpaceAfter = 0; s.LineHeight = 0;
                s.InTable = s.InTable;  // \intbl 별도 처리
                SyncParaState(s, ref p);
                break;

            case "intbl":
                s.InTable = true;
                break;

            case "ql": s.Alignment = Alignment.Left;    SyncParaState(s, ref p); break;
            case "qc": s.Alignment = Alignment.Center;  SyncParaState(s, ref p); break;
            case "qr": s.Alignment = Alignment.Right;   SyncParaState(s, ref p); break;
            case "qj": s.Alignment = Alignment.Justify; SyncParaState(s, ref p); break;

            case "sb": if (hasN) { s.SpaceBefore = n * TwipsToPt; SyncParaState(s, ref p); } break;
            case "sa": if (hasN) { s.SpaceAfter  = n * TwipsToPt; SyncParaState(s, ref p); } break;
            case "sl": if (hasN && n > 0) { s.LineHeight = n / 240.0; SyncParaState(s, ref p); } break;

            case "li": if (hasN) { s.IndentLeft  = n * TwipsToMm; SyncParaState(s, ref p); } break;
            case "ri": if (hasN) { s.IndentRight = n * TwipsToMm; SyncParaState(s, ref p); } break;
            case "fi": if (hasN) { s.IndentFirst = n * TwipsToMm; SyncParaState(s, ref p); } break;

            // ─ 글자 서식 ─
            case "b":          s.Bold          = !hasN || n != 0; break;
            case "i":          s.Italic        = !hasN || n != 0; break;
            case "ul":         s.Underline     = !hasN || n != 0; break;
            case "ulnone":     s.Underline     = false; break;
            case "strike":
            case "striked":    s.Strikethrough = !hasN || n != 0; break;
            case "super":      s.Superscript   = true; s.Subscript = false; break;
            case "sub":        s.Subscript     = true; s.Superscript = false; break;
            case "nosupersub": s.Superscript   = false; s.Subscript = false; break;
            case "fs":         if (hasN && n > 0) s.FontSize = n / 2.0; break;
            case "f":          if (hasN && n >= 0 && n < _fontTable.Count) s.FontIndex = n; break;
            case "cf":         s.ForeIdx = hasN ? Math.Max(0, n) : 0; break;
            case "cb":
            case "highlight":  s.BackIdx = hasN ? Math.Max(0, n) : 0; break;

            // ─ 탭 ─
            case "tab":
                FlushTextToRun(textBuf.ToString(), s, p); textBuf.Clear();
                FlushTextToRun("\t", s, p);
                break;

            // ─ 표 제어 ─
            case "trowd":
                if (tbl is null) tbl = new TableAccum();
                row = new RowAccum();
                cellWidths = new List<int>();
                break;

            case "cellx":
                if (hasN) cellWidths.Add(n);
                break;

            case "cell":
                // 셀 종료: 현재 단락 플러시
                if (cell is null) cell = new CellAccum();
                if (p.Runs.Count > 0 || true)
                {
                    var blk = BuildParagraph(p);
                    if (blk is not null) cell.Blocks.Add(blk);
                    p = NewPara(s);
                }
                if (row is null) row = new RowAccum();
                row.Cells.Add(cell);
                cell = new CellAccum();
                break;

            case "trrh":
                if (row is not null && hasN && n > 0)
                    row.HeightMm = n * TwipsToMm;
                break;

            case "row":
                if (tbl is not null && row is not null)
                {
                    tbl.Rows.Add(row);
                    row = null; cell = null; cellWidths = new();
                }
                // 표 종료는 \pard 이후 \trowd 없으면 CommitTable 호출
                break;

            case "trql": case "trqc": case "trqr":
                if (tbl is not null) tbl.HAlign = w == "trqc" ? TableHAlign.Center
                                                 : w == "trqr" ? TableHAlign.Right
                                                 : TableHAlign.Left;
                break;

            case "clvertalt": case "clvertalc": case "clvertalb":
                // 다음 셀 수직 정렬 — 현재 cell에 반영
                if (cell is null) cell = new CellAccum();
                cell.VerticalAlign = w == "clvertalc" ? CellVerticalAlign.Middle
                                   : w == "clvertalb" ? CellVerticalAlign.Bottom
                                   : CellVerticalAlign.Top;
                break;
        }
    }

    // ── 단락/표 커밋 ────────────────────────────────────────────────────────────

    private static void CommitPara(
        ref ParaAccum p, RtfState s,
        ref TableAccum? tbl, ref RowAccum? row, ref CellAccum? cell,
        Section section)
    {
        if (s.InTable)
        {
            // 표 내부 단락 — cellAcc에 임시 보관
            if (cell is null) cell = new CellAccum();
            var blk = BuildParagraph(p);
            if (blk is not null) cell.Blocks.Add(blk);
        }
        else
        {
            // 열린 표가 있으면 먼저 닫음
            if (tbl is not null)
            {
                if (row?.Cells.Count > 0) tbl.Rows.Add(row);
                section.Blocks.Add(BuildTable(tbl));
                tbl = null; row = null; cell = null;
            }
            var blk = BuildParagraph(p);
            if (blk is not null) section.Blocks.Add(blk);
        }
    }

    private static Paragraph? BuildParagraph(ParaAccum p)
    {
        if (p.Runs.Count == 0) return null;
        var para = new Paragraph
        {
            Style = new ParagraphStyle
            {
                Alignment        = p.Alignment,
                SpaceBeforePt    = p.SpaceBefore,
                SpaceAfterPt     = p.SpaceAfter,
                LineHeightFactor = p.LineHeight > 0 ? p.LineHeight : 0,
                IndentLeftMm     = p.IndentLeft,
                IndentRightMm    = p.IndentRight,
                IndentFirstLineMm = p.IndentFirst,
            }
        };
        foreach (var r in p.Runs)
            para.Runs.Add(new Run { Text = r.Text, Style = r.Style });
        return para;
    }

    private static Table BuildTable(TableAccum ta)
    {
        var tbl = new Table { HAlign = ta.HAlign };
        foreach (var ra in ta.Rows)
        {
            var row = new TableRow { HeightMm = ra.HeightMm };
            foreach (var ca in ra.Cells)
            {
                var cell = new TableCell { VerticalAlign = ca.VerticalAlign };
                foreach (var blk in ca.Blocks)
                    cell.Blocks.Add(blk);
                row.Cells.Add(cell);
            }
            tbl.Rows.Add(row);
        }
        return tbl;
    }

    // ── FlushText ────────────────────────────────────────────────────────────────

    private void FlushText(StringBuilder buf, RtfState s, ParaAccum p)
    {
        if (buf.Length == 0) return;
        var t = buf.ToString(); buf.Clear();
        FlushTextToRun(t, s, p);
    }

    private void FlushTextToRun(string t, RtfState s, ParaAccum p)
    {
        if (string.IsNullOrWhiteSpace(t)) return;  // 공백만이면 스킵

        var run = MakeRun(t, s);
        if (run is null) return;

        if (p.Runs.Count > 0)
        {
            var last = p.Runs[^1];
            if (StyleEquals(last.Style, run.Style)) { last.Text += t; return; }
        }
        p.Runs.Add(run);
    }

    private RunAccum? MakeRun(string text, RtfState s)
    {
        var style = new RunStyle
        {
            Bold          = s.Bold,
            Italic        = s.Italic,
            Underline     = s.Underline,
            Strikethrough = s.Strikethrough,
            Superscript   = s.Superscript,
            Subscript     = s.Subscript,
            FontSizePt    = s.FontSize > 0 ? s.FontSize : 11,
            FontFamily    = s.FontIndex >= 0 && s.FontIndex < _fontTable.Count
                                ? _fontTable[s.FontIndex] : null,
        };
        if (s.ForeIdx > 0 && s.ForeIdx < _colorTable.Count)
            style.Foreground = _colorTable[s.ForeIdx];
        if (s.BackIdx > 0 && s.BackIdx < _colorTable.Count)
            style.Background = _colorTable[s.BackIdx];

        return new RunAccum { Text = text, Style = style };
    }

    private static bool StyleEquals(RunStyle a, RunStyle b)
        => a.Bold == b.Bold && a.Italic == b.Italic && a.Underline == b.Underline
        && a.Strikethrough == b.Strikethrough && a.Superscript == b.Superscript
        && a.Subscript == b.Subscript && a.FontSizePt == b.FontSizePt
        && a.FontFamily == b.FontFamily
        && a.Foreground == b.Foreground && a.Background == b.Background;

    private static void SyncParaState(RtfState s, ref ParaAccum p)
    {
        p.Alignment   = s.Alignment;
        p.SpaceBefore = s.SpaceBefore;
        p.SpaceAfter  = s.SpaceAfter;
        p.LineHeight  = s.LineHeight;
        p.IndentLeft  = s.IndentLeft;
        p.IndentRight = s.IndentRight;
        p.IndentFirst = s.IndentFirst;
    }

    private static ParaAccum NewPara(RtfState s) => new()
    {
        Alignment   = s.Alignment,
        SpaceBefore = s.SpaceBefore,
        SpaceAfter  = s.SpaceAfter,
        LineHeight  = s.LineHeight,
        IndentLeft  = s.IndentLeft,
        IndentRight = s.IndentRight,
        IndentFirst = s.IndentFirst,
    };

    // ── 도형 파싱 (아웃라인) ────────────────────────────────────────────────────

    private static ShapeObject? ParseShapeGroup(string grp)
    {
        var shape = new ShapeObject();

        // 위치·크기 (twips → mm)
        int left = 0, top = 0, right = 0, bottom = 0;
        if (TryExtractInt(grp, @"\\shpleft(-?\d+)",   out int l)) left   = l;
        if (TryExtractInt(grp, @"\\shptop(-?\d+)",    out int t)) top    = t;
        if (TryExtractInt(grp, @"\\shpright(-?\d+)",  out int r)) right  = r;
        if (TryExtractInt(grp, @"\\shpbottom(-?\d+)", out int b)) bottom = b;

        shape.OverlayXMm = left   * TwipsToMm;
        shape.OverlayYMm = top    * TwipsToMm;
        shape.WidthMm    = Math.Max(10, (right  - left) * TwipsToMm);
        shape.HeightMm   = Math.Max(5,  (bottom - top)  * TwipsToMm);

        // 도형 종류 (\sp{\sn shapeType}{\sv N})
        var stm = Regex.Match(grp, @"\\sn\s+shapeType.*?\\sv\s+(-?\d+)", RegexOptions.Singleline);
        if (stm.Success && int.TryParse(stm.Groups[1].Value, out int stVal))
            shape.Kind = RtfShapeTypeToKind(stVal);

        // 채우기 색상 (ABGR int → #RRGGBB)
        var fcm = Regex.Match(grp, @"\\sn\s+fillColor.*?\\sv\s+(-?\d+)", RegexOptions.Singleline);
        if (fcm.Success && int.TryParse(fcm.Groups[1].Value, out int fc))
            shape.FillColor = AbgrIntToHex(fc);

        // 선 색상
        var lcm = Regex.Match(grp, @"\\sn\s+lineColor.*?\\sv\s+(-?\d+)", RegexOptions.Singleline);
        if (lcm.Success && int.TryParse(lcm.Groups[1].Value, out int lc))
            shape.StrokeColor = AbgrIntToHex(lc);

        // 선 두께 (\sp{\sn lineWidth}{\sv N}), 단위 EMU → pt (914400 EMU = 72pt)
        var lwm = Regex.Match(grp, @"\\sn\s+lineWidth.*?\\sv\s+(\d+)", RegexOptions.Singleline);
        if (lwm.Success && int.TryParse(lwm.Groups[1].Value, out int lw) && lw > 0)
            shape.StrokeThicknessPt = lw / 12700.0;  // 12700 EMU = 1pt

        return shape;
    }

    private static ShapeKind RtfShapeTypeToKind(int shapeType) => shapeType switch
    {
        1  => ShapeKind.Rectangle,
        2  => ShapeKind.RoundedRect,
        3  => ShapeKind.Ellipse,
        4  => ShapeKind.Polygon,       // Diamond
        5  => ShapeKind.Triangle,
        6  => ShapeKind.Triangle,      // RightTriangle
        20 => ShapeKind.Line,
        22 => ShapeKind.Line,          // BentConnector (화살표 포함)
        75 => ShapeKind.Star,
        _  => ShapeKind.Rectangle,
    };

    private static string AbgrIntToHex(int abgr)
    {
        byte r = (byte)( abgr        & 0xFF);
        byte g = (byte)((abgr >>  8) & 0xFF);
        byte b = (byte)((abgr >> 16) & 0xFF);
        return $"#{r:X2}{g:X2}{b:X2}";
    }

    private static bool TryExtractInt(string text, string pattern, out int value)
    {
        var m = Regex.Match(text, pattern);
        if (m.Success && int.TryParse(m.Groups[1].Value, out value)) return true;
        value = 0; return false;
    }

    // ── 유틸 ────────────────────────────────────────────────────────────────────

    private static (int start, int end) FindGroupBounds(string rtf, int hint)
    {
        int start = hint;
        for (int i = hint; i >= 0; i--)
        {
            if (rtf[i] == '{') { start = i; break; }
        }
        int depth = 0, end = -1;
        for (int i = start; i < rtf.Length; i++)
        {
            if (rtf[i] == '{') depth++;
            else if (rtf[i] == '}') { depth--; if (depth == 0) { end = i; break; } }
        }
        return (start, end);
    }

    private static void SkipControlWord(string rtf, ref int pos)
    {
        if (pos >= rtf.Length) return;
        char nc = rtf[pos];
        if (nc == '\'') { pos += 3; return; }
        if (!char.IsLetter(nc)) { pos++; return; }
        while (pos < rtf.Length && char.IsLetter(rtf[pos])) pos++;
        if (pos < rtf.Length && rtf[pos] == '-') pos++;
        while (pos < rtf.Length && char.IsDigit(rtf[pos])) pos++;
        if (pos < rtf.Length && rtf[pos] == ' ') pos++;
    }

    // ── PolyDonkyument builder ───────────────────────────────────────────────────

    private static PolyDonkyument BuildDocument(List<Block> blocks)
    {
        var doc     = new PolyDonkyument();
        var section = new Section();
        doc.Sections.Add(section);
        foreach (var b in blocks) section.Blocks.Add(b);
        if (section.Blocks.Count == 0) section.Blocks.Add(new Paragraph());
        return doc;
    }

    // ─────────────────────────── internal types ─────────────────────────────────

    private sealed class RtfState
    {
        public bool       Bold          { get; set; }
        public bool       Italic        { get; set; }
        public bool       Underline     { get; set; }
        public bool       Strikethrough { get; set; }
        public bool       Superscript   { get; set; }
        public bool       Subscript     { get; set; }
        public double     FontSize      { get; set; } = 11;
        public int        FontIndex     { get; set; }
        public int        ForeIdx       { get; set; }
        public int        BackIdx       { get; set; }
        public Alignment  Alignment     { get; set; } = Alignment.Left;
        public double     SpaceBefore   { get; set; }
        public double     SpaceAfter    { get; set; }
        public double     LineHeight    { get; set; }
        public double     IndentLeft    { get; set; }
        public double     IndentRight   { get; set; }
        public double     IndentFirst   { get; set; }
        public bool       InTable       { get; set; }

        public RtfState Clone() => new()
        {
            Bold = Bold, Italic = Italic, Underline = Underline,
            Strikethrough = Strikethrough, Superscript = Superscript, Subscript = Subscript,
            FontSize = FontSize, FontIndex = FontIndex, ForeIdx = ForeIdx, BackIdx = BackIdx,
            Alignment = Alignment, SpaceBefore = SpaceBefore, SpaceAfter = SpaceAfter,
            LineHeight = LineHeight, IndentLeft = IndentLeft, IndentRight = IndentRight,
            IndentFirst = IndentFirst, InTable = InTable,
        };
    }

    private sealed class ParaAccum
    {
        public Alignment      Alignment   { get; set; } = Alignment.Left;
        public double         SpaceBefore { get; set; }
        public double         SpaceAfter  { get; set; }
        public double         LineHeight  { get; set; }
        public double         IndentLeft  { get; set; }
        public double         IndentRight { get; set; }
        public double         IndentFirst { get; set; }
        public List<RunAccum> Runs        { get; } = new();
    }

    private sealed class RunAccum
    {
        public string   Text  { get; set; } = string.Empty;
        public RunStyle Style { get; set; } = new();
    }

    private sealed class TableAccum
    {
        public List<RowAccum> Rows   { get; } = new();
        public TableHAlign    HAlign { get; set; }
    }

    private sealed class RowAccum
    {
        public List<CellAccum> Cells    { get; } = new();
        public double          HeightMm { get; set; }
    }

    private sealed class CellAccum
    {
        public List<Block>     Blocks        { get; } = new();
        public CellVerticalAlign VerticalAlign { get; set; } = CellVerticalAlign.Top;
    }
}

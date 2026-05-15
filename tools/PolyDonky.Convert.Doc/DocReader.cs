using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using PolyDonky.Core;

namespace PolyDonky.Convert.Doc;

/// <summary>
/// RTF (Rich Text Format) → IWPF 변환기.
/// 텍스트·서식(굵기/기울임/밑줄/취소선/폰트/색상/크기/정렬/줄간격) 지원.
/// </summary>
public class DocReader
{
    private readonly List<string> _fontTable = new();
    private readonly List<Color>  _colorTable = new();

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
        var paras = ParseContent(rtf);
        return BuildDocument(paras);
    }

    // ──────────────────────────── table pre-pass ────────────────────────────────

    private void ParseTables(string rtf)
    {
        ParseFontTable(rtf);
        ParseColorTable(rtf);

        if (_colorTable.Count == 0)
            _colorTable.Add(Color.Black);
    }

    private void ParseFontTable(string rtf)
    {
        int idx = rtf.IndexOf(@"\fonttbl", StringComparison.Ordinal);
        if (idx < 0) return;

        var (gs, ge) = FindGroupBounds(rtf, idx);
        if (gs < 0) return;

        var dict = new Dictionary<int, string>();
        int pos = gs + 1;

        while (pos < ge)
        {
            if (rtf[pos] != '{') { pos++; continue; }
            var (is_, ie) = FindGroupBounds(rtf, pos);
            if (ie < 0) break;

            var inner = rtf.Substring(is_ + 1, ie - is_ - 1);
            var fmatch = Regex.Match(inner, @"\\f(\d+)");
            if (fmatch.Success && int.TryParse(fmatch.Groups[1].Value, out int fnum))
            {
                // 폰트 이름: 마지막 공백 이후, ; 제거
                var nmatch = Regex.Match(inner, @"[^\\{};]+\s([A-Za-z][^\\{};]*?);?\s*$");
                if (!nmatch.Success)
                    nmatch = Regex.Match(inner, @"\s([^\\\{\};]+?);?\s*$");
                var fname = nmatch.Success
                    ? nmatch.Groups[1].Value.Trim().TrimEnd(';')
                    : "Arial";
                dict[fnum] = fname;
            }
            pos = ie + 1;
        }

        if (dict.Count > 0)
        {
            int max = 0;
            foreach (var k in dict.Keys) if (k > max) max = k;
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
        // 각 ';' 사이가 하나의 색상 슬롯
        int start = 0;
        while (start < seg.Length)
        {
            int semi = seg.IndexOf(';', start);
            if (semi < 0) break;
            var part = seg.Substring(start, semi - start);
            var rm = Regex.Match(part, @"\\red(\d+)");
            var gm = Regex.Match(part, @"\\green(\d+)");
            var bm = Regex.Match(part, @"\\blue(\d+)");
            if (rm.Success && gm.Success && bm.Success)
            {
                _colorTable.Add(new Color(
                    (byte)int.Parse(rm.Groups[1].Value),
                    (byte)int.Parse(gm.Groups[1].Value),
                    (byte)int.Parse(bm.Groups[1].Value)));
            }
            else
            {
                // 빈 항목 (슬롯 0 = 기본 전경색)
                _colorTable.Add(Color.Black);
            }
            start = semi + 1;
        }
    }

    // ─────────────────────────── content pass ───────────────────────────────────

    private List<ParaAccum> ParseContent(string rtf)
    {
        var result = new List<ParaAccum>();
        var stateStack = new Stack<RtfState>();
        var cur  = new RtfState();
        var para = new ParaAccum();
        var text = new StringBuilder();

        // 무시할 그룹 깊이 추적
        int skipDepth = 0;
        // fonttbl/colortbl/info 등 헤더 그룹 깊이
        int headerSkipDepth = 0;
        bool nextGroupIgnore = false;

        int pos = 0;
        int n   = rtf.Length;

        while (pos < n)
        {
            char c = rtf[pos];

            if (c == '{')
            {
                if (nextGroupIgnore || skipDepth > 0)
                {
                    skipDepth++;
                    nextGroupIgnore = false;
                    pos++; continue;
                }
                if (headerSkipDepth > 0)
                {
                    headerSkipDepth++;
                    pos++; continue;
                }
                FlushText(text, cur, para);
                stateStack.Push(cur.Clone());
                pos++;
            }
            else if (c == '}')
            {
                if (skipDepth > 0)      { skipDepth--;      pos++; continue; }
                if (headerSkipDepth > 0){ headerSkipDepth--; pos++; continue; }
                FlushText(text, cur, para);
                if (stateStack.Count > 0) cur = stateStack.Pop();
                pos++;
            }
            else if (c == '\\')
            {
                pos++;
                if (pos >= n) break;
                char nc = rtf[pos];

                if (skipDepth > 0 || headerSkipDepth > 0)
                {
                    // 스킵 중에는 제어 단어 길이만 소비
                    SkipControlWord(rtf, ref pos);
                    continue;
                }

                if (nc == '*')
                {
                    // \* → 다음 그룹 무시
                    nextGroupIgnore = true;
                    pos++;
                }
                else if (nc == '\n' || nc == '\r')
                {
                    // 줄바꿈 → \par 처럼
                    FlushText(text, cur, para);
                    if (para.Runs.Count > 0 || result.Count == 0)
                    {
                        result.Add(para);
                        para = NewPara(cur);
                    }
                    pos++;
                }
                else if (nc == '\'')
                {
                    pos++;
                    if (pos + 1 < n)
                    {
                        if (int.TryParse(rtf.Substring(pos, 2),
                                System.Globalization.NumberStyles.HexNumber, null, out int code))
                        {
                            try { text.Append(Encoding.GetEncoding(1252).GetString(new[] { (byte)code })); }
                            catch { text.Append((char)code); }
                        }
                        pos += 2;
                    }
                }
                else if (nc == 'u' && pos + 1 < n && (char.IsDigit(rtf[pos + 1]) || rtf[pos + 1] == '-'))
                {
                    // \uN? Unicode
                    pos++;
                    bool neg = pos < n && rtf[pos] == '-';
                    if (neg) pos++;
                    var nb = new StringBuilder();
                    while (pos < n && char.IsDigit(rtf[pos])) { nb.Append(rtf[pos]); pos++; }
                    if (pos < n && rtf[pos] == '?') pos++;   // 대체 문자
                    if (nb.Length > 0 && short.TryParse(nb.ToString(), out short ucode))
                        text.Append((char)(ushort)(neg ? -ucode : ucode));
                }
                else if (char.IsLetter(nc))
                {
                    // 제어 단어
                    var wb = new StringBuilder();
                    while (pos < n && char.IsLetter(rtf[pos])) { wb.Append(rtf[pos]); pos++; }
                    var word = wb.ToString();

                    bool hasNum = false; bool negNum = false; int num = 0;
                    if (pos < n && rtf[pos] == '-') { negNum = true; pos++; }
                    if (pos < n && char.IsDigit(rtf[pos]))
                    {
                        hasNum = true;
                        while (pos < n && char.IsDigit(rtf[pos]))
                        { num = num * 10 + (rtf[pos] - '0'); pos++; }
                        if (negNum) num = -num;
                    }
                    else if (negNum) pos--; // '-' 반환
                    if (pos < n && rtf[pos] == ' ') pos++; // 구분 공백 소비

                    // 헤더 그룹 키워드
                    if (word is "fonttbl" or "colortbl" or "stylesheet" or "info"
                             or "listable" or "listtable" or "listoverridetable"
                             or "rsidtbl" or "generator" or "themedata"
                             or "colorschememapping" or "latentstyles" or "xmlnstbl")
                    {
                        headerSkipDepth = 1;
                        continue;
                    }

                    FlushText(text, cur, para);
                    ApplyWord(word, hasNum, num, ref cur, ref para, result);
                }
                else
                {
                    // 단일 제어 심볼
                    switch (nc)
                    {
                        case '-': text.Append('­'); break;
                        case '~': text.Append(' '); break;
                        case '_': text.Append('‑'); break;
                        case '\t': text.Append('\t'); break;
                        case '{':
                        case '}':
                        case '\\': text.Append(nc); break;
                    }
                    pos++;
                }
            }
            else
            {
                if (skipDepth == 0 && headerSkipDepth == 0)
                    text.Append(c);
                pos++;
            }
        }

        FlushText(text, cur, para);
        if (para.Runs.Count > 0)
            result.Add(para);

        return result;
    }

    private static void SkipControlWord(string rtf, ref int pos)
    {
        if (pos >= rtf.Length) return;
        char nc = rtf[pos];
        if (nc == '\'' ) { pos += 3; return; }
        if (!char.IsLetter(nc)) { pos++; return; }
        while (pos < rtf.Length && char.IsLetter(rtf[pos])) pos++;
        if (pos < rtf.Length && rtf[pos] == '-') pos++;
        while (pos < rtf.Length && char.IsDigit(rtf[pos])) pos++;
        if (pos < rtf.Length && rtf[pos] == ' ') pos++;
    }

    private void ApplyWord(string w, bool hasN, int n, ref RtfState s, ref ParaAccum p, List<ParaAccum> result)
    {
        switch (w)
        {
            case "par":
            case "line":
                result.Add(p);
                p = NewPara(s);
                break;
            case "pard":
                s.Alignment   = Alignment.Left;
                s.SpaceBefore = 0; s.SpaceAfter = 0; s.LineHeight = 0;
                p.Alignment   = s.Alignment;
                p.SpaceBefore = 0; p.SpaceAfter = 0; p.LineHeight = 0;
                break;
            case "ql": s.Alignment = Alignment.Left;    p.Alignment = s.Alignment; break;
            case "qc": s.Alignment = Alignment.Center;  p.Alignment = s.Alignment; break;
            case "qr": s.Alignment = Alignment.Right;   p.Alignment = s.Alignment; break;
            case "qj": s.Alignment = Alignment.Justify; p.Alignment = s.Alignment; break;
            case "sb": if (hasN) { s.SpaceBefore = n / 20.0; p.SpaceBefore = s.SpaceBefore; } break;
            case "sa": if (hasN) { s.SpaceAfter  = n / 20.0; p.SpaceAfter  = s.SpaceAfter;  } break;
            case "sl": if (hasN && n > 0) { s.LineHeight = n / 240.0; p.LineHeight = s.LineHeight; } break;
            case "b":          s.Bold          = !hasN || n != 0; break;
            case "i":          s.Italic        = !hasN || n != 0; break;
            case "ul":         s.Underline     = !hasN || n != 0; break;
            case "ulnone":     s.Underline     = false; break;
            case "strike":     s.Strikethrough = !hasN || n != 0; break;
            case "striked":    s.Strikethrough = !hasN || n != 0; break;
            case "fs":         if (hasN && n > 0) s.FontSize = n / 2.0; break;
            case "f":          if (hasN && n >= 0 && n < _fontTable.Count) s.FontIndex = n; break;
            case "cf":         s.ForeIdx = hasN ? n : 0; break;
            case "cb":
            case "highlight":  s.BackIdx = hasN ? n : 0; break;
            case "tab":        // FlushText는 이미 호출됨 — 탭 문자 삽입
                var tabRun = MakeRun("\t", s);
                if (tabRun is not null) p.Runs.Add(tabRun);
                break;
        }
    }

    // ───────────────────────────── helpers ──────────────────────────────────────

    private void FlushText(StringBuilder buf, RtfState s, ParaAccum p)
    {
        if (buf.Length == 0) return;
        var t = buf.ToString(); buf.Clear();

        // 앞뒤 공백 제거 (DocWriter가 제어 단어 구분용으로 추가한 공백)
        // 실제 의도된 공백이 사라질 수 있지만 RTF 특성상 불가피
        // → 완전히 공백만인 토큰은 버림
        if (string.IsNullOrWhiteSpace(t)) return;

        var run = MakeRun(t, s);
        if (run is null) return;

        // 마지막 Run과 서식이 같으면 병합
        if (p.Runs.Count > 0)
        {
            var last = p.Runs[^1];
            if (StyleEquals(last.Style, run.Style))
            { last.Text += t; return; }
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
        => a.Bold == b.Bold && a.Italic == b.Italic
        && a.Underline == b.Underline && a.Strikethrough == b.Strikethrough
        && a.FontSizePt == b.FontSizePt && a.FontFamily == b.FontFamily
        && a.Foreground == b.Foreground && a.Background == b.Background;

    private static ParaAccum NewPara(RtfState s) => new()
    {
        Alignment   = s.Alignment,
        SpaceBefore = s.SpaceBefore,
        SpaceAfter  = s.SpaceAfter,
        LineHeight  = s.LineHeight,
    };

    /// <summary>rtf[pos..] 또는 pos 앞쪽에서 그룹('{' ~ '}') 경계를 찾는다.</summary>
    private static (int start, int end) FindGroupBounds(string rtf, int hint)
    {
        // '{' 를 hint 위치부터 뒤로 탐색
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

    // ────────────────────────── document builder ────────────────────────────────

    private static PolyDonkyument BuildDocument(List<ParaAccum> paras)
    {
        var doc     = new PolyDonkyument();
        var section = new Section();
        doc.Sections.Add(section);

        foreach (var pa in paras)
        {
            var para = new Paragraph
            {
                Style = new ParagraphStyle
                {
                    Alignment       = pa.Alignment,
                    SpaceBeforePt   = pa.SpaceBefore,
                    SpaceAfterPt    = pa.SpaceAfter,
                    LineHeightFactor = pa.LineHeight > 0 ? pa.LineHeight : 0,
                }
            };
            foreach (var ra in pa.Runs)
                para.Runs.Add(new Run { Text = ra.Text, Style = ra.Style });

            if (para.Runs.Count > 0)
                section.Blocks.Add(para);
        }

        if (section.Blocks.Count == 0)
            section.Blocks.Add(new Paragraph());

        return doc;
    }

    // ─────────────────────────── internal types ─────────────────────────────────

    private sealed class RtfState
    {
        public bool       Bold          { get; set; }
        public bool       Italic        { get; set; }
        public bool       Underline     { get; set; }
        public bool       Strikethrough { get; set; }
        public double     FontSize      { get; set; } = 11;
        public int        FontIndex     { get; set; }
        public int        ForeIdx       { get; set; }
        public int        BackIdx       { get; set; }
        public Alignment  Alignment     { get; set; } = Alignment.Left;
        public double     SpaceBefore   { get; set; }
        public double     SpaceAfter    { get; set; }
        public double     LineHeight    { get; set; }

        public RtfState Clone() => new()
        {
            Bold = Bold, Italic = Italic, Underline = Underline, Strikethrough = Strikethrough,
            FontSize = FontSize, FontIndex = FontIndex, ForeIdx = ForeIdx, BackIdx = BackIdx,
            Alignment = Alignment, SpaceBefore = SpaceBefore, SpaceAfter = SpaceAfter,
            LineHeight = LineHeight,
        };
    }

    private sealed class ParaAccum
    {
        public Alignment        Alignment   { get; set; } = Alignment.Left;
        public double           SpaceBefore { get; set; }
        public double           SpaceAfter  { get; set; }
        public double           LineHeight  { get; set; }
        public List<RunAccum>   Runs        { get; } = new();
    }

    private sealed class RunAccum
    {
        public string   Text  { get; set; } = string.Empty;
        public RunStyle Style { get; set; } = new();
    }
}

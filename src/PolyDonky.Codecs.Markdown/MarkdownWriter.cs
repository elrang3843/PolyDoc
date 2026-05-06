using System.Linq;
using System.Text;
using PolyDonky.Core;

namespace PolyDonky.Codecs.Markdown;

/// <summary>
/// PolyDonkyument 를 CommonMark + GFM Markdown 으로 직렬화한다.
///
/// 매핑:
///   - OutlineLevel.H1~H6        → "# " ~ "###### "
///   - ParagraphStyle.QuoteLevel → "> " 접두어 (n 회 반복)
///   - ParagraphStyle.IsThematicBreak → "---"
///   - ParagraphStyle.CodeLanguage non-null → ```lang … ```
///   - ListMarker.Bullet         → "- "
///   - ListMarker.OrderedDecimal → "1. " (자동 번호)
///   - ListMarker.Checked        → "- [x] " / "- [ ] " (GFM 작업 목록)
///   - Run.Bold/Italic           → **/*
///   - Run.Strikethrough         → ~~text~~
///   - Run.Subscript / Superscript → ~sub~ / ^sup^
///   - Run.Url                   → [text](url)
///   - Mono FontFamily 만 가진 Run → `code`
///   - Run.LatexSource           → $latex$ (인라인) / $$…$$ (별행 = IsDisplayEquation)
///   - Table                     → GFM 파이프 표 (헤더 행 분리, 셀 정렬 적용)
///   - ImageBlock                → ![alt](path)
/// </summary>
public sealed class MarkdownWriter : IDocumentWriter
{
    public string FormatId => "md";

    public Encoding Encoding { get; init; } = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

    public void Write(PolyDonkyument document, Stream output)
    {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(output);

        using var writer = new StreamWriter(output, Encoding, leaveOpen: true) { NewLine = "\n" };
        writer.Write(ToMarkdown(document));
    }

    public static string ToMarkdown(PolyDonkyument document)
    {
        ArgumentNullException.ThrowIfNull(document);
        var sb = new StringBuilder();

        foreach (var section in document.Sections)
            WriteBlocks(sb, section.Blocks, indent: "");

        return sb.ToString();
    }

    // ── 블록 렌더링 ───────────────────────────────────────────────────

    private static void WriteBlocks(StringBuilder sb, IList<Block> blocks, string indent)
    {
        bool first = true;
        foreach (var block in blocks)
        {
            if (!first) sb.Append('\n');   // 블록 간 빈 줄
            first = false;

            switch (block)
            {
                case Paragraph p:
                    WriteParagraph(sb, p, indent);
                    break;

                case Table t:
                    WriteTable(sb, t, indent);
                    break;

                case ImageBlock img:
                    WriteImage(sb, img, indent);
                    break;
            }
        }
    }

    private static void WriteParagraph(StringBuilder sb, Paragraph p, string indent)
    {
        var quotePrefix = p.Style.QuoteLevel > 0
            ? string.Concat(Enumerable.Repeat("> ", p.Style.QuoteLevel))
            : string.Empty;

        if (p.Style.IsThematicBreak)
        {
            sb.Append(indent).Append(quotePrefix).Append("---\n");
            return;
        }

        if (p.Style.CodeLanguage is not null)
        {
            // 펜스드 코드 블록 — 코드 내 백틱 연속 최대 길이보다 1 더 긴 펜스 사용.
            var code     = p.GetPlainText();
            int maxBt    = LongestBacktickRun(code);
            int fenceLen = Math.Max(3, maxBt + 1);
            var fence    = new string('`', fenceLen);
            sb.Append(indent).Append(quotePrefix).Append(fence).Append(p.Style.CodeLanguage).Append('\n');
            foreach (var line in code.Split('\n'))
                sb.Append(indent).Append(quotePrefix).Append(line).Append('\n');
            sb.Append(indent).Append(quotePrefix).Append(fence).Append('\n');
            return;
        }

        // 별행 수식 — display equation 단일 run 이면 $$ 블록으로 출력 (한 줄 $$…$$ 는 inline 으로 파싱됨).
        if (p.Runs.Count == 1 && p.Runs[0].IsDisplayEquation && !string.IsNullOrEmpty(p.Runs[0].LatexSource))
        {
            sb.Append(indent).Append(quotePrefix).Append("$$\n");
            foreach (var line in p.Runs[0].LatexSource!.Split('\n'))
                sb.Append(indent).Append(quotePrefix).Append(line).Append('\n');
            sb.Append(indent).Append(quotePrefix).Append("$$\n");
            return;
        }

        // 리스트 들여쓰기.
        var listPrefix = string.Empty;
        if (p.Style.ListMarker is { } list)
        {
            string itemIndent = new(' ', list.Level * 2);
            string marker = list.Kind switch
            {
                ListKind.Bullet => list.Checked switch
                {
                    true  => "- [x] ",
                    false => "- [ ] ",
                    _     => "- ",
                },
                _ => list.Checked switch
                {
                    true  => $"{list.OrderedNumber ?? 1}. [x] ",
                    false => $"{list.OrderedNumber ?? 1}. [ ] ",
                    _     => $"{list.OrderedNumber ?? 1}. ",
                },
            };
            listPrefix = itemIndent + marker;
        }

        // 헤더.
        var headerPrefix = p.Style.Outline > OutlineLevel.Body
            ? new string('#', (int)p.Style.Outline) + " "
            : string.Empty;

        sb.Append(indent).Append(quotePrefix).Append(listPrefix).Append(headerPrefix);
        sb.Append(RenderRuns(p.Runs));
        sb.Append('\n');
    }

    private static void WriteImage(StringBuilder sb, ImageBlock img, string indent)
    {
        var alt  = EscapeText(img.Description ?? string.Empty);
        string path;
        if (!string.IsNullOrEmpty(img.ResourcePath))
        {
            path = img.ResourcePath;
        }
        else if (img.Data.Length > 0)
        {
            // 임베디드 바이너리 → data URI (Markdown 자체는 표준이 아니지만 브라우저·VSCode 등이 지원).
            var b64  = Convert.ToBase64String(img.Data);
            var mime = string.IsNullOrEmpty(img.MediaType) ? "application/octet-stream" : img.MediaType;
            path = $"data:{mime};base64,{b64}";
        }
        else
        {
            path = string.Empty;
        }
        sb.Append(indent).Append("![").Append(alt).Append("](").Append(path).Append(")\n");
    }

    private static void WriteTable(StringBuilder sb, Table t, string indent)
    {
        if (t.Rows.Count == 0) return;

        int cols = t.Columns.Count;
        if (cols == 0)
        {
            cols = t.Rows.Max(r => r.Cells.Count);
        }

        // 헤더 행 — 첫 번째 IsHeader 행 또는 첫 번째 행을 헤더로 처리.
        var headerRowIdx = 0;
        for (int i = 0; i < t.Rows.Count; i++)
        {
            if (t.Rows[i].IsHeader) { headerRowIdx = i; break; }
        }

        // 셀 텍스트 추출.
        string[] CellsOfRow(TableRow row)
        {
            var arr = new string[cols];
            for (int c = 0; c < cols; c++)
            {
                arr[c] = c < row.Cells.Count
                    ? EscapeCell(CellPlainText(row.Cells[c]))
                    : string.Empty;
            }
            return arr;
        }

        // 헤더.
        var hdr = CellsOfRow(t.Rows[headerRowIdx]);
        sb.Append(indent).Append("| ").Append(string.Join(" | ", hdr)).Append(" |\n");

        // 정렬 줄.
        sb.Append(indent).Append('|');
        for (int c = 0; c < cols; c++)
        {
            var align = c < t.Rows[headerRowIdx].Cells.Count
                ? t.Rows[headerRowIdx].Cells[c].TextAlign
                : CellTextAlign.Left;
            sb.Append(align switch
            {
                CellTextAlign.Center => " :---: |",
                CellTextAlign.Right  => " ---: |",
                _                    => " --- |",
            });
        }
        sb.Append('\n');

        // 본문 행.
        for (int i = 0; i < t.Rows.Count; i++)
        {
            if (i == headerRowIdx) continue;
            var rowCells = CellsOfRow(t.Rows[i]);
            sb.Append(indent).Append("| ").Append(string.Join(" | ", rowCells)).Append(" |\n");
        }
    }

    private static string CellPlainText(TableCell cell)
    {
        var sb = new StringBuilder();
        foreach (var b in cell.Blocks)
        {
            if (b is Paragraph p)
            {
                if (sb.Length > 0) sb.Append("<br>");
                sb.Append(RenderRuns(p.Runs));
            }
        }
        return sb.ToString();
    }

    // ── 인라인(Run) 렌더링 ────────────────────────────────────────────

    private static string RenderRuns(IList<Run> runs)
    {
        var sb = new StringBuilder();
        foreach (var run in runs)
            sb.Append(RenderRun(run));
        return sb.ToString();
    }

    private static string RenderRun(Run run)
    {
        // 수식.
        if (!string.IsNullOrEmpty(run.LatexSource))
        {
            return run.IsDisplayEquation
                ? $"$${run.LatexSource}$$"
                : $"${run.LatexSource}$";
        }

        var text = run.Text;
        var s    = run.Style;
        bool isMono = !string.IsNullOrEmpty(s.FontFamily) &&
                      s.FontFamily.Contains("monospace", StringComparison.OrdinalIgnoreCase);

        // 인라인 코드 — 다른 강조와 결합되지 않는 단순 케이스로 처리.
        if (isMono && !s.Bold && !s.Italic && !s.Strikethrough && !s.Subscript && !s.Superscript)
        {
            // 백틱이 텍스트에 있으면 더 긴 백틱 펜스 사용.
            int bt = LongestBacktickRun(text);
            string fence = new('`', bt + 1);
            string pad = text.StartsWith('`') || text.EndsWith('`') ? " " : "";
            var coded  = $"{fence}{pad}{text}{pad}{fence}";
            return run.Url is { Length: > 0 } u ? $"[{coded}]({EscapeUrl(u)})" : coded;
        }

        var escaped = EscapeText(text);

        // 위/아래 첨자, 취소선 — bold/italic 보다 안쪽에 둠.
        if (s.Subscript)     escaped = $"~{escaped}~";
        if (s.Superscript)   escaped = $"^{escaped}^";
        if (s.Strikethrough) escaped = $"~~{escaped}~~";

        // 굵게/기울임.
        if (s.Bold && s.Italic) escaped = $"***{escaped}***";
        else if (s.Bold)        escaped = $"**{escaped}**";
        else if (s.Italic)      escaped = $"*{escaped}*";

        // 하이퍼링크.
        if (run.Url is { Length: > 0 } url)
        {
            escaped = $"[{escaped}]({EscapeUrl(url)})";
        }

        return escaped;
    }

    // ── 이스케이프 ────────────────────────────────────────────────────

    private static string EscapeText(string text)
    {
        if (string.IsNullOrEmpty(text)) return text;
        var sb = new StringBuilder(text.Length);
        foreach (var ch in text)
        {
            // CommonMark ASCII punctuation 중 의미 있는 것만 이스케이프.
            // `^` 포함: EmphasisExtras extension 이 `^sup^` 구문을 활성화하므로 일반 텍스트의 ^ 를 보호.
            switch (ch)
            {
                case '\\': case '`': case '*': case '_':
                case '[':  case ']':
                case '<':  case '>':
                case '#':  case '~':
                case '|':  case '^':
                    sb.Append('\\').Append(ch);
                    break;
                default:
                    sb.Append(ch);
                    break;
            }
        }
        return sb.ToString();
    }

    private static string EscapeCell(string text)
    {
        // 표 셀에서는 줄바꿈을 <br> 로, 파이프는 \\| 로.
        return text.Replace("\n", "<br>").Replace("|", "\\|");
    }

    private static string EscapeUrl(string url)
    {
        // URL 안의 괄호와 공백을 이스케이프 / 인코딩.
        return url.Replace(" ", "%20").Replace(")", "\\)");
    }

    private static int LongestBacktickRun(string s)
    {
        int max = 0, cur = 0;
        foreach (var ch in s)
        {
            if (ch == '`') { cur++; if (cur > max) max = cur; }
            else cur = 0;
        }
        return max;
    }
}

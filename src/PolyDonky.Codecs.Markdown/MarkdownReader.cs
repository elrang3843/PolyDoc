using System.Linq;
using System.Text;
using Markdig;
using Markdig.Extensions.TaskLists;
using Markdig.Extensions.Mathematics;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using PolyDonky.Core;
using MarkdigTable     = Markdig.Extensions.Tables.Table;
using MarkdigTableRow  = Markdig.Extensions.Tables.TableRow;
using MarkdigTableCell = Markdig.Extensions.Tables.TableCell;
using PdBlock          = PolyDonky.Core.Block;
using PdThematicBreak  = PolyDonky.Core.ThematicBreakBlock;
using PdTable          = PolyDonky.Core.Table;
using PdTableRow       = PolyDonky.Core.TableRow;
using PdTableCell      = PolyDonky.Core.TableCell;

namespace PolyDonky.Codecs.Markdown;

/// <summary>
/// Markdig 기반 Markdown 리더 — CommonMark + GFM (advanced extensions) 풀 지원.
///
/// 지원 문법:
///   - ATX/Setext 헤더 H1~H6
///   - 단락, 소프트/하드 줄바꿈
///   - 강조: *italic*, **bold**, ***bold-italic***
///   - 취소선: ~~strike~~ (GFM)
///   - 위/아래 첨자: ^sup^ / ~sub~ (EmphasisExtras)
///   - 인라인 코드, 펜스드/들여쓰기 코드 블록 (info-string 언어 힌트 보존)
///   - 비순서/순서 리스트, 중첩, GFM 작업 목록 `- [x]` / `- [ ]`
///   - 인용 (>), 중첩 인용
///   - 가로선 (---, ___, ***)
///   - 링크: [text](url), 자동 링크 &lt;url&gt;
///   - 이미지: ![alt](path) → 단락 단독은 ImageBlock, 인라인 혼합은 fallback 텍스트
///   - GFM 파이프 표 + 셀 정렬
///   - 인라인 HTML / HTML 블록 → 모노스페이스 텍스트로 보존
///   - 수식: $inline$ → LaTeX 인라인 Run, $$display$$ → LaTeX 별행 단락
///   - 각주, 정의 리스트는 일반 단락으로 격하 (Core 모델 미지원).
/// </summary>
public sealed class MarkdownReader : IDocumentReader
{
    public string FormatId => "md";

    private static readonly MarkdownPipeline DefaultPipeline = new MarkdownPipelineBuilder()
        .UseAdvancedExtensions()
        .Build();

    public PolyDonkyument Read(Stream input)
    {
        ArgumentNullException.ThrowIfNull(input);
        using var reader = new StreamReader(input, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, leaveOpen: true);
        return FromMarkdown(reader.ReadToEnd());
    }

    public static PolyDonkyument FromMarkdown(string source)
    {
        ArgumentNullException.ThrowIfNull(source);
        var document = new PolyDonkyument();
        var section  = new Section();
        document.Sections.Add(section);

        var ast = Markdig.Markdown.Parse(source, DefaultPipeline);
        ProcessContainer(ast, section.Blocks, marker: null, quoteLevel: 0, listLevel: 0);
        return document;
    }

    // ── 블록 처리 ─────────────────────────────────────────────────────────

    private static void ProcessContainer(
        ContainerBlock container,
        IList<PdBlock>   target,
        ListMarker?    marker,
        int            quoteLevel,
        int            listLevel)
    {
        foreach (var block in container)
        {
            switch (block)
            {
                case HeadingBlock heading:
                {
                    var p = new Paragraph();
                    p.Style.Outline    = (OutlineLevel)Math.Clamp(heading.Level, 1, 6);
                    p.Style.ListMarker = CloneMarker(marker);
                    p.Style.QuoteLevel = quoteLevel;
                    AppendInline(p, heading.Inline);
                    target.Add(p);
                    break;
                }

                case ParagraphBlock para:
                {
                    // 단독 이미지 단락은 ImageBlock 으로 승격.
                    if (TryConvertSoloImageParagraph(para, marker, quoteLevel, target))
                        break;

                    var p = new Paragraph();
                    p.Style.ListMarker = CloneMarker(marker);
                    p.Style.QuoteLevel = quoteLevel;
                    AppendInline(p, para.Inline);
                    target.Add(p);
                    break;
                }

                case ListBlock list:
                {
                    ProcessList(list, target, quoteLevel, listLevel);
                    break;
                }

                case QuoteBlock quote:
                {
                    ProcessContainer(quote, target, marker, quoteLevel + 1, listLevel);
                    break;
                }

                // MathBlock extends FencedCodeBlock — 반드시 먼저 매치.
                case MathBlock mathBlock:
                {
                    var mp = new Paragraph();
                    mp.Style.QuoteLevel = quoteLevel;
                    mp.Style.ListMarker = CloneMarker(marker);
                    var latex = string.Join('\n',
                        mathBlock.Lines.Lines.Take(mathBlock.Lines.Count).Select(l => l.ToString()));
                    mp.Runs.Add(new Run
                    {
                        Text              = latex,
                        LatexSource       = latex,
                        IsDisplayEquation = true,
                        Style             = new RunStyle(),
                    });
                    target.Add(mp);
                    break;
                }

                case FencedCodeBlock fenced:
                {
                    target.Add(BuildCodeParagraph(fenced.Lines, marker, quoteLevel, fenced.Info ?? ""));
                    break;
                }

                case CodeBlock indented:
                {
                    target.Add(BuildCodeParagraph(indented.Lines, marker, quoteLevel, ""));
                    break;
                }

                case Markdig.Syntax.ThematicBreakBlock:
                {
                    target.Add(new PdThematicBreak());
                    break;
                }

                case HtmlBlock html:
                {
                    var hp = new Paragraph();
                    hp.Style.QuoteLevel = quoteLevel;
                    hp.Style.ListMarker = CloneMarker(marker);
                    var text = string.Join('\n',
                        html.Lines.Lines.Take(html.Lines.Count).Select(l => l.ToString()));
                    hp.AddText(text, MonoStyle());
                    target.Add(hp);
                    break;
                }

                case MarkdigTable mdTable:
                {
                    target.Add(BuildTable(mdTable));
                    break;
                }

                // 정의 리스트, 각주 그룹, 사용자 정의 컨테이너 — 자식 블록을 평탄화.
                case ContainerBlock cb:
                {
                    ProcessContainer(cb, target, marker, quoteLevel, listLevel);
                    break;
                }
            }
        }
    }

    private static void ProcessList(ListBlock list, IList<PdBlock> target, int quoteLevel, int parentListLevel)
    {
        int counter = 0;
        foreach (var item in list)
        {
            counter++;
            if (item is not ListItemBlock listItem) continue;

            // GFM 작업 목록 — 첫 ParagraphBlock 의 첫 Inline 이 TaskList 면 추출.
            bool? checkedState = null;
            if (listItem.Count > 0
                && listItem[0] is ParagraphBlock pb
                && pb.Inline?.FirstChild is TaskList tl)
            {
                checkedState = tl.Checked;
                tl.Remove();
            }

            var lm = list.IsOrdered
                ? new ListMarker
                {
                    Kind          = ListKind.OrderedDecimal,
                    OrderedNumber = listItem.Order > 0 ? listItem.Order : counter,
                    Level         = parentListLevel,
                    Checked       = checkedState,
                }
                : new ListMarker
                {
                    Kind    = ListKind.Bullet,
                    Level   = parentListLevel,
                    Checked = checkedState,
                };

            ProcessContainer(listItem, target, lm, quoteLevel, parentListLevel + 1);
        }
    }

    private static bool TryConvertSoloImageParagraph(
        ParagraphBlock para,
        ListMarker?    marker,
        int            quoteLevel,
        IList<PdBlock>   target)
    {
        if (para.Inline is null) return false;
        // 단독 이미지: 첫 Inline 이 IsImage 인 LinkInline 이고 다른 자식이 없을 때.
        var first = para.Inline.FirstChild;
        if (first is not LinkInline { IsImage: true } imgLink) return false;
        if (first.NextSibling is not null) return false;

        var alt = ExtractInlineText(imgLink);
        var img = new ImageBlock
        {
            Description  = string.IsNullOrEmpty(alt) ? null : alt,
            ResourcePath = imgLink.Url,
            MediaType    = GuessMediaType(imgLink.Url),
            HAlign       = ImageHAlign.Left,
            WrapMode     = ImageWrapMode.Inline,
        };
        target.Add(img);

        // 리스트/인용 컨텍스트에서 이미지 단독은 그 컨텍스트를 잃지만(블록 모델 한계) —
        // 나중에 인용/리스트 안 이미지 보존이 필요하면 별도 wrap 도입.
        _ = marker;
        _ = quoteLevel;
        return true;
    }

    private static PdTable BuildTable(MarkdigTable mdTable)
    {
        var t = new PdTable();

        int maxCols = 0;
        foreach (var rowBlock in mdTable.OfType<MarkdigTableRow>())
            maxCols = Math.Max(maxCols, rowBlock.Count);

        // 컬럼 정의 (Markdig 의 ColumnDefinitions 에 정렬 정보 있음)
        var aligns = new List<CellTextAlign>(maxCols);
        for (int i = 0; i < maxCols; i++)
        {
            t.Columns.Add(new TableColumn());
            var align = i < mdTable.ColumnDefinitions.Count
                ? mdTable.ColumnDefinitions[i].Alignment switch
                {
                    Markdig.Extensions.Tables.TableColumnAlign.Center => CellTextAlign.Center,
                    Markdig.Extensions.Tables.TableColumnAlign.Right  => CellTextAlign.Right,
                    _                                                 => CellTextAlign.Left,
                }
                : CellTextAlign.Left;
            aligns.Add(align);
        }

        foreach (var rowBlock in mdTable.OfType<MarkdigTableRow>())
        {
            var row = new PdTableRow { IsHeader = rowBlock.IsHeader };
            int colIdx = 0;
            foreach (var cellBlock in rowBlock.OfType<MarkdigTableCell>())
            {
                var cell = new PdTableCell
                {
                    TextAlign  = colIdx < aligns.Count ? aligns[colIdx] : CellTextAlign.Left,
                    ColumnSpan = cellBlock.ColumnSpan > 0 ? cellBlock.ColumnSpan : 1,
                    RowSpan    = cellBlock.RowSpan    > 0 ? cellBlock.RowSpan    : 1,
                };
                ProcessContainer(cellBlock, cell.Blocks, marker: null, quoteLevel: 0, listLevel: 0);
                if (cell.Blocks.Count == 0) cell.Blocks.Add(new Paragraph());
                row.Cells.Add(cell);
                colIdx++;
            }
            t.Rows.Add(row);
        }

        return t;
    }

    private static Paragraph BuildCodeParagraph(
        Markdig.Helpers.StringLineGroup lines,
        ListMarker?                    marker,
        int                            quoteLevel,
        string                         language)
    {
        var paragraph = new Paragraph();
        paragraph.Style.ListMarker   = CloneMarker(marker);
        paragraph.Style.QuoteLevel   = quoteLevel;
        paragraph.Style.CodeLanguage = language;

        var slices = lines.Lines.Take(lines.Count).Select(l => l.ToString());
        var code   = string.Join('\n', slices);
        paragraph.AddText(code, MonoStyle());
        return paragraph;
    }

    // ── 인라인 처리 ──────────────────────────────────────────────────────

    private static void AppendInline(Paragraph paragraph, ContainerInline? inline)
    {
        if (inline is null)
        {
            paragraph.AddText(string.Empty);
            return;
        }

        WalkInline(paragraph, inline, new RunStyle(), url: null);
        if (paragraph.Runs.Count == 0)
            paragraph.AddText(string.Empty);
    }

    private static void WalkInline(Paragraph paragraph, Inline node, RunStyle style, string? url)
    {
        switch (node)
        {
            case LiteralInline lit:
            {
                var s = Clone(style);
                paragraph.Runs.Add(new Run { Text = lit.Content.ToString(), Style = s, Url = url });
                break;
            }

            case EmphasisInline em:
            {
                var s = Clone(style);
                char delim = em.DelimiterChar;
                int  cnt   = em.DelimiterCount;
                if (delim == '~')
                {
                    if (cnt >= 2) s.Strikethrough = true;  // GFM ~~strike~~
                    else          s.Subscript    = true;  // ExtraEmphasis ~sub~
                }
                else if (delim == '^')
                {
                    s.Superscript = true;                  // ExtraEmphasis ^sup^
                }
                else if (cnt >= 2)
                {
                    s.Bold = true;
                }
                else
                {
                    s.Italic = true;
                }
                foreach (var c in em) WalkInline(paragraph, c, s, url);
                break;
            }

            case CodeInline code:
            {
                var s = Clone(style);
                s.FontFamily = "Consolas, D2Coding, monospace";
                paragraph.Runs.Add(new Run { Text = code.Content, Style = s, Url = url });
                break;
            }

            case LineBreakInline lb:
            {
                // 하드 브레이크는 줄바꿈, 소프트 브레이크는 공백.
                var ch = lb.IsHard ? "\n" : " ";
                paragraph.Runs.Add(new Run { Text = ch, Style = Clone(style), Url = url });
                break;
            }

            case AutolinkInline auto:
            {
                var s = Clone(style);
                s.Underline = true;
                paragraph.Runs.Add(new Run { Text = auto.Url, Style = s, Url = auto.Url });
                break;
            }

            case LinkInline link when link.IsImage:
            {
                // 인라인 이미지(혼합 단락) — 모델 한계로 텍스트 fallback.
                var alt = ExtractInlineText(link);
                paragraph.AddText($"[{alt}]({link.Url})", Clone(style));
                break;
            }

            case LinkInline link:
            {
                var s = Clone(style);
                s.Underline = true;
                foreach (var c in link) WalkInline(paragraph, c, s, link.Url);
                break;
            }

            case HtmlInline html:
            {
                paragraph.AddText(html.Tag, MonoStyle());
                break;
            }

            case HtmlEntityInline ent:
            {
                paragraph.AddText(ent.Transcoded.ToString(), Clone(style));
                break;
            }

            case MathInline math:
            {
                var s = Clone(style);
                paragraph.Runs.Add(new Run
                {
                    Text        = math.Content.ToString(),
                    LatexSource = math.Content.ToString(),
                    Style       = s,
                    Url         = url,
                });
                break;
            }

            case TaskList:
            {
                // 작업 목록 — ProcessList 에서 이미 추출되어 도달하지 않아야 함.
                break;
            }

            case ContainerInline container:
            {
                foreach (var c in container) WalkInline(paragraph, c, style, url);
                break;
            }
        }
    }

    private static string ExtractInlineText(ContainerInline container)
    {
        var sb = new StringBuilder();
        Walk(container);
        return sb.ToString();

        void Walk(Inline node)
        {
            switch (node)
            {
                case LiteralInline lit: sb.Append(lit.Content.ToString()); break;
                case CodeInline code:   sb.Append(code.Content);           break;
                case LineBreakInline:   sb.Append(' ');                     break;
                case ContainerInline c: foreach (var x in c) Walk(x);       break;
            }
        }
    }

    // ── 유틸 ────────────────────────────────────────────────────────────

    private static RunStyle MonoStyle() => new() { FontFamily = "Consolas, D2Coding, monospace" };

    // Core 의 정식 RunStyle.Clone() 사용.
    private static RunStyle Clone(RunStyle s) => s.Clone();

    // Core 의 정식 ListMarker.Clone() 사용.
    private static ListMarker? CloneMarker(ListMarker? m) => m?.Clone();

    private static string GuessMediaType(string? url)
    {
        if (string.IsNullOrEmpty(url)) return "application/octet-stream";
        var ext = Path.GetExtension(url).ToLowerInvariant();
        return ext switch
        {
            ".png"  => "image/png",
            ".jpg"  => "image/jpeg",
            ".jpeg" => "image/jpeg",
            ".gif"  => "image/gif",
            ".bmp"  => "image/bmp",
            ".tif"  => "image/tiff",
            ".tiff" => "image/tiff",
            ".webp" => "image/webp",
            ".svg"  => "image/svg+xml",
            _       => "application/octet-stream",
        };
    }
}

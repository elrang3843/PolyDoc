using System.Linq;
using System.Text;
using Markdig;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using PolyDoc.Core;

namespace PolyDoc.Codecs.Markdown;

/// <summary>
/// Markdig 기반 Markdown 리더. CommonMark 풀 파싱 후 PolyDoc 모델로 매핑한다.
///
/// Phase A 의 PolyDoc 모델 범위에 맞춰 다음을 처리한다:
///   - 단락 / 헤더 (H1~H6) / 순서 · 비순서 리스트
///   - 인라인 강조 (em / strong / 둘 다)
///   - 인라인 코드 (monospace 폰트 힌트로 표시)
///   - 코드블록 (fenced / indented) — 모노스페이스 단락으로 격하
///   - 인용 — 일반 단락으로 격하 (모델에 인용 블록 도입 후 향상)
/// 표·이미지·링크는 PolyDoc.Core 가 표/이미지/하이퍼링크 노드를 도입한 뒤 매핑한다.
/// </summary>
public sealed class MarkdownReader : IDocumentReader
{
    public string FormatId => "md";

    private static readonly MarkdownPipeline DefaultPipeline = new MarkdownPipelineBuilder().Build();

    public PolyDocument Read(Stream input)
    {
        ArgumentNullException.ThrowIfNull(input);
        using var reader = new StreamReader(input, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, leaveOpen: true);
        return FromMarkdown(reader.ReadToEnd());
    }

    public static PolyDocument FromMarkdown(string source)
    {
        ArgumentNullException.ThrowIfNull(source);
        var document = new PolyDocument();
        var section = new Section();
        document.Sections.Add(section);

        var ast = Markdig.Markdown.Parse(source, DefaultPipeline);
        ProcessContainer(ast, section, marker: null);
        return document;
    }

    private static void ProcessContainer(ContainerBlock container, Section section, ListMarker? marker)
    {
        foreach (var block in container)
        {
            switch (block)
            {
                case HeadingBlock heading:
                    {
                        var paragraph = new Paragraph();
                        paragraph.Style.Outline = (OutlineLevel)Math.Clamp(heading.Level, 1, 6);
                        paragraph.Style.ListMarker = CloneMarker(marker);
                        AppendInline(paragraph, heading.Inline);
                        section.Blocks.Add(paragraph);
                        break;
                    }
                case ParagraphBlock para:
                    {
                        var paragraph = new Paragraph();
                        paragraph.Style.ListMarker = CloneMarker(marker);
                        AppendInline(paragraph, para.Inline);
                        section.Blocks.Add(paragraph);
                        break;
                    }
                case ListBlock list:
                    {
                        int counter = 0;
                        foreach (var item in list)
                        {
                            counter++;
                            if (item is not ListItemBlock listItem)
                            {
                                continue;
                            }
                            var listMarker = list.IsOrdered
                                ? new ListMarker
                                {
                                    Kind = ListKind.OrderedDecimal,
                                    OrderedNumber = listItem.Order > 0 ? listItem.Order : counter,
                                }
                                : new ListMarker { Kind = ListKind.Bullet };
                            ProcessContainer(listItem, section, listMarker);
                        }
                        break;
                    }
                case QuoteBlock quote:
                    {
                        // Phase A 모델 범위: 인용은 일반 단락으로 격하.
                        ProcessContainer(quote, section, marker);
                        break;
                    }
                case FencedCodeBlock fenced:
                    {
                        section.Blocks.Add(BuildCodeParagraph(fenced.Lines, marker));
                        break;
                    }
                case CodeBlock indented:
                    {
                        section.Blocks.Add(BuildCodeParagraph(indented.Lines, marker));
                        break;
                    }
                // ThematicBreakBlock / HtmlBlock / 표·링크 등은 향후 Core 모델 확장 후 매핑.
            }
        }
    }

    private static Paragraph BuildCodeParagraph(Markdig.Helpers.StringLineGroup lines, ListMarker? marker)
    {
        var paragraph = new Paragraph();
        paragraph.Style.ListMarker = CloneMarker(marker);
        var slices = lines.Lines.Take(lines.Count).Select(l => l.ToString());
        var code = string.Join('\n', slices);
        var style = new RunStyle { FontFamily = "Consolas, D2Coding, monospace" };
        paragraph.AddText(code, style);
        return paragraph;
    }

    private static void AppendInline(Paragraph paragraph, ContainerInline? inline)
    {
        if (inline is null)
        {
            paragraph.AddText(string.Empty);
            return;
        }

        WalkInline(paragraph, inline, new RunStyle());
        if (paragraph.Runs.Count == 0)
        {
            paragraph.AddText(string.Empty);
        }
    }

    private static void WalkInline(Paragraph paragraph, Inline node, RunStyle style)
    {
        switch (node)
        {
            case LiteralInline lit:
                paragraph.AddText(lit.Content.ToString(), Clone(style));
                break;

            case EmphasisInline em:
                {
                    var emStyle = Clone(style);
                    if (em.DelimiterCount >= 2)
                    {
                        emStyle.Bold = true;
                    }
                    else
                    {
                        emStyle.Italic = true;
                    }
                    foreach (var child in em)
                    {
                        WalkInline(paragraph, child, emStyle);
                    }
                    break;
                }

            case CodeInline code:
                {
                    var codeStyle = Clone(style);
                    codeStyle.FontFamily = "Consolas, D2Coding, monospace";
                    paragraph.AddText(code.Content, codeStyle);
                    break;
                }

            case LineBreakInline:
                paragraph.AddText(" ", Clone(style));
                break;

            case LinkInline link:
                {
                    var linkStyle = Clone(style);
                    linkStyle.Underline = true;
                    foreach (var child in link)
                    {
                        WalkInline(paragraph, child, linkStyle);
                    }
                    break;
                }

            case ContainerInline container:
                foreach (var child in container)
                {
                    WalkInline(paragraph, child, style);
                }
                break;
        }
    }

    private static RunStyle Clone(RunStyle s) => new()
    {
        FontFamily = s.FontFamily,
        FontSizePt = s.FontSizePt,
        Bold = s.Bold,
        Italic = s.Italic,
        Underline = s.Underline,
        Strikethrough = s.Strikethrough,
        Overline = s.Overline,
        Superscript = s.Superscript,
        Subscript = s.Subscript,
        Foreground = s.Foreground,
        Background = s.Background,
        WidthPercent = s.WidthPercent,
        LetterSpacingPx = s.LetterSpacingPx,
    };

    private static ListMarker? CloneMarker(ListMarker? marker)
        => marker is null ? null : new ListMarker
        {
            Kind = marker.Kind,
            Level = marker.Level,
            OrderedNumber = marker.OrderedNumber,
        };
}

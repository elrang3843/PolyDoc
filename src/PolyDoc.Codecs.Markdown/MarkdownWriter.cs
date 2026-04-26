using System.Text;
using PolyDoc.Core;

namespace PolyDoc.Codecs.Markdown;

/// <summary>
/// PolyDocument 를 CommonMark Markdown 으로 직렬화한다.
/// 매핑 규칙:
///   - OutlineLevel.H1~H6  → "# " ~ "###### " 헤더
///   - ListMarker.Bullet   → "- "
///   - ListMarker.Ordered* → "1. " (Markdown 은 자동 번호 부여)
///   - Run.Bold            → **...**
///   - Run.Italic          → *...*
/// </summary>
public sealed class MarkdownWriter : IDocumentWriter
{
    public string FormatId => "md";

    public Encoding Encoding { get; init; } = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

    public void Write(PolyDocument document, Stream output)
    {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(output);

        using var writer = new StreamWriter(output, Encoding, leaveOpen: true) { NewLine = "\n" };
        writer.Write(ToMarkdown(document));
    }

    public static string ToMarkdown(PolyDocument document)
    {
        ArgumentNullException.ThrowIfNull(document);
        var sb = new StringBuilder();
        bool first = true;

        foreach (var paragraph in document.EnumerateParagraphs())
        {
            if (!first)
            {
                sb.Append('\n');   // 블록 간 빈 줄
            }
            sb.Append(RenderParagraph(paragraph));
            sb.Append('\n');
            first = false;
        }

        return sb.ToString();
    }

    private static string RenderParagraph(Paragraph paragraph)
    {
        var prefix = string.Empty;
        if (paragraph.Style.Outline is var lvl and > OutlineLevel.Body)
        {
            prefix = new string('#', (int)lvl) + " ";
        }
        else if (paragraph.Style.ListMarker is { } list)
        {
            prefix = list.Kind switch
            {
                ListKind.Bullet => "- ",
                _ => $"{list.OrderedNumber ?? 1}. ",
            };
        }

        return prefix + RenderRuns(paragraph);
    }

    private static string RenderRuns(Paragraph paragraph)
    {
        var sb = new StringBuilder();
        foreach (var run in paragraph.Runs)
        {
            var text = run.Text;
            if (run.Style.Bold && run.Style.Italic)
            {
                sb.Append("***").Append(text).Append("***");
            }
            else if (run.Style.Bold)
            {
                sb.Append("**").Append(text).Append("**");
            }
            else if (run.Style.Italic)
            {
                sb.Append('*').Append(text).Append('*');
            }
            else
            {
                sb.Append(text);
            }
        }
        return sb.ToString();
    }
}

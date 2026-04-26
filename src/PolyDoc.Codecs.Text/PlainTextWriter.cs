using System.Text;
using PolyDoc.Core;

namespace PolyDoc.Codecs.Text;

/// <summary>PolyDocument 의 모든 문단을 LF 로 이어 단일 평문으로 직렬화한다.</summary>
public sealed class PlainTextWriter : IDocumentWriter
{
    public string FormatId => "txt";

    /// <summary>출력 인코딩. 기본 UTF-8 (BOM 없음).</summary>
    public Encoding Encoding { get; init; } = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

    /// <summary>줄 구분자. 기본 LF.</summary>
    public string NewLine { get; init; } = "\n";

    public void Write(PolyDocument document, Stream output)
    {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(output);

        using var writer = new StreamWriter(output, Encoding, leaveOpen: true)
        {
            NewLine = NewLine,
        };
        WriteTo(document, writer);
    }

    public static string ToText(PolyDocument document, string newLine = "\n")
    {
        ArgumentNullException.ThrowIfNull(document);
        var sb = new StringBuilder();
        bool first = true;
        foreach (var paragraph in document.EnumerateParagraphs())
        {
            if (!first)
            {
                sb.Append(newLine);
            }
            sb.Append(paragraph.GetPlainText());
            first = false;
        }
        return sb.ToString();
    }

    private static void WriteTo(PolyDocument document, StreamWriter writer)
    {
        bool first = true;
        foreach (var paragraph in document.EnumerateParagraphs())
        {
            if (!first)
            {
                writer.Write(writer.NewLine);
            }
            writer.Write(paragraph.GetPlainText());
            first = false;
        }
    }
}

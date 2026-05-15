using System;
using System.Text;
using PolyDonky.Core;

namespace PolyDonky.Convert.Doc;

/// <summary>
/// IWPF → DOC (97-2003) 변환기. 현재 단계: 텍스트만.
/// RTF (Rich Text Format) 형식으로 생성하며, Word에서 .doc으로 열 수 있다.
/// </summary>
public class DocWriter
{
    public void Write(PolyDonkyument doc, Stream output)
    {
        var text = ExtractAllText(doc);
        var rtf = GenerateRtf(text);

        using (var writer = new StreamWriter(output, Encoding.Default, leaveOpen: true))
        {
            writer.Write(rtf);
        }
    }

    private string ExtractAllText(PolyDonkyument doc)
    {
        var sb = new StringBuilder();
        foreach (var section in doc.Sections)
        {
            foreach (var block in section.Blocks)
            {
                if (block is Paragraph para)
                {
                    foreach (var run in para.Runs)
                    {
                        if (!string.IsNullOrEmpty(run.Text))
                        {
                            // Escape RTF special characters
                            sb.Append(EscapeRtf(run.Text));
                        }
                    }
                    sb.Append(@"\par");  // Paragraph end
                    sb.AppendLine();
                }
            }
        }
        return sb.ToString();
    }

    private string EscapeRtf(string text)
    {
        if (string.IsNullOrEmpty(text))
            return text;

        var sb = new StringBuilder();
        foreach (char c in text)
        {
            switch (c)
            {
                case '\\':
                case '{':
                case '}':
                    sb.Append('\\').Append(c);
                    break;
                case '\n':
                    sb.Append(@"\par ");
                    break;
                case '\r':
                    // Skip carriage returns
                    break;
                case '\t':
                    sb.Append(@"\tab ");
                    break;
                default:
                    if (c < 128)
                        sb.Append(c);
                    else
                        // Encode non-ASCII as Unicode escape
                        sb.Append($@"\u{(short)c}?");
                    break;
            }
        }
        return sb.ToString();
    }

    private string GenerateRtf(string content)
    {
        var sb = new StringBuilder();

        // RTF Header
        sb.AppendLine(@"{\rtf1\ansi\ansicpg1252\deff0");
        sb.AppendLine(@"{\fonttbl{\f0\fnil\fcharset0 Arial;}}");
        sb.AppendLine(@"{\colortbl;\red0\green0\blue0;}");
        sb.AppendLine(@"\viewkind4\uc1\pard\f0\fs20");

        // Content
        sb.Append(content);

        // If no trailing \par, add one
        if (!content.EndsWith(@"\par"))
            sb.Append(@"\par");

        // RTF Footer
        sb.AppendLine();
        sb.Append("}");

        return sb.ToString();
    }
}

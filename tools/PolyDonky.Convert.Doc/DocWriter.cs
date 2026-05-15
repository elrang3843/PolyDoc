using System;
using System.Collections.Generic;
using System.Text;
using PolyDonky.Core;

namespace PolyDonky.Convert.Doc;

/// <summary>
/// IWPF → DOC (97-2003) 변환기. 텍스트 + 기본 포매팅 지원.
/// RTF (Rich Text Format) 형식으로 생성하며, Word에서 .doc으로 열 수 있다.
/// </summary>
public class DocWriter
{
    private List<RtfColor> _colorTable = new();

    public void Write(PolyDonkyument doc, Stream output)
    {
        _colorTable.Clear();
        _colorTable.Add(new RtfColor(0, 0, 0));  // Color 0: default black

        var paragraphs = ExtractParagraphs(doc);
        var rtf = GenerateRtf(paragraphs);

        using (var writer = new StreamWriter(output, Encoding.Default, leaveOpen: true))
        {
            writer.Write(rtf);
        }
    }

    private List<ParagraphInfo> ExtractParagraphs(PolyDonkyument doc)
    {
        var paragraphs = new List<ParagraphInfo>();

        foreach (var section in doc.Sections)
        {
            foreach (var block in section.Blocks)
            {
                if (block is Paragraph para)
                {
                    var info = new ParagraphInfo
                    {
                        Style = para.Style ?? new ParagraphStyle(),
                        Runs = new List<RunInfo>()
                    };

                    foreach (var run in para.Runs)
                    {
                        if (!string.IsNullOrEmpty(run.Text))
                        {
                            var runInfo = new RunInfo
                            {
                                Text = run.Text,
                                Style = run.Style ?? new RunStyle()
                            };

                            // 색상 테이블에 색 추가
                            if (runInfo.Style.Foreground.HasValue)
                            {
                                var color = new RtfColor(
                                    runInfo.Style.Foreground.Value.R,
                                    runInfo.Style.Foreground.Value.G,
                                    runInfo.Style.Foreground.Value.B
                                );
                                int idx = _colorTable.FindIndex(c => c.R == color.R && c.G == color.G && c.B == color.B);
                                if (idx < 0)
                                {
                                    _colorTable.Add(color);
                                    runInfo.ColorIndex = _colorTable.Count - 1;
                                }
                                else
                                {
                                    runInfo.ColorIndex = idx;
                                }
                            }

                            info.Runs.Add(runInfo);
                        }
                    }

                    if (info.Runs.Count > 0)
                        paragraphs.Add(info);
                }
            }
        }

        return paragraphs;
    }

    private struct RtfColor
    {
        public byte R { get; }
        public byte G { get; }
        public byte B { get; }

        public RtfColor(byte r, byte g, byte b)
        {
            R = r;
            G = g;
            B = b;
        }
    }

    private string GenerateRtf(List<ParagraphInfo> paragraphs)
    {
        var sb = new StringBuilder();

        // RTF Header
        sb.AppendLine(@"{\rtf1\ansi\ansicpg1252\deff0");

        // Font table
        sb.AppendLine(@"{\fonttbl{\f0\fnil\fcharset0 Arial;{\*\fname Arial;}{\falt Helvetica;}}}");

        // Color table
        sb.Append(@"{\colortbl");
        foreach (var color in _colorTable)
        {
            sb.Append($@"\red{color.R}\green{color.G}\blue{color.B};");
        }
        sb.AppendLine("}");

        sb.AppendLine(@"\viewkind4\uc1");

        // Content
        foreach (var para in paragraphs)
        {
            // 문단 속성
            sb.Append(@"\pard");

            // 문단 정렬
            switch (para.Style.Alignment)
            {
                case Alignment.Center:
                    sb.Append(@"\qc");
                    break;
                case Alignment.Right:
                    sb.Append(@"\qr");
                    break;
                case Alignment.Justify:
                    sb.Append(@"\qj");
                    break;
                default: // Left
                    sb.Append(@"\ql");
                    break;
            }

            // 문단 간격
            if (para.Style.SpaceBeforePt > 0)
                sb.Append($@"\sb{(int)(para.Style.SpaceBeforePt * 20)}");
            if (para.Style.SpaceAfterPt > 0)
                sb.Append($@"\sa{(int)(para.Style.SpaceAfterPt * 20)}");

            // 줄 높이
            if (para.Style.LineHeightFactor > 0)
                sb.Append($@"\sl{(int)(para.Style.LineHeightFactor * 240)}\slmult1");

            // 기본 글자 속성
            sb.Append(@"\f0 ");

            // Run들 처리
            foreach (var run in para.Runs)
            {
                // 글자 색
                if (run.ColorIndex > 0)
                    sb.Append($@"\cf{run.ColorIndex}");
                else
                    sb.Append(@"\cf0");

                // 글자 크기 (RTF는 반포인트 단위, 즉 포인트 * 2)
                double fontSize = run.Style.FontSizePt > 0 ? run.Style.FontSizePt : 11;
                sb.Append($@"\fs{(int)(fontSize * 2)}");

                // 글자 스타일
                if (run.Style.Bold)
                    sb.Append(@"\b");
                if (run.Style.Italic)
                    sb.Append(@"\i");
                if (run.Style.Underline)
                    sb.Append(@"\ul");
                if (run.Style.Strikethrough)
                    sb.Append(@"\strike");

                // 텍스트
                sb.Append(" ");
                sb.Append(EscapeRtf(run.Text));

                // 스타일 종료
                if (run.Style.Bold)
                    sb.Append(@"\b0");
                if (run.Style.Italic)
                    sb.Append(@"\i0");
                if (run.Style.Underline)
                    sb.Append(@"\ul0");
                if (run.Style.Strikethrough)
                    sb.Append(@"\strike0");

                sb.Append(" ");
            }

            sb.AppendLine(@"\par");
        }

        // RTF Footer
        sb.Append("}");

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

    private class ParagraphInfo
    {
        public ParagraphStyle Style { get; set; } = new();
        public List<RunInfo> Runs { get; set; } = new();
    }

    private class RunInfo
    {
        public string Text { get; set; } = string.Empty;
        public RunStyle Style { get; set; } = new();
        public int ColorIndex { get; set; } = 0;
    }
}

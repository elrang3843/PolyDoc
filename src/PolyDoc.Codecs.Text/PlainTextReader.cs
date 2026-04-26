using System.Text;
using PolyDoc.Core;

namespace PolyDoc.Codecs.Text;

/// <summary>
/// 평문 텍스트를 PolyDocument 로 변환한다. 빈 줄을 만나면 새 문단을 시작하지 않고
/// 한 줄 = 한 문단의 단순 매핑을 사용한다 (TXT 의 의미론은 그게 전부이므로).
/// </summary>
public sealed class PlainTextReader : IDocumentReader
{
    public string FormatId => "txt";

    /// <summary>BOM 이 없을 때 사용할 인코딩. 기본 UTF-8.</summary>
    public Encoding DefaultEncoding { get; init; } = Encoding.UTF8;

    public PolyDocument Read(Stream input)
    {
        ArgumentNullException.ThrowIfNull(input);

        // BOM 자동 감지: detectEncodingFromByteOrderMarks=true
        using var reader = new StreamReader(input, DefaultEncoding, detectEncodingFromByteOrderMarks: true, leaveOpen: true);
        var text = reader.ReadToEnd();
        return FromText(text);
    }

    public static PolyDocument FromText(string text)
    {
        ArgumentNullException.ThrowIfNull(text);
        var document = new PolyDocument();
        var section = new Section();
        document.Sections.Add(section);

        // CRLF / CR / LF 모두 한 줄 구분으로 처리.
        var lines = text.Split('\n');
        for (int i = 0; i < lines.Length; i++)
        {
            var line = lines[i];
            if (line.Length > 0 && line[^1] == '\r')
            {
                line = line[..^1];
            }
            // 마지막 빈 줄은 trailing newline 의 결과이므로 추가하지 않는다.
            if (i == lines.Length - 1 && line.Length == 0)
            {
                break;
            }
            section.Blocks.Add(Paragraph.Of(line));
        }

        return document;
    }
}

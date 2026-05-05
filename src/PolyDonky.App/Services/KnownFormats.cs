using System;
using System.IO;
using PolyDonky.Codecs.Docx;
using PolyDonky.Codecs.Html;
using PolyDonky.Codecs.Hwpx;
using PolyDonky.Codecs.Markdown;
using PolyDonky.Codecs.Text;
using PolyDonky.Core;
using PolyDonky.Iwpf;
using PdXmlReader = PolyDonky.Codecs.Xml.XmlReader;
using PdXmlWriter = PolyDonky.Codecs.Xml.XmlWriter;

namespace PolyDonky.App.Services;

/// <summary>
/// 파일 확장자에서 적절한 codec 을 선택해 read/write 를 수행한다.
/// IWPF / DOCX / MD / TXT 는 메인 앱이 직접 처리하고, 다른 포맷은 추후 외부 CLI 컨버터를 통해 처리한다.
/// </summary>
/// <remarks>
/// 클래스 이름은 <c>DocumentFormat.OpenXml</c> 패키지 네임스페이스(<c>DocumentFormat</c>)와의
/// 단순 이름 충돌을 피하기 위해 <c>KnownFormats</c> 로 둔다. 컴파일러는 단순 이름 <c>DocumentFormat</c>
/// 을 namespace 로 우선 해석할 수 있어, 같은 이름의 정적 클래스를 두면 CS0234 가 발생한다.
/// </remarks>
public static class KnownFormats
{
    public static IDocumentReader? PickReader(string path)
    {
        return GetExtensionId(path) switch
        {
            "iwpf" => new IwpfReader(),
            "md" or "markdown" => new MarkdownReader(),
            "txt" => new PlainTextReader(),
            "docx" => new DocxReader(),
            "hwpx" => new HwpxReader(),
            "html" or "htm" => new HtmlReader(),
            "xml" or "xhtml" => new PdXmlReader(),
            _ => null,
        };
    }

    public static IDocumentWriter? PickWriter(string path)
    {
        return GetExtensionId(path) switch
        {
            "iwpf" => new IwpfWriter(),
            "md" or "markdown" => new MarkdownWriter(),
            "txt" => new PlainTextWriter(),
            "docx" => new DocxWriter(),
            "hwpx" => new HwpxWriter(),
            "html" or "htm" => new HtmlWriter(),
            "xml" or "xhtml" => new PdXmlWriter(),
            _ => null,
        };
    }

    public static bool IsSupportedNatively(string path)
        => PickReader(path) is not null;

    /// <summary>"외부 컨버터 위탁" 대상 — 레거시 바이너리(HWP/DOC) 만 외부 CLI 위탁 유지.</summary>
    public static bool RequiresExternalConverter(string path)
    {
        return GetExtensionId(path) switch
        {
            "hwp" or "doc" => true,
            _ => false,
        };
    }

    public const string OpenFilter =
        // 직접 처리
        "PolyDonky 직접 지원 (IWPF·DOCX·HWPX·HTML·XML·MD·TXT)|*.iwpf;*.docx;*.hwpx;*.html;*.htm;*.xml;*.xhtml;*.md;*.markdown;*.txt|" +
        "PolyDonky 문서 (*.iwpf)|*.iwpf|" +
        "Word DOCX (*.docx)|*.docx|" +
        "한글 HWPX (*.hwpx)|*.hwpx|" +
        "HTML (*.html;*.htm)|*.html;*.htm|" +
        "XML / XHTML (*.xml;*.xhtml)|*.xml;*.xhtml|" +
        "Markdown (*.md;*.markdown)|*.md;*.markdown|" +
        "텍스트 (*.txt)|*.txt|" +
        // 외부 컨버터 위탁
        "한글 HWP (*.hwp) — 외부 컨버터 필요|*.hwp|" +
        "Word 레거시 (*.doc) — 외부 컨버터 필요|*.doc|" +
        "모든 파일 (*.*)|*.*";

    public const string SaveFilter =
        "PolyDonky 문서 (*.iwpf)|*.iwpf|" +
        "Word DOCX (*.docx)|*.docx|" +
        "한글 HWPX (*.hwpx)|*.hwpx|" +
        "HTML (*.html)|*.html|" +
        "XML / XHTML (*.xml)|*.xml|" +
        "Markdown (*.md)|*.md|" +
        "텍스트 (*.txt)|*.txt";

    private static string GetExtensionId(string path)
        => Path.GetExtension(path).TrimStart('.').ToLowerInvariant();
}

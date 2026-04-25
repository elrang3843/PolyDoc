using System;
using System.IO;
using PolyDoc.Codecs.Markdown;
using PolyDoc.Codecs.Text;
using PolyDoc.Core;
using PolyDoc.Iwpf;

namespace PolyDoc.App.Services;

/// <summary>
/// 파일 확장자에서 적절한 codec 을 선택해 read/write 를 수행한다.
/// IWPF / MD / TXT 만 메인 앱이 직접 처리하고, 다른 포맷은 추후 외부 CLI 컨버터를 통해 처리한다.
/// </summary>
public static class DocumentFormat
{
    public static IDocumentReader? PickReader(string path)
    {
        return GetExtensionId(path) switch
        {
            "iwpf" => new IwpfReader(),
            "md" or "markdown" => new MarkdownReader(),
            "txt" => new PlainTextReader(),
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
            _ => null,
        };
    }

    public static bool IsSupportedNatively(string path)
        => PickReader(path) is not null;

    /// <summary>"외부 컨버터 위탁" 대상 — Phase D 에서 외부 CLI 호출로 연결.</summary>
    public static bool RequiresExternalConverter(string path)
    {
        return GetExtensionId(path) switch
        {
            "hwp" or "hwpx" or "doc" or "docx" or "html" or "htm" => true,
            _ => false,
        };
    }

    public const string OpenFilter =
        "모든 지원 문서|*.iwpf;*.md;*.markdown;*.txt;*.hwp;*.hwpx;*.doc;*.docx;*.html;*.htm|" +
        "PolyDoc 문서 (*.iwpf)|*.iwpf|" +
        "Markdown (*.md;*.markdown)|*.md;*.markdown|" +
        "텍스트 (*.txt)|*.txt|" +
        "한글 (*.hwp;*.hwpx)|*.hwp;*.hwpx|" +
        "Word (*.doc;*.docx)|*.doc;*.docx|" +
        "HTML (*.html;*.htm)|*.html;*.htm|" +
        "모든 파일 (*.*)|*.*";

    public const string SaveFilter =
        "PolyDoc 문서 (*.iwpf)|*.iwpf|" +
        "Markdown (*.md)|*.md|" +
        "텍스트 (*.txt)|*.txt|" +
        "HWPX (*.hwpx)|*.hwpx|" +
        "DOCX (*.docx)|*.docx|" +
        "HTML (*.html)|*.html";

    private static string GetExtensionId(string path)
        => Path.GetExtension(path).TrimStart('.').ToLowerInvariant();
}

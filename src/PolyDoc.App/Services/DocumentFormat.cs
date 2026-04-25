using System;
using System.IO;
using PolyDoc.Codecs.Docx;
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
            "docx" => new DocxReader(),
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
            _ => null,
        };
    }

    public static bool IsSupportedNatively(string path)
        => PickReader(path) is not null;

    /// <summary>"외부 컨버터 위탁" 대상 — Phase D 에서 외부 CLI 호출로 연결.</summary>
    /// <remarks>
    /// Phase C 에서 DOCX 는 OpenXml SDK 로 직접 처리하므로 목록에서 제외.
    /// HWPX 는 Phase C 후반에 자체 구현으로 처리될 예정 (KS X 6101).
    /// 레거시 바이너리(HWP/DOC)와 HTML 만 외부 컨버터 위탁이 유지된다.
    /// </remarks>
    public static bool RequiresExternalConverter(string path)
    {
        return GetExtensionId(path) switch
        {
            "hwp" or "hwpx" or "doc" or "html" or "htm" => true,
            _ => false,
        };
    }

    public const string OpenFilter =
        // 직접 처리 (Phase A·C)
        "PolyDoc 직접 지원 (IWPF·DOCX·MD·TXT)|*.iwpf;*.docx;*.md;*.markdown;*.txt|" +
        "PolyDoc 문서 (*.iwpf)|*.iwpf|" +
        "Word DOCX (*.docx)|*.docx|" +
        "Markdown (*.md;*.markdown)|*.md;*.markdown|" +
        "텍스트 (*.txt)|*.txt|" +
        // 외부 컨버터 위탁 (Phase D 이후)
        "한글 HWPX (*.hwpx) — 외부 컨버터 필요|*.hwpx|" +
        "한글 HWP (*.hwp) — 외부 컨버터 필요|*.hwp|" +
        "Word 레거시 (*.doc) — 외부 컨버터 필요|*.doc|" +
        "HTML (*.html;*.htm) — 외부 컨버터 필요|*.html;*.htm|" +
        "모든 파일 (*.*)|*.*";

    public const string SaveFilter =
        // 직접 처리 (Phase A·C)
        "PolyDoc 문서 (*.iwpf)|*.iwpf|" +
        "Word DOCX (*.docx)|*.docx|" +
        "Markdown (*.md)|*.md|" +
        "텍스트 (*.txt)|*.txt|" +
        // 외부 컨버터 위탁 (Phase D 이후)
        "한글 HWPX (*.hwpx) — 외부 컨버터 필요|*.hwpx|" +
        "HTML (*.html) — 외부 컨버터 필요|*.html";

    private static string GetExtensionId(string path)
        => Path.GetExtension(path).TrimStart('.').ToLowerInvariant();
}

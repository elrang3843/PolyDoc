using System.Text;
using System.Text.RegularExpressions;
using AngleSharp;
using AngleSharp.Css;
using AngleSharp.Css.Dom;
using AngleSharp.Dom;
using PolyDonky.Codecs.Html;
using PolyDonky.Core;
using PolyDonky.Iwpf;

// PolyDonky.Convert.Html — HTML ↔ IWPF 변환 전용 콘솔 도구.
// CLAUDE.md §3 의 외부 변환 모듈 분리 원칙: 메인 앱은 IWPF/MD/TXT 만 직접 처리하고
// HTML 은 이 CLI 가 처리한다.
//
// 사용법:
//   PolyDonky.Convert.Html <input> <output> [--fragment] [--title <text>]
//   PolyDonky.Convert.Html --version | -v
//   PolyDonky.Convert.Html --help    | -h | /?
//
// 변환 쌍:
//   *.html|*.htm → *.iwpf : import (인코딩 자동 감지, 블록 한도 없음)
//   *.iwpf       → *.html|*.htm : export (기본 완전 HTML5 문서)
//
// 옵션 (export 시):
//   --fragment       <!DOCTYPE>·<html>·<head>·<body> 래퍼 없이 fragment 만 출력
//   --title <text>   <title> 요소 텍스트 (지정 안 하면 첫 H1 또는 기본값)
//
// 종료 코드:
//   0 성공
//   2 인자 오류
//   3 지원하지 않는 변환 쌍
//   4 입출력 실패
//   5 변환 실패

const int ExitOk            = 0;
const int ExitBadArgs       = 2;
const int ExitUnsupportedOp = 3;
const int ExitIoError       = 4;
const int ExitConvertError  = 5;

try { Console.OutputEncoding = Encoding.UTF8; } catch { /* redirected pipe 등 무시 */ }

// 레거시 코드페이지(cp949 / EUC-KR / Shift-JIS 등) 지원 등록 — 한국어 HTML 자동 감지에 필수.
try { Encoding.RegisterProvider(CodePagesEncodingProvider.Instance); }
catch { /* provider 부재 환경 무시 */ }

// Ctrl+C 시 임시파일 정리.
string? tempCleanupPath = null;
Console.CancelKeyPress += (_, e) =>
{
    if (tempCleanupPath is not null && File.Exists(tempCleanupPath))
    {
        try { File.Delete(tempCleanupPath); } catch { /* 무시 */ }
    }
    Console.Error.Flush();
    Console.Out.Flush();
    e.Cancel = false;
};

// ── 인자 파싱 ────────────────────────────────────────────────────────
var positional = new List<string>(2);
bool fragmentOut = false;
string? titleOut = null;

for (int i = 0; i < args.Length; i++)
{
    var a = args[i];
    switch (a)
    {
        case "--version" or "-v":
            Console.WriteLine("PolyDonky.Convert.Html 1.0");
            return ExitOk;
        case "--help" or "-h" or "/?":
            PrintHelp();
            return ExitOk;
        case "--fragment":
            fragmentOut = true;
            break;
        case "--title":
            if (i + 1 >= args.Length)
            {
                Console.Error.WriteLine("--title 다음에 텍스트가 와야 합니다.");
                return ExitBadArgs;
            }
            titleOut = args[++i];
            break;
        default:
            if (a.StartsWith("--title=", StringComparison.Ordinal))
                titleOut = a["--title=".Length..];
            else if (a.StartsWith("--", StringComparison.Ordinal))
            {
                Console.Error.WriteLine($"알 수 없는 옵션: {a}");
                return ExitBadArgs;
            }
            else if (positional.Count < 2)
                positional.Add(a);
            else
            {
                Console.Error.WriteLine("위치 인자는 2개여야 합니다 (input, output).");
                return ExitBadArgs;
            }
            break;
    }
}

if (positional.Count != 2)
{
    Console.Error.WriteLine("Usage: PolyDonky.Convert.Html <input> <output> [--fragment] [--title <text>]");
    Console.Error.WriteLine("  '--help' 로 자세한 도움말과 종료 코드 안내를 볼 수 있습니다.");
    return ExitBadArgs;
}

string inPath, outPath;
try
{
    inPath  = Path.GetFullPath(positional[0]);
    outPath = Path.GetFullPath(positional[1]);
}
catch (Exception ex)
{
    Console.Error.WriteLine($"경로 해석 실패: {ex.Message}");
    return ExitBadArgs;
}

string Ext(string p) => Path.GetExtension(p).TrimStart('.').ToLowerInvariant();
string inExt  = Ext(inPath);
string outExt = Ext(outPath);

// ── 검증 ────────────────────────────────────────────────────────────
if (string.Equals(inPath, outPath, StringComparison.OrdinalIgnoreCase))
{
    Console.Error.WriteLine($"입력과 출력 경로가 같습니다 — 자기 자신을 덮어쓸 수 없습니다: {inPath}");
    Console.Error.Flush();
    return ExitBadArgs;
}

bool isImport = (inExt is "html" or "htm") && outExt == "iwpf";
bool isExport = inExt == "iwpf" && (outExt is "html" or "htm");
if (!isImport && !isExport)
{
    Console.Error.WriteLine($"지원하지 않는 변환: .{inExt} → .{outExt}");
    Console.Error.WriteLine("  지원: .html|.htm → .iwpf, .iwpf → .html|.htm");
    Console.Error.Flush();
    return ExitUnsupportedOp;
}

if (fragmentOut && !isExport)
{
    Console.Error.WriteLine("--fragment 는 export 모드(*.iwpf → *.html) 에서만 사용 가능합니다.");
    Console.Error.Flush();
    return ExitBadArgs;
}
if (titleOut is not null && !isExport)
{
    Console.Error.WriteLine("--title 은 export 모드(*.iwpf → *.html) 에서만 사용 가능합니다.");
    Console.Error.Flush();
    return ExitBadArgs;
}

if (!File.Exists(inPath))
{
    Console.Error.WriteLine($"입력 파일이 없습니다: {inPath}");
    Console.Error.Flush();
    return ExitIoError;
}

if (new FileInfo(inPath).Length == 0)
{
    Console.Error.WriteLine($"입력 파일이 비어 있습니다(0 byte): {inPath}");
    Console.Error.Flush();
    return ExitIoError;
}

var outDir = Path.GetDirectoryName(outPath);
if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir))
{
    try { Directory.CreateDirectory(outDir); }
    catch (Exception ex)
    {
        Console.Error.WriteLine($"출력 디렉터리 생성 실패: {outDir}\n  → {ex.Message}");
        Console.Error.Flush();
        return ExitIoError;
    }
}

// ── HTML 사전 검사 (import 만) ────────────────────────────────────────
byte[]?  htmlBytes  = null;
Encoding htmlEnc    = Encoding.UTF8;
string   encLabel   = "UTF-8";

if (isImport)
{
    try
    {
        htmlBytes = File.ReadAllBytes(inPath);
    }
    catch (Exception ex)
    {
        Console.Error.WriteLine($"입력 파일 읽기 실패: {ex.Message}");
        Console.Error.Flush();
        return ExitIoError;
    }

    if (LooksBinary(htmlBytes))
    {
        Console.Error.WriteLine(
            $"HTML 텍스트 파일이 아닙니다 (NUL 바이트 다수 감지 — 바이너리 파일?): {inPath}");
        Console.Error.Flush();
        return ExitConvertError;
    }

    (htmlEnc, encLabel) = DetectEncoding(htmlBytes);
}

// ── 변환 본체 (원자적 쓰기) ──────────────────────────────────────────
var tempOut = outPath + ".tmp-" + Guid.NewGuid().ToString("N").Substring(0, 8);
tempCleanupPath = tempOut;

try
{
    if (isImport)
    {
        WriteProgress(0, $"HTML 읽는 중 (인코딩 {encLabel})");
        // 한도 0 = 무제한 — CLI 호출자가 명시적으로 변환을 시작했으므로 잘림 없이 처리.
        var htmlBaseDir = Path.GetDirectoryName(inPath) ?? ".";
        var text = htmlEnc.GetString(htmlBytes!);

        // CSS 전처리: 외부 stylesheet 인라인 → 캐스케이드 계산 후 style="" 인라이닝.
        WriteProgress(10, "CSS 인라이닝 중");
        text = InlineExternalStylesheets(text, htmlBaseDir);
        text = ComputeAndInlineCss(text);

        var doc = HtmlReader.FromHtml(text, maxBlocks: 0);

        // HTML 파일 기준 상대 경로 이미지를 디스크에서 읽어 data 로 내장.
        WriteProgress(30, "이미지 내장 중");
        EmbedLocalImages(doc, htmlBaseDir);

        WriteProgress(60, "IWPF 로 변환 중");
        using (var ofs = File.Create(tempOut))
            new IwpfWriter().Write(doc, ofs);
    }
    else // isExport
    {
        WriteProgress(0, "IWPF 읽는 중");
        PolyDonkyument doc;
        using (var fs = File.OpenRead(inPath))
            doc = new IwpfReader().Read(fs);

        WriteProgress(60, fragmentOut ? "HTML fragment 로 변환 중" : "HTML 로 변환 중");
        var writer = new HtmlWriter { FullDocument = !fragmentOut, DocumentTitle = titleOut };
        using (var ofs = File.Create(tempOut))
            writer.Write(doc, ofs);
    }

    if (File.Exists(outPath)) File.Delete(outPath);
    File.Move(tempOut, outPath);
    tempCleanupPath = null;
    WriteProgress(100, "완료");
    Console.WriteLine($"OK: {Path.GetFileName(inPath)} → {Path.GetFileName(outPath)}");
    Console.Out.Flush();
    return ExitOk;
}
catch (FileNotFoundException ex)
{
    Console.Error.WriteLine($"파일을 찾을 수 없습니다: {ex.FileName ?? inPath}");
    return ExitIoError;
}
catch (DirectoryNotFoundException ex)
{
    Console.Error.WriteLine($"디렉터리를 찾을 수 없습니다: {ex.Message}");
    return ExitIoError;
}
catch (UnauthorizedAccessException ex)
{
    Console.Error.WriteLine($"권한 거부: {ex.Message}");
    return ExitIoError;
}
catch (DecoderFallbackException ex)
{
    Console.Error.WriteLine(
        $"문자열 디코딩 실패 (감지된 인코딩 {encLabel} 으로는 일부 바이트를 변환할 수 없음): {inPath}\n  세부: {ex.Message}");
    return ExitConvertError;
}
catch (System.IO.InvalidDataException ex)
{
    Console.Error.WriteLine(
        $"파일 형식이 유효하지 않습니다: {inPath}\n  세부: {ex.Message}");
    return ExitConvertError;
}
catch (IOException ex)
{
    Console.Error.WriteLine($"I/O 실패: {ex.Message}");
    return ExitIoError;
}
catch (Exception ex)
{
    Console.Error.WriteLine($"변환 실패: {ex.GetType().Name}: {ex.Message}");
    return ExitConvertError;
}
finally
{
    try { if (File.Exists(tempOut)) File.Delete(tempOut); } catch { /* 무시 */ }
    Console.Error.Flush();
    Console.Out.Flush();
}

// ── 헬퍼 ────────────────────────────────────────────────────────────

static void WriteProgress(int percent, string message)
{
    Console.WriteLine($"PROGRESS:{percent}:{message}");
    Console.Out.Flush();
}

/// <summary>NUL 바이트가 5% 이상이면 바이너리로 간주.</summary>
static bool LooksBinary(byte[] bytes)
{
    int sample = Math.Min(bytes.Length, 1024);
    if (sample == 0) return false;
    int nul = 0;
    for (int i = 0; i < sample; i++)
        if (bytes[i] == 0) nul++;
    return nul * 20 > sample;
}

/// <summary>
/// 입력 바이트에서 HTML 인코딩 감지:
///   1. BOM (UTF-8 / UTF-16 LE/BE / UTF-32 LE/BE)
///   2. 첫 4KB 안에서 &lt;meta charset="X"&gt; 또는 http-equiv Content-Type charset
///   3. 위 모두 실패 → UTF-8 (HTML5 기본)
/// 감지된 .NET <see cref="Encoding"/> 와 사람이 읽는 라벨을 함께 반환.
/// </summary>
static (Encoding enc, string label) DetectEncoding(byte[] bytes)
{
    if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
        return (new UTF8Encoding(false), "UTF-8 (BOM)");
    if (bytes.Length >= 4)
    {
        if (bytes[0] == 0x00 && bytes[1] == 0x00 && bytes[2] == 0xFE && bytes[3] == 0xFF)
            return (new UTF32Encoding(true,  true), "UTF-32 BE (BOM)");
        if (bytes[0] == 0xFF && bytes[1] == 0xFE && bytes[2] == 0x00 && bytes[3] == 0x00)
            return (new UTF32Encoding(false, true), "UTF-32 LE (BOM)");
    }
    if (bytes.Length >= 2)
    {
        if (bytes[0] == 0xFE && bytes[1] == 0xFF) return (Encoding.BigEndianUnicode, "UTF-16 BE (BOM)");
        if (bytes[0] == 0xFF && bytes[1] == 0xFE) return (Encoding.Unicode,           "UTF-16 LE (BOM)");
    }

    int sniffLen = Math.Min(bytes.Length, 4096);
    // ASCII 로 1차 디코딩해 <meta> 만 찾는다 — 이 단계에선 비-ASCII 바이트는 모두 '?' 로 보여도 괜찮다.
    var head = Encoding.ASCII.GetString(bytes, 0, sniffLen);
    var match = Regex.Match(head,
        @"<meta\s+[^>]*charset\s*=\s*[""']?([\w\-]+)[""']?",
        RegexOptions.IgnoreCase);
    if (match.Success)
    {
        var name = match.Groups[1].Value;
        try
        {
            var enc = Encoding.GetEncoding(name);
            return (enc, $"{enc.WebName} (meta charset)");
        }
        catch
        {
            return (Encoding.UTF8, $"UTF-8 (meta charset '{name}' 미지원, 기본값으로 처리)");
        }
    }

    return (new UTF8Encoding(false), "UTF-8 (기본값)");
}

/// <summary>
/// 문서 전체를 순회하며 ResourcePath 만 있고 Data 가 없는 ImageBlock 을
/// 디스크에서 읽어 Data 에 내장한다.
/// http(s):// · data: · // 등 외부 URL 은 건너뛴다.
/// </summary>
static void EmbedLocalImages(PolyDonkyument doc, string baseDir)
{
    foreach (var section in doc.Sections)
        EmbedImagesInBlocks(section.Blocks, baseDir);
}

static void EmbedImagesInBlocks(IList<Block> blocks, string baseDir)
{
    foreach (var block in blocks)
    {
        switch (block)
        {
            case ImageBlock img when img.Data.Length == 0
                                  && !string.IsNullOrEmpty(img.ResourcePath)
                                  && !IsExternalUrl(img.ResourcePath):
                TryEmbedImageFromDisk(img, baseDir);
                break;

            case Table t:
                foreach (var row in t.Rows)
                    foreach (var cell in row.Cells)
                        EmbedImagesInBlocks(cell.Blocks, baseDir);
                break;
        }
    }
}

static bool IsExternalUrl(string path) =>
    path.StartsWith("//",       StringComparison.Ordinal) ||
    path.StartsWith("http://",  StringComparison.OrdinalIgnoreCase) ||
    path.StartsWith("https://", StringComparison.OrdinalIgnoreCase) ||
    path.StartsWith("ftp://",   StringComparison.OrdinalIgnoreCase) ||
    path.StartsWith("data:",    StringComparison.OrdinalIgnoreCase) ||
    path.StartsWith("mailto:",  StringComparison.OrdinalIgnoreCase);

static void TryEmbedImageFromDisk(ImageBlock img, string baseDir)
{
    try
    {
        // 쿼리스트링/프래그먼트 제거 후 경로 정규화.
        var resourcePath = img.ResourcePath!;
        var sep = resourcePath.IndexOfAny(['?', '#']);
        if (sep >= 0) resourcePath = resourcePath[..sep];

        // URL 인코딩 디코드 (%20 → 공백 등).
        resourcePath = Uri.UnescapeDataString(resourcePath);

        var fullPath = Path.GetFullPath(Path.Combine(baseDir, resourcePath));
        if (!File.Exists(fullPath)) return;

        img.Data         = File.ReadAllBytes(fullPath);
        img.ResourcePath = null;

        // MediaType 이 모호(application/octet-stream)거나 비어 있으면 확장자로 재추정.
        if (string.IsNullOrEmpty(img.MediaType) || img.MediaType == "application/octet-stream")
            img.MediaType = GuessMediaTypeByExt(Path.GetExtension(fullPath));
    }
    catch { /* 로드 실패 시 ResourcePath 유지 — 메인 앱이 placeholder 를 표시 */ }
}

static string GuessMediaTypeByExt(string ext) => ext.ToLowerInvariant() switch
{
    ".png"              => "image/png",
    ".jpg" or ".jpeg"   => "image/jpeg",
    ".gif"              => "image/gif",
    ".bmp"              => "image/bmp",
    ".tif" or ".tiff"   => "image/tiff",
    ".webp"             => "image/webp",
    ".svg"              => "image/svg+xml",
    _                   => "application/octet-stream",
};

/// <summary>
/// HTML 내 &lt;link rel="stylesheet" href="local.css"&gt; 를 파일 내용으로 인라인화.
/// 외부 URL(http://, //, data: 등) 은 그대로 둔다.
/// </summary>
static string InlineExternalStylesheets(string html, string baseDir)
{
    return Regex.Replace(html,
        @"<link\b[^>]*\brel\s*=\s*[""']stylesheet[""'][^>]*>",
        match =>
        {
            var hrefMatch = Regex.Match(match.Value,
                @"\bhref\s*=\s*[""']([^""']+)[""']",
                RegexOptions.IgnoreCase);
            if (!hrefMatch.Success) return "";
            var href = hrefMatch.Groups[1].Value;
            if (IsExternalUrl(href)) return match.Value;
            try
            {
                var cssPath = Path.GetFullPath(
                    Path.Combine(baseDir, Uri.UnescapeDataString(href)));
                if (!File.Exists(cssPath)) return "";
                var css = File.ReadAllText(cssPath, Encoding.UTF8);
                return $"<style>{css}</style>";
            }
            catch { return ""; }
        },
        RegexOptions.IgnoreCase | RegexOptions.Singleline);
}

/// <summary>
/// AngleSharp.Css 로 CSS 캐스케이드를 계산한 뒤 모든 요소의 style="" 속성으로 인라이닝.
/// 문서 CSS 규칙에 등장하는 속성만 처리(UA 기본값 제외).
/// 기존 인라인 style="" 은 CSS 규칙보다 우선(캐스케이드 규칙 준수).
/// </summary>
static string ComputeAndInlineCss(string html)
    => ComputeAndInlineCssAsync(html).GetAwaiter().GetResult();

static async Task<string> ComputeAndInlineCssAsync(string html)
{
    var config = Configuration.Default.WithCss();
    var context = BrowsingContext.New(config);
    var document = await context.OpenAsync(req => req.Content(html));

    // 문서 CSS 규칙에서 사용된 속성명 수집 — UA 기본값 제외의 핵심.
    var usedProps = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    foreach (var sheet in document.StyleSheets.OfType<ICssStyleSheet>())
        CollectUsedProperties(sheet.Rules, usedProps);

    if (usedProps.Count == 0) return html;

    var skipTags = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        { "script", "style", "head", "title", "meta", "link", "base", "noscript", "template" };

    foreach (var element in document.All)
    {
        if (skipTags.Contains(element.LocalName)) continue;

        ICssStyleDeclaration? computed;
        try { computed = element.ComputeCurrentStyle(); }
        catch { continue; }
        if (computed is null) continue;

        var applied = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var prop in usedProps)
        {
            var value = computed.GetPropertyValue(prop);
            if (!string.IsNullOrEmpty(value))
                applied[prop] = value;
        }

        if (applied.Count == 0) continue;

        // 기존 인라인 style="" 이 우선 — 캐스케이드 규칙(inline > author).
        var existing = element.GetAttribute("style");
        if (!string.IsNullOrWhiteSpace(existing))
        {
            foreach (var part in existing.Split(';',
                StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            {
                var colon = part.IndexOf(':');
                if (colon <= 0) continue;
                applied[part[..colon].Trim()] = part[(colon + 1)..].Trim();
            }
        }

        element.SetAttribute("style",
            string.Join(";", applied.Select(kv => $"{kv.Key}:{kv.Value}")));
    }

    return document.DocumentElement?.OuterHtml ?? html;
}

static void CollectUsedProperties(ICssRuleList rules, HashSet<string> result)
{
    foreach (var rule in rules)
    {
        if (rule is ICssStyleRule styleRule)
        {
            foreach (ICssProperty prop in styleRule.Style)
                result.Add(prop.Name);
        }
        else if (rule is ICssGroupingRule groupRule)
            CollectUsedProperties(groupRule.Rules, result);
    }
}

static void PrintHelp()
{
    Console.WriteLine("PolyDonky.Convert.Html — HTML ↔ IWPF 변환기");
    Console.WriteLine();
    Console.WriteLine("사용법:");
    Console.WriteLine("  PolyDonky.Convert.Html <input> <output> [--fragment] [--title <text>]");
    Console.WriteLine("  PolyDonky.Convert.Html --version | -v");
    Console.WriteLine("  PolyDonky.Convert.Html --help    | -h | /?");
    Console.WriteLine();
    Console.WriteLine("변환 쌍:");
    Console.WriteLine("  *.html|*.htm → *.iwpf : import (블록 한도 없음, 인코딩 자동 감지)");
    Console.WriteLine("  *.iwpf       → *.html|*.htm : export (기본 완전 HTML5 문서)");
    Console.WriteLine();
    Console.WriteLine("옵션 (export 시):");
    Console.WriteLine("  --fragment      <!DOCTYPE>·<html>·<head>·<body> 래퍼 없이 fragment 만 출력");
    Console.WriteLine("  --title <text>  <title> 요소 텍스트 지정 (생략 시 첫 H1 또는 기본값)");
    Console.WriteLine();
    Console.WriteLine("인코딩 감지 (import 시):");
    Console.WriteLine("  • BOM (UTF-8 / UTF-16 / UTF-32) 우선");
    Console.WriteLine("  • 첫 4KB 안의 <meta charset=\"X\"> 또는 http-equiv Content-Type charset");
    Console.WriteLine("  • cp949·EUC-KR·Shift-JIS 등 레거시 코드페이지 지원");
    Console.WriteLine("  • 위 모두 실패 → UTF-8 (HTML5 기본)");
    Console.WriteLine();
    Console.WriteLine("바이너리 거부:");
    Console.WriteLine("  앞 1KB 안에 NUL 바이트가 5% 이상이면 HTML 이 아닌 것으로 판단해 거부 (exit 5)");
    Console.WriteLine();
    Console.WriteLine("종료 코드:");
    Console.WriteLine("  0  성공");
    Console.WriteLine("  2  인자 오류");
    Console.WriteLine("  3  지원하지 않는 변환 쌍");
    Console.WriteLine("  4  입출력 실패");
    Console.WriteLine("  5  변환 실패 (HTML 파싱·내부 예외·바이너리 입력)");
    Console.WriteLine();
    Console.WriteLine("진행 표시:");
    Console.WriteLine("  표준 출력에 PROGRESS:<percent>:<message> 형식으로 emit (감지된 인코딩 포함)");
    Console.WriteLine();
    Console.WriteLine("출력 안전성:");
    Console.WriteLine("  임시파일에 쓴 뒤 원자적 rename — 도중 종료(Ctrl+C 포함) 시 반쪽 파일이 남지 않음");
}

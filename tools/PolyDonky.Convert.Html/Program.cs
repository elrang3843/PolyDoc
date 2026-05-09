using System.Globalization;
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
/// CSS 캐스케이드를 author 규칙 매칭으로만 계산해 각 요소의 style="" 로 인라이닝.
/// UA 초기값(list-style-type:disc, display:block 등) 은 명시적으로 매치되지 않은 한 절대 출력하지 않는다 —
/// 이전 구현(<see cref="IElement.ComputeCurrentStyle"/>) 은 매치 여부와 무관하게 모든 사용 속성에 대해
/// 초기값/상속값을 반환해 list-style-type:none 등이 잘못 전파되는 부작용이 있었다.
/// 기존 인라인 style="" 은 author 규칙보다 우선(캐스케이드 규칙 준수).
/// </summary>
static string ComputeAndInlineCss(string html)
    => ComputeAndInlineCssAsync(html).GetAwaiter().GetResult();

static async Task<string> ComputeAndInlineCssAsync(string html)
{
    var config = Configuration.Default.WithCss();
    var context = BrowsingContext.New(config);
    var document = await context.OpenAsync(req => req.Content(html));

    // 모든 author 스타일 규칙(매체 쿼리 등 grouping 안쪽 포함) 을 (selector, declarations) 쌍으로 평탄화.
    var rules = new List<(string Selector, ICssStyleDeclaration Decl)>();
    foreach (var sheet in document.StyleSheets.OfType<ICssStyleSheet>())
        CollectStyleRules(sheet.Rules, rules);

    if (rules.Count == 0) return html;

    // ::before / ::after pseudo-elements + counter() 해소 — 변환 시점에 실제 <span> 으로 굳혀
    // 이후 cascade 패스가 inline style 을 추가 적용한다.
    ResolvePseudoAndCounters(document, rules);

    // ::first-letter — 매치되는 요소의 첫 글자를 별도 <span> 으로 분리해 스타일 적용.
    ResolveFirstLetter(document, rules);

    var skipTags = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        { "script", "style", "head", "title", "meta", "link", "base", "noscript", "template" };

    // 콤마 분리된 selector 를 분기별로 펼쳐 specificity 를 가지는 평탄 리스트로 만든다.
    // 같은 declaration 에 대해 분기별로 (selector, specificity, sourceIndex) 를 가짐.
    var indexedRules = new List<(string Selector, int Specificity, int SourceIndex, ICssStyleDeclaration Decl)>();
    for (int idx = 0; idx < rules.Count; idx++)
    {
        var (selectorList, decl) = rules[idx];
        foreach (var branch in SplitTopLevelCommas(selectorList))
        {
            var trimmed = branch.Trim();
            if (trimmed.Length == 0) continue;
            indexedRules.Add((trimmed, ComputeSpecificity(trimmed), idx, decl));
        }
    }

    foreach (var element in document.All)
    {
        if (skipTags.Contains(element.LocalName)) continue;

        // 캐스케이드: 매치되는 모든 규칙을 (specificity asc, sourceIndex asc) 로 정렬해
        // 순서대로 덮어쓴다 — 마지막에 적용되는 규칙이 이김. CSS 사양의 specificity > source-order 와 일치.
        var matchedRules = new List<(int Specificity, int SourceIndex, ICssStyleDeclaration Decl)>();
        foreach (var r in indexedRules)
        {
            bool matched;
            try { matched = element.Matches(r.Selector); }
            catch { continue; }
            if (matched) matchedRules.Add((r.Specificity, r.SourceIndex, r.Decl));
        }
        matchedRules.Sort((a, b) =>
        {
            int s = a.Specificity.CompareTo(b.Specificity);
            return s != 0 ? s : a.SourceIndex.CompareTo(b.SourceIndex);
        });

        var applied = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var r in matchedRules)
        {
            foreach (ICssProperty p in r.Decl)
            {
                if (string.IsNullOrEmpty(p.Value)) continue;
                applied[p.Name] = p.Value;
            }
        }

        // 기존 인라인 style="" 이 author 규칙보다 우선.
        // 인라인 단축속성(margin / padding / border 등) 은 같은 영역의 longhand(margin-top 등) 를
        // 함께 무효화해야 한다.
        var existing = element.GetAttribute("style");
        if (!string.IsNullOrWhiteSpace(existing))
        {
            foreach (var part in existing.Split(';',
                StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            {
                var colon = part.IndexOf(':');
                if (colon <= 0) continue;
                var prop = part[..colon].Trim();
                var val  = part[(colon + 1)..].Trim();
                foreach (var sibling in ExpandShorthand(prop))
                    applied.Remove(sibling);
                applied[prop] = val;
            }
        }

        if (applied.Count == 0) continue;
        element.SetAttribute("style",
            string.Join(";", applied.Select(kv => $"{kv.Key}:{kv.Value}")));
    }

    return document.DocumentElement?.OuterHtml ?? html;
}

/// <summary>CSS selector 의 specificity 를 (a*100 + b*10 + c) 단일 정수로 근사. inline style 은 별도 처리이므로 제외.
/// a = id 개수, b = class/attr/pseudo-class 개수, c = type/pseudo-element 개수.</summary>
static int ComputeSpecificity(string selector)
{
    int a = 0, b = 0, c = 0;
    int i = 0;
    while (i < selector.Length)
    {
        char ch = selector[i];
        if (ch == '#') { a++; i++; i = SkipIdent(selector, i); }
        else if (ch == '.') { b++; i++; i = SkipIdent(selector, i); }
        else if (ch == '[')
        {
            b++;
            while (i < selector.Length && selector[i] != ']') i++;
            if (i < selector.Length) i++;
        }
        else if (ch == ':')
        {
            // ::pseudo-element 는 c 에 카운트, :pseudo-class 는 b.
            if (i + 1 < selector.Length && selector[i + 1] == ':')
            { c++; i += 2; i = SkipIdent(selector, i); }
            else
            {
                // :before/:after/:first-letter/:first-line 은 사양상 pseudo-element (c).
                int start = i + 1;
                int end = SkipIdent(selector, start);
                var name = selector.Substring(start, end - start).ToLowerInvariant();
                if (name is "before" or "after" or "first-letter" or "first-line") c++;
                else b++;
                i = end;
                // pseudo-class 함수 :nth-child(...) — 괄호 건너뜀.
                if (i < selector.Length && selector[i] == '(')
                {
                    int depth = 1; i++;
                    while (i < selector.Length && depth > 0)
                    {
                        if (selector[i] == '(') depth++;
                        else if (selector[i] == ')') depth--;
                        i++;
                    }
                }
            }
        }
        else if (char.IsLetter(ch))
        {
            // 타입 selector (a, div, td 등) — c 에 카운트.
            c++;
            i = SkipIdent(selector, i);
        }
        else if (ch == '*')
        {
            // universal selector — specificity 0.
            i++;
        }
        else
        {
            i++; // 결합자 ' ', '>', '+', '~' 등 또는 공백.
        }
    }
    return a * 10000 + b * 100 + c;
}

static int SkipIdent(string s, int i)
{
    while (i < s.Length && (char.IsLetterOrDigit(s[i]) || s[i] == '-' || s[i] == '_')) i++;
    return i;
}

/// <summary>"a, b, c" 형식의 selector 를 콤마 분기로 분리. 함수 인자 안의 콤마는 무시.</summary>
static IEnumerable<string> SplitTopLevelCommas(string selector)
{
    int depth = 0;
    int start = 0;
    for (int i = 0; i < selector.Length; i++)
    {
        char ch = selector[i];
        if (ch == '(' || ch == '[') depth++;
        else if (ch == ')' || ch == ']') depth--;
        else if (ch == ',' && depth == 0)
        {
            yield return selector.Substring(start, i - start);
            start = i + 1;
        }
    }
    if (start < selector.Length)
        yield return selector.Substring(start);
}

/// <summary>
/// CSS 단축 속성 이름이 주어지면 그것이 가리는 longhand 속성 이름들을 반환.
/// 캐스케이드 우선순위 처리에서 단축속성을 인라인에 둘 때 이전에 적용된 longhand 값을 청소하기 위해 사용.
/// 알려지지 않은 속성은 빈 배열.
/// </summary>
static IReadOnlyList<string> ExpandShorthand(string prop)
{
    string p = prop.ToLowerInvariant();
    return p switch
    {
        "margin"     => new[] { "margin-top", "margin-right", "margin-bottom", "margin-left" },
        "padding"    => new[] { "padding-top", "padding-right", "padding-bottom", "padding-left" },
        "border-top"    => new[] { "border-top-width",    "border-top-style",    "border-top-color" },
        "border-right"  => new[] { "border-right-width",  "border-right-style",  "border-right-color" },
        "border-bottom" => new[] { "border-bottom-width", "border-bottom-style", "border-bottom-color" },
        "border-left"   => new[] { "border-left-width",   "border-left-style",   "border-left-color" },
        "border" => new[]
        {
            "border-top",         "border-right",         "border-bottom",         "border-left",
            "border-top-width",   "border-right-width",   "border-bottom-width",   "border-left-width",
            "border-top-style",   "border-right-style",   "border-bottom-style",   "border-left-style",
            "border-top-color",   "border-right-color",   "border-bottom-color",   "border-left-color",
            "border-width", "border-style", "border-color",
        },
        "border-width" => new[] { "border-top-width", "border-right-width", "border-bottom-width", "border-left-width" },
        "border-style" => new[] { "border-top-style", "border-right-style", "border-bottom-style", "border-left-style" },
        "border-color" => new[] { "border-top-color", "border-right-color", "border-bottom-color", "border-left-color" },
        "background"   => new[] { "background-color", "background-image", "background-repeat",
                                  "background-position", "background-size", "background-attachment", "background-origin", "background-clip" },
        "font"         => new[] { "font-style", "font-variant", "font-weight", "font-stretch",
                                  "font-size", "line-height", "font-family" },
        "list-style"   => new[] { "list-style-type", "list-style-position", "list-style-image" },
        _ => Array.Empty<string>(),
    };
}

/// <summary>
/// CSS ::before / ::after 가상 요소 + counter() 함수를 변환 시점에 실제 &lt;span&gt; 노드로 굳혀
/// HtmlReader 가 일반 인라인 텍스트로 처리하도록 한다. 동적 효과 없는 정적 문서이므로
/// CSS 의 가상 요소는 한 번 굳히면 이후 라운드트립에 보존된다.
/// 지원 패턴:
///  - <c>content: 'literal'</c> / <c>content: "literal"</c>
///  - <c>content: counter(name)</c> (counter-increment / counter-reset 지원)
///  - 여러 토큰 연결: <c>content: counter(line) ' '</c>
///  - 쓸 수 없는 함수(<c>attr()</c>, <c>url()</c>, <c>open-quote</c> 등)는 무시.
/// 카운터 상태는 문서 깊이 우선 순회 동안 단일 글로벌 dictionary 로 추적 — CSS 의 scoped counter
/// 사양 일부를 단순화한 근사. 워드프로세서 도메인의 일반적인 줄 번호/리스트 번호엔 충분.
/// </summary>
static void ResolvePseudoAndCounters(IDocument document, List<(string Selector, ICssStyleDeclaration Decl)> rules)
{
    var pseudoBefore = new List<(string baseSel, ICssStyleDeclaration decl)>();
    var pseudoAfter  = new List<(string baseSel, ICssStyleDeclaration decl)>();
    foreach (var (sel, decl) in rules)
    {
        if (TryStripPseudo(sel, "before", out var b)) pseudoBefore.Add((b, decl));
        else if (TryStripPseudo(sel, "after", out var a)) pseudoAfter.Add((a, decl));
    }
    if (pseudoBefore.Count == 0 && pseudoAfter.Count == 0)
        return;

    var counters = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
    if (document.DocumentElement is { } root)
        WalkPseudo(root, rules, pseudoBefore, pseudoAfter, counters);
}

/// <summary>::first-letter 해소 — 매치되는 요소의 첫 텍스트 자식의 맨 앞 한 글자(또는 잇따른 구두점/따옴표 포함)
/// 를 별도 &lt;span&gt; 으로 잘라내고 인라인 style 을 적용한다. ::first-line 은 줄바꿈 위치에 의존하므로
/// 변환 시점에 정확히 알 수 없어 미지원 (필요 시 사용자가 직접 첫 줄을 분리 마크업).</summary>
static void ResolveFirstLetter(IDocument document, List<(string Selector, ICssStyleDeclaration Decl)> rules)
{
    var firstLetter = new List<(string baseSel, ICssStyleDeclaration decl)>();
    foreach (var (sel, decl) in rules)
    {
        if (TryStripPseudo(sel, "first-letter", out var b)) firstLetter.Add((b, decl));
    }
    if (firstLetter.Count == 0) return;

    foreach (var el in document.All.ToList())
    {
        if (el.GetAttribute("data-pd-pseudo") is not null) continue;

        Dictionary<string, string>? merged = null;
        foreach (var (sel, decl) in firstLetter)
        {
            bool m;
            try { m = el.Matches(sel); } catch { continue; }
            if (!m) continue;
            merged ??= new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (ICssProperty p in decl)
            {
                if (string.IsNullOrEmpty(p.Value)) continue;
                merged[p.Name] = p.Value;
            }
        }
        if (merged is null) continue;

        // 첫 텍스트 자식 노드를 찾는다 (공백 무시).
        IText? firstText = null;
        foreach (var n in el.ChildNodes)
        {
            if (n is IText t && !string.IsNullOrWhiteSpace(t.Data)) { firstText = t; break; }
            if (n is IElement) break;     // 첫 자식이 다른 요소면 우리 영역 아님
        }
        if (firstText is null) continue;

        var text = firstText.Data;
        int firstNonWs = 0;
        while (firstNonWs < text.Length && char.IsWhiteSpace(text[firstNonWs])) firstNonWs++;
        if (firstNonWs >= text.Length) continue;

        // 한 글자 (서로게이트 페어 처리 — 한글/이모지 안전).
        int letterEnd = firstNonWs + 1;
        if (char.IsHighSurrogate(text[firstNonWs]) && firstNonWs + 1 < text.Length)
            letterEnd = firstNonWs + 2;

        var prefix    = text.Substring(0, firstNonWs);
        var letter    = text.Substring(firstNonWs, letterEnd - firstNonWs);
        var remainder = text.Substring(letterEnd);

        var owner = el.Owner!;
        var span = owner.CreateElement("span");
        span.SetAttribute("data-pd-pseudo", "first-letter");
        span.TextContent = letter;
        var styleParts = new List<string>();
        foreach (var (n, v) in merged)
        {
            if (n.StartsWith("counter-", StringComparison.OrdinalIgnoreCase)) continue;
            styleParts.Add($"{n}:{v}");
        }
        if (styleParts.Count > 0) span.SetAttribute("style", string.Join(";", styleParts));

        // 텍스트 노드 분할: prefix(텍스트) → span → remainder(텍스트). InsertBefore + NextSibling 으로
        // 정확히 firstText 바로 뒤에 삽입한다 (DOM 표준).
        firstText.Data = prefix;
        var parent = firstText.Parent!;
        var anchor = firstText.NextSibling;
        if (anchor is not null) parent.InsertBefore(span, anchor);
        else                    parent.AppendChild(span);
        if (!string.IsNullOrEmpty(remainder))
        {
            var rem = owner.CreateTextNode(remainder);
            var anchor2 = span.NextSibling;
            if (anchor2 is not null) parent.InsertBefore(rem, anchor2);
            else                     parent.AppendChild(rem);
        }
    }
}

static bool IsPseudoSelector(string sel)
{
    if (string.IsNullOrEmpty(sel)) return false;
    return sel.IndexOf("::before", StringComparison.OrdinalIgnoreCase) >= 0
        || sel.IndexOf("::after",  StringComparison.OrdinalIgnoreCase) >= 0
        || sel.EndsWith(":before", StringComparison.OrdinalIgnoreCase)
        || sel.EndsWith(":after",  StringComparison.OrdinalIgnoreCase);
}

static bool TryStripPseudo(string sel, string pseudoName, out string baseSelector)
{
    baseSelector = sel;
    if (string.IsNullOrEmpty(sel)) return false;
    string[] suffixes = { "::" + pseudoName, ":" + pseudoName };
    foreach (var suffix in suffixes)
    {
        if (sel.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
        {
            baseSelector = sel.Substring(0, sel.Length - suffix.Length);
            if (string.IsNullOrWhiteSpace(baseSelector)) baseSelector = "*";
            return true;
        }
    }
    return false;
}

static void WalkPseudo(IElement el, List<(string, ICssStyleDeclaration)> rules,
    List<(string, ICssStyleDeclaration)> pseudoBefore,
    List<(string, ICssStyleDeclaration)> pseudoAfter,
    Dictionary<string, int> counters)
{
    // 우리가 직접 삽입한 합성 pseudo span 은 사용자 규칙 매치에서 제외 — 무한 재귀 방지.
    if (el.GetAttribute("data-pd-pseudo") is not null) return;

    // 일반 요소의 counter-reset / counter-increment (regular 규칙 매치).
    foreach (var (sel, decl) in rules)
    {
        if (IsPseudoSelector(sel)) continue; // pseudo 는 분리해 처리
        bool m;
        try { m = el.Matches(sel); } catch { continue; }
        if (!m) continue;
        var reset = decl.GetPropertyValue("counter-reset");
        if (!string.IsNullOrEmpty(reset)) ApplyCounterTokens(reset, counters, isReset: true);
        var incr = decl.GetPropertyValue("counter-increment");
        if (!string.IsNullOrEmpty(incr)) ApplyCounterTokens(incr, counters, isReset: false);
    }

    // ::before pseudo — 매치되는 모든 규칙을 캐스케이드 순서로 머지(같은 속성은 뒤가 이김), 단일 span 으로 삽입.
    var beforeMerged = MergePseudoRules(el, pseudoBefore);
    if (beforeMerged is { Count: > 0 })
        InjectPseudoSpanFromProps(el, beforeMerged, isBefore: true, counters);

    // 자식 재귀 (스냅샷 — 삽입 중 인덱스 흔들림 방지)
    var children = el.Children.ToList();
    foreach (var child in children)
        WalkPseudo(child, rules, pseudoBefore, pseudoAfter, counters);

    // ::after pseudo
    var afterMerged = MergePseudoRules(el, pseudoAfter);
    if (afterMerged is { Count: > 0 })
        InjectPseudoSpanFromProps(el, afterMerged, isBefore: false, counters);
}

static Dictionary<string,string>? MergePseudoRules(IElement el,
    List<(string baseSel, ICssStyleDeclaration decl)> pseudoRules)
{
    Dictionary<string,string>? props = null;
    foreach (var (sel, decl) in pseudoRules)
    {
        bool m;
        try { m = el.Matches(sel); } catch { continue; }
        if (!m) continue;
        props ??= new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase);
        foreach (ICssProperty p in decl)
        {
            if (string.IsNullOrEmpty(p.Value)) continue;
            props[p.Name] = p.Value;
        }
    }
    return props;
}

static void InjectPseudoSpanFromProps(IElement el, Dictionary<string,string> props, bool isBefore,
    Dictionary<string,int> counters)
{
    if (props.TryGetValue("counter-increment", out var inc))   ApplyCounterTokens(inc,  counters, isReset: false);
    if (props.TryGetValue("counter-reset",     out var reset)) ApplyCounterTokens(reset, counters, isReset: true);

    if (!props.TryGetValue("content", out var contentVal) || string.IsNullOrEmpty(contentVal)) return;
    var text = ResolvePseudoContent(contentVal, counters);
    if (text is null) return;

    var styleParts = new List<string>();
    foreach (var (n, v) in props)
    {
        if (n.Equals("content", StringComparison.OrdinalIgnoreCase)) continue;
        if (n.StartsWith("counter-", StringComparison.OrdinalIgnoreCase)) continue;
        styleParts.Add($"{n}:{v}");
    }

    var owner = el.Owner;
    if (owner is null) return;
    var span = owner.CreateElement("span");
    span.SetAttribute("data-pd-pseudo", isBefore ? "before" : "after");
    span.TextContent = text;
    if (styleParts.Count > 0)
        span.SetAttribute("style", string.Join(";", styleParts));

    if (isBefore)
        el.InsertBefore(span, el.FirstChild);
    else
        el.AppendChild(span);
}

static void ApplyCounterTokens(string value, Dictionary<string, int> counters, bool isReset)
{
    var tokens = value.Split(new[] { ' ', '\t', ',' }, StringSplitOptions.RemoveEmptyEntries);
    int i = 0;
    while (i < tokens.Length)
    {
        var name = tokens[i++];
        int n = isReset ? 0 : 1;
        if (i < tokens.Length && int.TryParse(tokens[i], NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed))
        { n = parsed; i++; }
        if (isReset) counters[name] = n;
        else
        {
            counters.TryGetValue(name, out var cur);
            counters[name] = cur + n;
        }
    }
}

static string? ResolvePseudoContent(string val, Dictionary<string, int> counters)
{
    var sb = new StringBuilder();
    int i = 0;
    while (i < val.Length)
    {
        while (i < val.Length && char.IsWhiteSpace(val[i])) i++;
        if (i >= val.Length) break;

        char c = val[i];
        if (c == '\'' || c == '"')
        {
            int end = val.IndexOf(c, i + 1);
            if (end < 0) return null;
            sb.Append(val.Substring(i + 1, end - i - 1));
            i = end + 1;
        }
        else if (StartsWithCi(val, i, "counter("))
        {
            int closing = val.IndexOf(')', i);
            if (closing < 0) return null;
            var inside = val.Substring(i + "counter(".Length, closing - i - "counter(".Length);
            var args = inside.Split(',');
            var name = args[0].Trim();
            counters.TryGetValue(name, out var cur);
            sb.Append(cur);
            i = closing + 1;
        }
        else if (StartsWithCi(val, i, "counters("))
        {
            // counters(name, sep) — 같은 이름의 모든 중첩 counter 를 sep 으로 join.
            // 본 구현은 단일 글로벌 counter 라 sep 없이 현재 값 하나만.
            int closing = val.IndexOf(')', i);
            if (closing < 0) return null;
            var inside = val.Substring(i + "counters(".Length, closing - i - "counters(".Length);
            var args = inside.Split(',');
            var name = args[0].Trim();
            counters.TryGetValue(name, out var cur);
            sb.Append(cur);
            i = closing + 1;
        }
        else if (StartsWithCi(val, i, "attr(") || StartsWithCi(val, i, "url("))
        {
            int closing = val.IndexOf(')', i);
            if (closing < 0) return null;
            i = closing + 1;
        }
        else
        {
            // 알 수 없는 토큰 — 한 글자 건너뜀.
            i++;
        }
    }
    return sb.ToString();
}

static bool StartsWithCi(string haystack, int idx, string needle)
    => idx + needle.Length <= haystack.Length &&
       string.Compare(haystack, idx, needle, 0, needle.Length, StringComparison.OrdinalIgnoreCase) == 0;

static void CollectStyleRules(ICssRuleList ruleList, List<(string, ICssStyleDeclaration)> result)
{
    foreach (var rule in ruleList)
    {
        if (rule is ICssStyleRule styleRule)
            result.Add((styleRule.SelectorText, styleRule.Style));
        else if (rule is ICssGroupingRule groupRule)
            CollectStyleRules(groupRule.Rules, result);
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

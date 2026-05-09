using System.Text;
using System.Text.RegularExpressions;
using PolyDonky.Core;
using PolyDonky.Iwpf;
using PdXmlReader = PolyDonky.Codecs.Xml.XmlReader;
using PdXmlWriter = PolyDonky.Codecs.Xml.XmlWriter;

// PolyDonky.Convert.Xml — XML / XHTML ↔ IWPF 변환 전용 콘솔 도구.
// CLAUDE.md §3 의 외부 변환 모듈 분리 원칙: 메인 앱은 IWPF/MD/TXT 만 직접 처리하고
// XML/XHTML 은 이 CLI 가 처리한다.
//
// 사용법:
//   PolyDonky.Convert.Xml <input> <output> [--fragment] [--title <text>]
//   PolyDonky.Convert.Xml --version | -v
//   PolyDonky.Convert.Xml --help    | -h | /?
//
// 변환 쌍:
//   *.xml|*.xhtml → *.iwpf : import (XHTML 자동 감지, 일반 XML 은 텍스트 추출)
//   *.iwpf        → *.xml|*.xhtml : export (XHTML5 polyglot markup)
//
// 옵션 (export 시):
//   --fragment       <?xml ?>·<!DOCTYPE>·<html> 래퍼 없이 fragment 만 출력
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

try { Console.OutputEncoding = Encoding.UTF8; } catch { /* 무시 */ }

// 레거시 코드페이지(cp949 / EUC-KR / Shift-JIS 등) 지원 등록 — XML 선언이 가리키는 인코딩 처리.
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
            Console.WriteLine("PolyDonky.Convert.Xml 1.0");
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
    Console.Error.WriteLine("Usage: PolyDonky.Convert.Xml <input> <output> [--fragment] [--title <text>]");
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

bool isImport = (inExt is "xml" or "xhtml") && outExt == "iwpf";
bool isExport = inExt == "iwpf" && (outExt is "xml" or "xhtml");
if (!isImport && !isExport)
{
    Console.Error.WriteLine($"지원하지 않는 변환: .{inExt} → .{outExt}");
    Console.Error.WriteLine("  지원: .xml|.xhtml → .iwpf, .iwpf → .xml|.xhtml");
    Console.Error.Flush();
    return ExitUnsupportedOp;
}

if (fragmentOut && !isExport)
{
    Console.Error.WriteLine("--fragment 는 export 모드(*.iwpf → *.xml) 에서만 사용 가능합니다.");
    Console.Error.Flush();
    return ExitBadArgs;
}
if (titleOut is not null && !isExport)
{
    Console.Error.WriteLine("--title 은 export 모드(*.iwpf → *.xml) 에서만 사용 가능합니다.");
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

// ── XML 사전 검사 (import 만) ────────────────────────────────────────
byte[]?  xmlBytes = null;
Encoding xmlEnc   = Encoding.UTF8;
string   encLabel = "UTF-8";

if (isImport)
{
    try
    {
        xmlBytes = File.ReadAllBytes(inPath);
    }
    catch (Exception ex)
    {
        Console.Error.WriteLine($"입력 파일 읽기 실패: {ex.Message}");
        Console.Error.Flush();
        return ExitIoError;
    }

    if (LooksBinary(xmlBytes))
    {
        Console.Error.WriteLine(
            $"XML 텍스트 파일이 아닙니다 (NUL 바이트 다수 감지 — 바이너리 파일?): {inPath}");
        Console.Error.Flush();
        return ExitConvertError;
    }

    (xmlEnc, encLabel) = DetectEncoding(xmlBytes);
}

// ── 변환 본체 (원자적 쓰기) ──────────────────────────────────────────
var tempOut = outPath + ".tmp-" + Guid.NewGuid().ToString("N").Substring(0, 8);
tempCleanupPath = tempOut;

try
{
    if (isImport)
    {
        WriteProgress(0, $"XML 읽는 중 (인코딩 {encLabel})");
        var text = xmlEnc.GetString(xmlBytes!);
        var doc  = PdXmlReader.FromXml(text);

        WriteProgress(60, "IWPF 로 변환 중");
        using (var ofs = File.Create(tempOut))
            new IwpfWriter().Write(doc, ofs);
    }
    else
    {
        WriteProgress(0, "IWPF 읽는 중");
        PolyDonkyument doc;
        using (var fs = File.OpenRead(inPath))
            doc = new IwpfReader().Read(fs);

        WriteProgress(60, fragmentOut ? "XHTML fragment 로 변환 중" : "XHTML 로 변환 중");
        var writer = new PdXmlWriter { FullDocument = !fragmentOut, DocumentTitle = titleOut };
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
catch (System.Xml.XmlException ex)
{
    // DTD 금지 정책 위반은 메시지에 'DTD' 가 포함됨 — 사용자에게 XXE 방어 정책 알림.
    if (ex.Message.Contains("DTD", StringComparison.OrdinalIgnoreCase))
    {
        Console.Error.WriteLine(
            $"DTD 가 포함된 XML 은 보안 정책상 거부됩니다 (XXE/외부 엔티티 공격 차단): {inPath}\n" +
            $"  줄 {ex.LineNumber}, 위치 {ex.LinePosition}: {ex.Message}\n" +
            "  → DOCTYPE 선언과 외부 엔티티 참조를 제거한 뒤 다시 시도하세요.");
    }
    else
    {
        Console.Error.WriteLine(
            $"XML 형식이 유효하지 않습니다: {inPath}\n" +
            $"  줄 {ex.LineNumber}, 위치 {ex.LinePosition}: {ex.Message}");
    }
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
/// 입력 바이트에서 XML 인코딩 감지:
///   1. BOM (UTF-8 / UTF-16 LE/BE / UTF-32 LE/BE)
///   2. 첫 256 바이트 안의 XML 선언 <c>&lt;?xml ... encoding="X" ... ?&gt;</c>
///   3. 위 모두 실패 → UTF-8 (XML 1.0 spec 기본값)
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

    // XML 선언은 파일 첫 줄에 와야 한다 — 256 바이트면 충분.
    int sniffLen = Math.Min(bytes.Length, 256);
    var head = Encoding.ASCII.GetString(bytes, 0, sniffLen);
    var match = Regex.Match(head,
        @"<\?xml\s+[^>]*encoding\s*=\s*[""']([\w\-]+)[""']",
        RegexOptions.IgnoreCase);
    if (match.Success)
    {
        var name = match.Groups[1].Value;
        try
        {
            var enc = Encoding.GetEncoding(name);
            return (enc, $"{enc.WebName} (XML declaration)");
        }
        catch
        {
            return (Encoding.UTF8, $"UTF-8 (XML declaration encoding '{name}' 미지원, 기본값으로 처리)");
        }
    }

    return (new UTF8Encoding(false), "UTF-8 (XML 1.0 기본값)");
}

static void PrintHelp()
{
    Console.WriteLine("PolyDonky.Convert.Xml — XML / XHTML ↔ IWPF 변환기");
    Console.WriteLine();
    Console.WriteLine("사용법:");
    Console.WriteLine("  PolyDonky.Convert.Xml <input> <output> [--fragment] [--title <text>]");
    Console.WriteLine("  PolyDonky.Convert.Xml --version | -v");
    Console.WriteLine("  PolyDonky.Convert.Xml --help    | -h | /?");
    Console.WriteLine();
    Console.WriteLine("변환 쌍:");
    Console.WriteLine("  *.xml|*.xhtml → *.iwpf : import (XHTML 자동 감지, 인코딩 자동 감지)");
    Console.WriteLine("  *.iwpf        → *.xml|*.xhtml : export (XHTML5 polyglot markup)");
    Console.WriteLine();
    Console.WriteLine("옵션 (export 시):");
    Console.WriteLine("  --fragment      <?xml ?>·<!DOCTYPE>·<html> 래퍼 없이 fragment 만 출력");
    Console.WriteLine("  --title <text>  <title> 요소 텍스트 지정 (생략 시 첫 H1 또는 기본값)");
    Console.WriteLine();
    Console.WriteLine("인코딩 감지 (import 시):");
    Console.WriteLine("  • BOM (UTF-8 / UTF-16 / UTF-32) 우선");
    Console.WriteLine("  • 첫 256 바이트 안의 XML 선언 <?xml encoding=\"X\"?>");
    Console.WriteLine("  • cp949·EUC-KR·Shift-JIS 등 레거시 코드페이지 지원");
    Console.WriteLine("  • 위 모두 실패 → UTF-8 (XML 1.0 기본값)");
    Console.WriteLine();
    Console.WriteLine("바이너리 거부:");
    Console.WriteLine("  앞 1KB 안에 NUL 바이트가 5% 이상이면 XML 이 아닌 것으로 판단해 거부 (exit 5)");
    Console.WriteLine();
    Console.WriteLine("보안 정책:");
    Console.WriteLine("  DTD 처리 비활성 — XXE / 외부 엔티티 공격 차단. DOCTYPE 선언이 있으면 거부.");
    Console.WriteLine();
    Console.WriteLine("종료 코드:");
    Console.WriteLine("  0  성공");
    Console.WriteLine("  2  인자 오류");
    Console.WriteLine("  3  지원하지 않는 변환 쌍");
    Console.WriteLine("  4  입출력 실패");
    Console.WriteLine("  5  변환 실패 (XML 형식 오류·DTD 거부·바이너리 입력)");
    Console.WriteLine();
    Console.WriteLine("진행 표시:");
    Console.WriteLine("  표준 출력에 PROGRESS:<percent>:<message> 형식으로 emit (감지된 인코딩 포함)");
    Console.WriteLine();
    Console.WriteLine("출력 안전성:");
    Console.WriteLine("  임시파일에 쓴 뒤 원자적 rename — 도중 종료(Ctrl+C 포함) 시 반쪽 파일이 남지 않음");
}

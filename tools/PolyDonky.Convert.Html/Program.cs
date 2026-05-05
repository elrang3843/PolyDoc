using System.Text;
using PolyDonky.Codecs.Html;
using PolyDonky.Core;
using PolyDonky.Iwpf;

// PolyDonky.Convert.Html — HTML ↔ IWPF 변환 전용 콘솔 도구.
// CLAUDE.md §3 의 외부 변환 모듈 분리 원칙: 메인 앱은 IWPF/MD/TXT 만 직접 처리하고
// HTML 은 이 CLI 가 처리한다.
//
// 사용법:
//   PolyDonky.Convert.Html <input> <output>
//   PolyDonky.Convert.Html --version | -v
//   PolyDonky.Convert.Html --help    | -h | /?
//
// 변환 쌍:
//   *.html|*.htm → *.iwpf : import (블록 한도 없음)
//   *.iwpf       → *.html|*.htm : export
//
// 종료 코드:
//   0 성공
//   2 인자 오류
//   3 지원하지 않는 변환 쌍
//   4 입출력 실패 (파일 없음·권한·디렉터리 없음·잠금)
//   5 변환 실패 (HTML 파싱·내부 예외)

const int ExitOk            = 0;
const int ExitBadArgs       = 2;
const int ExitUnsupportedOp = 3;
const int ExitIoError       = 4;
const int ExitConvertError  = 5;

try { Console.OutputEncoding = Encoding.UTF8; } catch { /* redirected pipe 등 무시 */ }

if (args.Length == 1 && (args[0] is "--version" or "-v"))
{
    Console.WriteLine("PolyDonky.Convert.Html 1.0");
    return ExitOk;
}

if (args.Length == 1 && (args[0] is "--help" or "-h" or "/?"))
{
    PrintHelp();
    return ExitOk;
}

if (args.Length != 2)
{
    Console.Error.WriteLine("Usage: PolyDonky.Convert.Html <input> <output>");
    Console.Error.WriteLine("  Supported: .html|.htm <-> .iwpf");
    Console.Error.WriteLine("  '--help' 로 자세한 도움말과 종료 코드 안내를 볼 수 있습니다.");
    return ExitBadArgs;
}

string inPath, outPath;
try
{
    inPath  = Path.GetFullPath(args[0]);
    outPath = Path.GetFullPath(args[1]);
}
catch (Exception ex)
{
    Console.Error.WriteLine($"경로 해석 실패: {ex.Message}");
    return ExitBadArgs;
}

string Ext(string p) => Path.GetExtension(p).TrimStart('.').ToLowerInvariant();
string inExt  = Ext(inPath);
string outExt = Ext(outPath);

// ── 검증 순서: 비용 낮은 검사부터 ────────────────────────────────────
if (string.Equals(inPath, outPath, StringComparison.OrdinalIgnoreCase))
{
    Console.Error.WriteLine($"입력과 출력 경로가 같습니다 — 자기 자신을 덮어쓸 수 없습니다: {inPath}");
    return ExitBadArgs;
}

bool isImport = (inExt is "html" or "htm") && outExt == "iwpf";
bool isExport = inExt == "iwpf" && (outExt is "html" or "htm");
if (!isImport && !isExport)
{
    Console.Error.WriteLine($"지원하지 않는 변환: .{inExt} → .{outExt}");
    Console.Error.WriteLine("  지원: .html|.htm → .iwpf, .iwpf → .html|.htm");
    return ExitUnsupportedOp;
}

if (!File.Exists(inPath))
{
    Console.Error.WriteLine($"입력 파일이 없습니다: {inPath}");
    return ExitIoError;
}

var inInfo = new FileInfo(inPath);
if (inInfo.Length == 0)
{
    Console.Error.WriteLine($"입력 파일이 비어 있습니다(0 byte): {inPath}");
    return ExitIoError;
}

var outDir = Path.GetDirectoryName(outPath);
if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir))
{
    try { Directory.CreateDirectory(outDir); }
    catch (Exception ex)
    {
        Console.Error.WriteLine($"출력 디렉터리 생성 실패: {outDir}\n  → {ex.Message}");
        return ExitIoError;
    }
}

// ── 변환 본체 (원자적 쓰기) ──────────────────────────────────────────
var tempOut = outPath + ".tmp-" + Guid.NewGuid().ToString("N").Substring(0, 8);

try
{
    if (isImport)
    {
        // 한도 0 = 무제한 — 사용자가 명시적으로 큰 HTML 을 변환하려고 했으므로 잘림 없이 모두 처리.
        WriteProgress(0, "HTML 읽는 중");
        var reader = new HtmlReader { MaxBlocks = 0 };
        PolyDonkyument doc;
        using (var fs = File.OpenRead(inPath))
            doc = reader.Read(fs);

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

        WriteProgress(60, "HTML 로 변환 중");
        using (var ofs = File.Create(tempOut))
            new HtmlWriter().Write(doc, ofs);
    }

    if (File.Exists(outPath)) File.Delete(outPath);
    File.Move(tempOut, outPath);
    WriteProgress(100, "완료");
    Console.WriteLine($"OK: {Path.GetFileName(inPath)} → {Path.GetFileName(outPath)}");
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
}

// ── 진행 막대 출력 ────────────────────────────────────────────────────
static void WriteProgress(int percent, string message)
{
    Console.WriteLine($"PROGRESS:{percent}:{message}");
    Console.Out.Flush();
}

static void PrintHelp()
{
    Console.WriteLine("PolyDonky.Convert.Html — HTML ↔ IWPF 변환기");
    Console.WriteLine();
    Console.WriteLine("사용법:");
    Console.WriteLine("  PolyDonky.Convert.Html <input> <output>");
    Console.WriteLine("  PolyDonky.Convert.Html --version | -v");
    Console.WriteLine("  PolyDonky.Convert.Html --help    | -h | /?");
    Console.WriteLine();
    Console.WriteLine("변환 쌍:");
    Console.WriteLine("  *.html|*.htm → *.iwpf : import (블록 한도 없음)");
    Console.WriteLine("  *.iwpf       → *.html|*.htm : export");
    Console.WriteLine();
    Console.WriteLine("종료 코드:");
    Console.WriteLine("  0  성공");
    Console.WriteLine("  2  인자 오류");
    Console.WriteLine("  3  지원하지 않는 변환 쌍");
    Console.WriteLine("  4  입출력 실패");
    Console.WriteLine("  5  변환 실패 (HTML 파싱·내부 예외)");
    Console.WriteLine();
    Console.WriteLine("진행 표시:");
    Console.WriteLine("  표준 출력에 PROGRESS:<percent>:<message> 형식으로 emit");
    Console.WriteLine();
    Console.WriteLine("출력 안전성:");
    Console.WriteLine("  임시파일에 쓴 뒤 원자적 rename — 도중 종료 시 반쪽 파일이 남지 않음");
}

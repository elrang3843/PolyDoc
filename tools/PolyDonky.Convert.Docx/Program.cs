using System.Text;
using DocumentFormat.OpenXml.Packaging;
using PolyDonky.Codecs.Docx;
using PolyDonky.Core;
using PolyDonky.Iwpf;

// PolyDonky.Convert.Docx — DOCX ↔ IWPF 변환 전용 콘솔 도구.
// CLAUDE.md §3 의 외부 변환 모듈 분리 원칙: 메인 앱은 IWPF/MD/TXT 만 직접 처리하고
// DOCX 는 이 CLI 가 처리한다.
//
// 사용법:
//   PolyDonky.Convert.Docx <input> <output>
//   PolyDonky.Convert.Docx --version | -v
//   PolyDonky.Convert.Docx --help    | -h | /?
//
// 변환 쌍:
//   *.docx → *.iwpf : import (Word 2013 / AppVersion 15.0 이상)
//   *.iwpf → *.docx : export
//
// 종료 코드:
//   0 성공
//   2 인자 오류
//   3 지원하지 않는 변환 쌍
//   4 입출력 실패 (파일 없음·권한·디렉터리 없음·디스크 잠금 등)
//   5 변환 실패 (DOCX 구조 손상·내부 예외)
//   6 지원하지 않는 옛 버전 (Word 2013 미만)

const int ExitOk            = 0;
const int ExitBadArgs       = 2;
const int ExitUnsupportedOp = 3;
const int ExitIoError       = 4;
const int ExitConvertError  = 5;
const int ExitOldVersion    = 6;

// Windows cmd.exe 의 기본 cp949 에서 한글 깨짐 방지.
try { Console.OutputEncoding = Encoding.UTF8; } catch { /* 일부 환경(redirected pipe) 무시 */ }

// Ctrl+C 시 임시파일 정리 — finally 블록이 실행되지 않는 SIGINT 경로에서도
// `.tmp-XXXXXXXX` 파일이 남지 않게 한다.
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

if (args.Length == 1 && (args[0] is "--version" or "-v"))
{
    Console.WriteLine("PolyDonky.Convert.Docx 1.0");
    return ExitOk;
}

if (args.Length == 1 && (args[0] is "--help" or "-h" or "/?"))
{
    PrintHelp();
    return ExitOk;
}

if (args.Length != 2)
{
    Console.Error.WriteLine("Usage: PolyDonky.Convert.Docx <input> <output>");
    Console.Error.WriteLine("  Supported: .docx <-> .iwpf");
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
// 1. 동일 경로 검사
if (string.Equals(inPath, outPath, StringComparison.OrdinalIgnoreCase))
{
    Console.Error.WriteLine($"입력과 출력 경로가 같습니다 — 자기 자신을 덮어쓸 수 없습니다: {inPath}");
    return ExitBadArgs;
}

// 2. 변환 쌍 검사
bool isImport = inExt == "docx" && outExt == "iwpf";
bool isExport = inExt == "iwpf" && outExt == "docx";
if (!isImport && !isExport)
{
    Console.Error.WriteLine($"지원하지 않는 변환: .{inExt} → .{outExt}");
    Console.Error.WriteLine("  지원: .docx → .iwpf, .iwpf → .docx");
    return ExitUnsupportedOp;
}

// 3. 입력 파일 존재
if (!File.Exists(inPath))
{
    Console.Error.WriteLine($"입력 파일이 없습니다: {inPath}");
    return ExitIoError;
}

// 4. 입력 파일 비어있지 않음
var inInfo = new FileInfo(inPath);
if (inInfo.Length == 0)
{
    Console.Error.WriteLine($"입력 파일이 비어 있습니다(0 byte): {inPath}");
    return ExitIoError;
}

// 5. 출력 디렉터리 보장
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

// ── 변환 본체 ────────────────────────────────────────────────────────
// 원자적 쓰기 — 변환 도중 비정상 종료 시 반쪽 출력이 남지 않도록 임시파일 → 최종 이름 이동.
var tempOut = outPath + ".tmp-" + Guid.NewGuid().ToString("N").Substring(0, 8);
tempCleanupPath = tempOut;

try
{
    if (isImport)
    {
        // 버전 정책 — Word 2013 (AppVersion 15.0) 이상.
        var av = DocxVersionPolicy.ReadAppVersion(inPath);
        if (!DocxVersionPolicy.IsSupported(av))
        {
            Console.Error.WriteLine(
                $"지원하지 않는 버전 — DOCX AppVersion {av:0.0}. " +
                $"PolyDonky 는 Word 2013 (AppVersion {DocxVersionPolicy.MinSupportedAppVersion:0.0}) 이상만 처리합니다.");
            return ExitOldVersion;
        }

        WriteProgress(0, $"DOCX 읽는 중 (AppVersion {av?.ToString("0.0") ?? "?"})");
        PolyDonkyument doc;
        using (var fs = File.OpenRead(inPath))
            doc = new DocxReader().Read(fs);

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

        WriteProgress(60, "DOCX 로 변환 중");
        using (var ofs = File.Create(tempOut))
            new DocxWriter().Write(doc, ofs);
    }

    // 원자적 이동 — 동일 파일 시스템에선 원자적, 그 외는 copy+delete fallback.
    if (File.Exists(outPath)) File.Delete(outPath);
    File.Move(tempOut, outPath);
    tempCleanupPath = null;  // Move 성공 후엔 SIGINT 핸들러가 outPath 를 지우지 않게 함.
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
catch (OpenXmlPackageException ex)
{
    Console.Error.WriteLine(
        $"DOCX 패키지가 손상되었거나 유효하지 않습니다: {inPath}\n" +
        $"  세부: {ex.Message}");
    return ExitConvertError;
}
catch (System.IO.FileFormatException ex)
{
    // System.IO.Packaging 이 ZIP/OPC 컨테이너를 인식하지 못할 때 — 사실상 DOCX 가 아님.
    Console.Error.WriteLine(
        $"DOCX 형식이 아닙니다 (OOXML 컨테이너로 인식 안 됨): {inPath}\n" +
        $"  세부: {ex.Message}");
    return ExitConvertError;
}
catch (System.IO.InvalidDataException ex)
{
    Console.Error.WriteLine(
        $"파일 형식이 유효하지 않습니다 (ZIP/OOXML 파싱 실패): {inPath}\n" +
        $"  세부: {ex.Message}");
    return ExitConvertError;
}
catch (IOException ex)
{
    // 디스크 가득·다른 프로세스의 잠금·네트워크 끊김 등.
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
    // 임시파일 정리 — File.Move 가 성공했으면 이미 사라졌으므로 Exists 체크가 안전.
    try { if (File.Exists(tempOut)) File.Delete(tempOut); } catch { /* 무시 */ }
    Console.Error.Flush();
    Console.Out.Flush();
}

// ── 진행 막대 출력 ────────────────────────────────────────────────────
// 메인 앱(Services/ExternalConverter) 가 stdout 을 한 줄씩 읽고 "PROGRESS:" 접두로
// 시작하는 줄을 파싱해 IProgress 로 보고한다.
static void WriteProgress(int percent, string message)
{
    Console.WriteLine($"PROGRESS:{percent}:{message}");
    Console.Out.Flush();
}

static void PrintHelp()
{
    Console.WriteLine("PolyDonky.Convert.Docx — DOCX ↔ IWPF 변환기");
    Console.WriteLine();
    Console.WriteLine("사용법:");
    Console.WriteLine("  PolyDonky.Convert.Docx <input> <output>");
    Console.WriteLine("  PolyDonky.Convert.Docx --version | -v");
    Console.WriteLine("  PolyDonky.Convert.Docx --help    | -h | /?");
    Console.WriteLine();
    Console.WriteLine("변환 쌍:");
    Console.WriteLine("  *.docx → *.iwpf : import (Word 2013 / AppVersion 15.0 이상)");
    Console.WriteLine("  *.iwpf → *.docx : export");
    Console.WriteLine();
    Console.WriteLine("종료 코드:");
    Console.WriteLine("  0  성공");
    Console.WriteLine("  2  인자 오류");
    Console.WriteLine("  3  지원하지 않는 변환 쌍");
    Console.WriteLine("  4  입출력 실패 (파일 없음·권한·디렉터리 없음·잠금)");
    Console.WriteLine("  5  변환 실패 (DOCX 구조 손상·내부 예외)");
    Console.WriteLine("  6  지원하지 않는 옛 버전");
    Console.WriteLine();
    Console.WriteLine("진행 표시:");
    Console.WriteLine("  표준 출력에 PROGRESS:<percent>:<message> 형식으로 emit");
    Console.WriteLine();
    Console.WriteLine("출력 안전성:");
    Console.WriteLine("  임시파일에 쓴 뒤 원자적 rename — 도중 종료 시 반쪽 파일이 남지 않음");
}

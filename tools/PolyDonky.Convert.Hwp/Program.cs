using System.Text;
using PolyDonky.Codecs.Docx;
using PolyDonky.Convert.Common;
using PolyDonky.Core;
using PolyDonky.Iwpf;

// PolyDonky.Convert.Hwp — HWP ↔ IWPF 변환 전용 콘솔 도구.
// CLAUDE.md §3 의 외부 변환 모듈 분리 원칙: 메인 앱은 IWPF/MD/TXT 만 직접 처리하고
// HWP 는 이 CLI 가 처리한다.
//
// 변환 파이프라인:
//   *.hwp → *.iwpf : LibreOffice headless (hwp→docx) → DocxReader → IwpfWriter
//   *.iwpf → *.hwp : IwpfReader → DocxWriter → LibreOffice headless (docx→hwp)
//
// 주의: LibreOffice 의 HWP 지원 수준에 따라 충실도가 제한될 수 있습니다.
//        한컴 오피스에서 최종 확인을 권장합니다.
//
// 사용법:
//   PolyDonky.Convert.Hwp <input> <output>
//   PolyDonky.Convert.Hwp --version | -v
//   PolyDonky.Convert.Hwp --help    | -h | /?
//
// 종료 코드:
//   0 성공  2 인자 오류  3 지원하지 않는 변환 쌍
//   4 입출력 실패  5 변환 실패  6 LibreOffice 미설치
// (상수는 PolyDonky.Convert.Common.ConverterExitCodes 에 정의됨)

try { Console.OutputEncoding = Encoding.UTF8; } catch { }

string? tempCleanupDir = null;
Console.CancelKeyPress += (_, e) =>
{
    if (tempCleanupDir is not null && Directory.Exists(tempCleanupDir))
        try { Directory.Delete(tempCleanupDir, recursive: true); } catch { }
    Console.Error.Flush();
    Console.Out.Flush();
    e.Cancel = false;
};

if (args.Length == 1 && (args[0] is "--version" or "-v"))
{
    Console.WriteLine("PolyDonky.Convert.Hwp 1.0");
    return ConverterExitCodes.Ok;
}

if (args.Length == 1 && (args[0] is "--help" or "-h" or "/?"))
{
    PrintHelp();
    return ConverterExitCodes.Ok;
}

if (args.Length != 2)
{
    Console.Error.WriteLine("Usage: PolyDonky.Convert.Hwp <input> <output>");
    Console.Error.WriteLine("  Supported: .hwp <-> .iwpf");
    return ConverterExitCodes.BadArgs;
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
    return ConverterExitCodes.BadArgs;
}

string Ext(string p) => Path.GetExtension(p).TrimStart('.').ToLowerInvariant();
string inExt  = Ext(inPath);
string outExt = Ext(outPath);

if (string.Equals(inPath, outPath, StringComparison.OrdinalIgnoreCase))
{
    Console.Error.WriteLine("입력과 출력 경로가 같습니다.");
    return ConverterExitCodes.BadArgs;
}

bool isImport = inExt == "hwp" && outExt == "iwpf";
bool isExport = inExt == "iwpf" && outExt == "hwp";
if (!isImport && !isExport)
{
    Console.Error.WriteLine($"지원하지 않는 변환: .{inExt} → .{outExt}");
    Console.Error.WriteLine("  지원: .hwp → .iwpf, .iwpf → .hwp");
    return ConverterExitCodes.UnsupportedOp;
}

if (!File.Exists(inPath))
{
    Console.Error.WriteLine($"입력 파일이 없습니다: {inPath}");
    return ConverterExitCodes.IoError;
}

if (new FileInfo(inPath).Length == 0)
{
    Console.Error.WriteLine($"입력 파일이 비어 있습니다(0 byte): {inPath}");
    return ConverterExitCodes.IoError;
}

// LibreOffice 유무 확인
var sOffice = LibreOfficeBridge.FindSOffice();
if (sOffice is null)
{
    Console.Error.WriteLine(
        "LibreOffice 를 찾을 수 없습니다. LibreOffice 를 설치하거나 " +
        "LIBREOFFICE_PATH 환경변수에 설치 폴더 경로를 지정하세요.");
    return ConverterExitCodes.OldVersion;  // 6 = 지원 범위 밖 (LibreOffice 없음)
}

var outDir = Path.GetDirectoryName(outPath);
if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir))
{
    try { Directory.CreateDirectory(outDir); }
    catch (Exception ex)
    {
        Console.Error.WriteLine($"출력 디렉터리 생성 실패: {outDir}\n  → {ex.Message}");
        return ConverterExitCodes.IoError;
    }
}

var tmpDir  = Path.Combine(Path.GetTempPath(), $"polydonky-hwp-{Guid.NewGuid():N}");
tempCleanupDir = tmpDir;
var tempOut = outPath + ".tmp-" + Guid.NewGuid().ToString("N")[..8];

try
{
    if (isImport)
    {
        // HWP → (LibreOffice) → 임시 DOCX → (DocxReader) → IWPF
        ConverterProgress.Write(0, "LibreOffice 로 HWP → DOCX 변환 중");

        var progress = new Progress<(int Percent, string Message)>(t =>
            ConverterProgress.Write(t.Percent, t.Message));

        var tmpDocx = await LibreOfficeBridge.ConvertToDocxAsync(
            inPath, tmpDir, sOfficePath: sOffice, progress: progress);

        ConverterProgress.Write(50, "DOCX 읽는 중");
        PolyDonkyument doc;
        using (var fs = File.OpenRead(tmpDocx))
            doc = new DocxReader().Read(fs);

        ConverterProgress.Write(80, "IWPF 로 변환 중");
        using (var ofs = File.Create(tempOut))
            new IwpfWriter().Write(doc, ofs);
    }
    else // isExport
    {
        // IWPF → (IwpfReader) → (DocxWriter) → 임시 DOCX → (LibreOffice) → HWP
        ConverterProgress.Write(0, "IWPF 읽는 중");
        PolyDonkyument doc;
        using (var fs = File.OpenRead(inPath))
            doc = new IwpfReader().Read(fs);

        ConverterProgress.Write(30, "DOCX 로 변환 중");
        var tmpDocxPath = Path.Combine(tmpDir, Path.GetFileNameWithoutExtension(outPath) + ".docx");
        Directory.CreateDirectory(tmpDir);
        using (var ofs = File.Create(tmpDocxPath))
            new DocxWriter().Write(doc, ofs);

        ConverterProgress.Write(60, "LibreOffice 로 DOCX → HWP 변환 중");
        var progress = new Progress<(int Percent, string Message)>(t =>
            ConverterProgress.Write(t.Percent, t.Message));

        var hwpResult = await LibreOfficeBridge.ConvertFromDocxAsync(
            tmpDocxPath, tmpDir, "hwp", sOfficePath: sOffice, progress: progress);

        File.Copy(hwpResult, tempOut, overwrite: true);
    }

    if (File.Exists(outPath)) File.Delete(outPath);
    File.Move(tempOut, outPath);
    tempCleanupDir = null;
    ConverterProgress.Write(100, "완료");
    Console.WriteLine($"OK: {Path.GetFileName(inPath)} → {Path.GetFileName(outPath)}");
    Console.Out.Flush();
    return ConverterExitCodes.Ok;
}
catch (FileNotFoundException ex)
{
    Console.Error.WriteLine($"파일을 찾을 수 없습니다: {ex.FileName ?? inPath}");
    return ConverterExitCodes.IoError;
}
catch (IOException ex)
{
    Console.Error.WriteLine($"I/O 실패: {ex.Message}");
    return ConverterExitCodes.IoError;
}
catch (Exception ex)
{
    Console.Error.WriteLine($"변환 실패: {ex.GetType().Name}: {ex.Message}");
    Console.Error.WriteLine(ex.StackTrace);
    return ConverterExitCodes.ConvertError;
}
finally
{
    if (File.Exists(tempOut)) try { File.Delete(tempOut); } catch { }
    if (tempCleanupDir is not null && Directory.Exists(tempCleanupDir))
        try { Directory.Delete(tempCleanupDir, recursive: true); } catch { }
    Console.Error.Flush();
    Console.Out.Flush();
}

static void PrintHelp()
{
    Console.WriteLine("PolyDonky.Convert.Hwp — HWP ↔ IWPF 변환기");
    Console.WriteLine();
    Console.WriteLine("사용법:");
    Console.WriteLine("  PolyDonky.Convert.Hwp <input> <output>");
    Console.WriteLine();
    Console.WriteLine("변환 쌍:");
    Console.WriteLine("  *.hwp  → *.iwpf : import (LibreOffice → DOCX → IWPF)");
    Console.WriteLine("  *.iwpf → *.hwp  : export (IWPF → DOCX → LibreOffice)");
    Console.WriteLine();
    Console.WriteLine("주의: LibreOffice HWP 지원 수준에 따라 충실도가 제한될 수 있습니다.");
    Console.WriteLine();
    Console.WriteLine("종료 코드:");
    Console.WriteLine("  0  성공");
    Console.WriteLine("  2  인자 오류");
    Console.WriteLine("  3  지원하지 않는 변환 쌍");
    Console.WriteLine("  4  입출력 실패");
    Console.WriteLine("  5  변환 실패");
    Console.WriteLine("  6  LibreOffice 미설치 또는 탐지 불가");
    Console.WriteLine();
    Console.WriteLine("LibreOffice 경로 지정:");
    Console.WriteLine("  LIBREOFFICE_PATH 환경변수에 LibreOffice 설치 폴더를 지정");
}

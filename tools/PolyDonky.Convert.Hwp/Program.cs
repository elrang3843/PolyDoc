using System.Text;
using PolyDonky.Codecs.Docx;
using PolyDonky.Convert.Common;
using PolyDonky.Convert.Hwp;
using PolyDonky.Core;
using PolyDonky.Iwpf;

// PolyDonky.Convert.Hwp — HWP ↔ IWPF 변환 전용 콘솔 도구.
// CLAUDE.md §3 의 외부 변환 모듈 분리 원칙: 메인 앱은 IWPF/MD/TXT 만 직접 처리하고
// HWP 는 이 CLI 가 처리한다.
//
// 변환 파이프라인:
//   *.hwp → *.iwpf : (아직 미구현, LibreOffice 필요)
//   *.iwpf → *.hwp : IwpfReader → HwpWriter (커스텀 구현)
//
// 사용법:
//   PolyDonky.Convert.Hwp <input> <output>
//   PolyDonky.Convert.Hwp --version | -v
//   PolyDonky.Convert.Hwp --help    | -h | /?
//
// 종료 코드:
//   0 성공  2 인자 오류  3 지원하지 않는 변환 쌍
//   4 입출력 실패  5 변환 실패
// (상수는 PolyDonky.Convert.Common.ConverterExitCodes 에 정의됨)

try { Console.OutputEncoding = Encoding.UTF8; } catch { }

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

// 현재 단계: IWPF → HWP만 지원 (읽기는 나중)
if (inExt == "hwp" && outExt == "iwpf")
{
    Console.Error.WriteLine("HWP → IWPF 변환은 아직 미구현되었습니다.");
    return ConverterExitCodes.UnsupportedOp;
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

var tempOut = outPath + ".tmp-" + Guid.NewGuid().ToString("N")[..8];

try
{
    if (isImport)
    {
        Console.Error.WriteLine("HWP → IWPF 변환은 아직 미구현되었습니다.");
        return ConverterExitCodes.UnsupportedOp;
    }
    else // isExport
    {
        // IWPF → HwpWriter → HWP
        ConverterProgress.Write(0, "IWPF 읽는 중");
        PolyDonkyument doc;
        using (var fs = File.OpenRead(inPath))
            doc = new IwpfReader().Read(fs);

        ConverterProgress.Write(50, "HWP 로 변환 중");
        using (var ofs = File.Create(tempOut))
            new HwpWriter().Write(doc, ofs);
    }

    if (File.Exists(outPath)) File.Delete(outPath);
    File.Move(tempOut, outPath);
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
    Console.WriteLine("변환 쌍 (현재 단계):");
    Console.WriteLine("  *.iwpf → *.hwp  : export (텍스트)");
    Console.WriteLine();
    Console.WriteLine("향후:");
    Console.WriteLine("  *.hwp  → *.iwpf : import (파서 구현 예정)");
    Console.WriteLine();
    Console.WriteLine("종료 코드:");
    Console.WriteLine("  0  성공");
    Console.WriteLine("  2  인자 오류");
    Console.WriteLine("  3  지원하지 않는 변환 쌍");
    Console.WriteLine("  4  입출력 실패");
    Console.WriteLine("  5  변환 실패");
}

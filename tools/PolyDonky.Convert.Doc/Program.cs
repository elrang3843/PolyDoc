using System.Text;
using PolyDonky.Convert.Common;
using PolyDonky.Convert.Doc;
using PolyDonky.Core;
using PolyDonky.Iwpf;

// PolyDonky.Convert.Doc — IWPF ↔ RTF 변환 전용 콘솔 도구.
// CLAUDE.md §3 의 외부 변환 모듈 분리 원칙: 메인 앱은 IWPF/MD/TXT 만 직접 처리하고
// RTF 는 이 CLI 가 처리한다.
//
// 변환 파이프라인:
//   *.iwpf → *.rtf : IwpfReader → DocWriter (RTF 생성기) → 저장
//
// 참고: DOC (Word 97-2003 OLE2 포맷)는 v1.0.0 이후에 지원 예정 (Aspose.Words)
//
// 사용법:
//   PolyDonky.Convert.Doc <input> <output>
//   PolyDonky.Convert.Doc --version | -v
//   PolyDonky.Convert.Doc --help    | -h | /?
//
// 종료 코드:
//   0 성공  2 인자 오류  3 지원하지 않는 변환 쌍
//   4 입출력 실패  5 변환 실패
// (상수는 PolyDonky.Convert.Common.ConverterExitCodes 에 정의됨)
//
// RTF 형식 특징:
//   - 배경색(하이라이트) 완벽 지원
//   - Word 97 이상에서 100% 호환
//   - 텍스트 기반 형식, 가볍고 안정적

try { Console.OutputEncoding = Encoding.UTF8; } catch { }

if (args.Length == 1 && (args[0] is "--version" or "-v"))
{
    Console.WriteLine("PolyDonky.Convert.Doc 1.0");
    return ConverterExitCodes.Ok;
}

if (args.Length == 1 && (args[0] is "--help" or "-h" or "/?"))
{
    PrintHelp();
    return ConverterExitCodes.Ok;
}

if (args.Length != 2)
{
    Console.Error.WriteLine("Usage: PolyDonky.Convert.Doc <input> <output>");
    Console.Error.WriteLine("  Supported: .iwpf → .rtf");
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

// 현재 단계: IWPF → RTF만 지원
bool isExport = inExt == "iwpf" && outExt == "rtf";
if (!isExport)
{
    Console.Error.WriteLine($"지원하지 않는 변환: .{inExt} → .{outExt}");
    Console.Error.WriteLine("  지원: .iwpf → .rtf");
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

try
{
    ConverterProgress.Write(0, "IWPF 읽는 중");
    PolyDonkyument doc;
    using (var fs = File.OpenRead(inPath))
        doc = new IwpfReader().Read(fs);

    ConverterProgress.Write(50, "RTF 로 변환 중");
    using (var ofs = File.Create(outPath))
        new DocWriter().Write(doc, ofs);

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
    Console.Error.Flush();
    Console.Out.Flush();
}

static void PrintHelp()
{
    Console.WriteLine("PolyDonky.Convert.Doc — IWPF ↔ RTF 변환기");
    Console.WriteLine();
    Console.WriteLine("사용법:");
    Console.WriteLine("  PolyDonky.Convert.Doc <input> <output>");
    Console.WriteLine();
    Console.WriteLine("변환 쌍 (현재 단계):");
    Console.WriteLine("  *.iwpf → *.rtf  : export (배경색 지원)");
    Console.WriteLine();
    Console.WriteLine("향후 (v1.0.0 이후 언젠가):");
    Console.WriteLine("  *.iwpf → *.doc  : export (Word 97-2003 OLE2 바이너리 포맷)");
    Console.WriteLine();
    Console.WriteLine("종료 코드:");
    Console.WriteLine("  0  성공");
    Console.WriteLine("  2  인자 오류");
    Console.WriteLine("  3  지원하지 않는 변환 쌍");
    Console.WriteLine("  4  입출력 실패");
    Console.WriteLine("  5  변환 실패");
}

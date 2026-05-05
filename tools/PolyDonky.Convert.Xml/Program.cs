using System.Text;
using PolyDonky.Core;
using PolyDonky.Iwpf;
using PdXmlReader = PolyDonky.Codecs.Xml.XmlReader;
using PdXmlWriter = PolyDonky.Codecs.Xml.XmlWriter;

// PolyDonky.Convert.Xml — XML / XHTML ↔ IWPF 변환 전용 콘솔 도구.
// CLAUDE.md §3 의 외부 변환 모듈 분리 원칙: 메인 앱은 IWPF/MD/TXT 만 직접 처리하고
// XML/XHTML 은 이 CLI 가 처리한다.
//
// 사용법:
//   PolyDonky.Convert.Xml <input> <output>
//   PolyDonky.Convert.Xml --version | -v
//   PolyDonky.Convert.Xml --help    | -h | /?
//
// 변환 쌍:
//   *.xml|*.xhtml → *.iwpf : import (XHTML 자동 감지, 일반 XML 은 텍스트 추출)
//   *.iwpf        → *.xml|*.xhtml : export (XHTML5 polyglot markup)
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

if (args.Length == 1 && (args[0] is "--version" or "-v"))
{
    Console.WriteLine("PolyDonky.Convert.Xml 1.0");
    return ExitOk;
}

if (args.Length == 1 && (args[0] is "--help" or "-h" or "/?"))
{
    PrintHelp();
    return ExitOk;
}

if (args.Length != 2)
{
    Console.Error.WriteLine("Usage: PolyDonky.Convert.Xml <input> <output>");
    Console.Error.WriteLine("  Supported: .xml|.xhtml <-> .iwpf");
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

if (string.Equals(inPath, outPath, StringComparison.OrdinalIgnoreCase))
{
    Console.Error.WriteLine($"입력과 출력 경로가 같습니다 — 자기 자신을 덮어쓸 수 없습니다: {inPath}");
    return ExitBadArgs;
}

bool isImport = (inExt is "xml" or "xhtml") && outExt == "iwpf";
bool isExport = inExt == "iwpf" && (outExt is "xml" or "xhtml");
if (!isImport && !isExport)
{
    Console.Error.WriteLine($"지원하지 않는 변환: .{inExt} → .{outExt}");
    Console.Error.WriteLine("  지원: .xml|.xhtml → .iwpf, .iwpf → .xml|.xhtml");
    return ExitUnsupportedOp;
}

if (!File.Exists(inPath))
{
    Console.Error.WriteLine($"입력 파일이 없습니다: {inPath}");
    return ExitIoError;
}

if (new FileInfo(inPath).Length == 0)
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

var tempOut = outPath + ".tmp-" + Guid.NewGuid().ToString("N").Substring(0, 8);

try
{
    if (isImport)
    {
        WriteProgress(0, "XML 읽는 중");
        PolyDonkyument doc;
        using (var fs = File.OpenRead(inPath))
            doc = new PdXmlReader().Read(fs);

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

        WriteProgress(60, "XHTML 로 변환 중");
        using (var ofs = File.Create(tempOut))
            new PdXmlWriter().Write(doc, ofs);
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
catch (System.Xml.XmlException ex)
{
    Console.Error.WriteLine(
        $"XML 형식이 유효하지 않습니다: {inPath}\n" +
        $"  줄 {ex.LineNumber}, 위치 {ex.LinePosition}: {ex.Message}");
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
}

static void WriteProgress(int percent, string message)
{
    Console.WriteLine($"PROGRESS:{percent}:{message}");
    Console.Out.Flush();
}

static void PrintHelp()
{
    Console.WriteLine("PolyDonky.Convert.Xml — XML / XHTML ↔ IWPF 변환기");
    Console.WriteLine();
    Console.WriteLine("사용법:");
    Console.WriteLine("  PolyDonky.Convert.Xml <input> <output>");
    Console.WriteLine("  PolyDonky.Convert.Xml --version | -v");
    Console.WriteLine("  PolyDonky.Convert.Xml --help    | -h | /?");
    Console.WriteLine();
    Console.WriteLine("변환 쌍:");
    Console.WriteLine("  *.xml|*.xhtml → *.iwpf : import");
    Console.WriteLine("  *.iwpf        → *.xml|*.xhtml : export (XHTML5 polyglot markup)");
    Console.WriteLine();
    Console.WriteLine("종료 코드:");
    Console.WriteLine("  0  성공");
    Console.WriteLine("  2  인자 오류");
    Console.WriteLine("  3  지원하지 않는 변환 쌍");
    Console.WriteLine("  4  입출력 실패");
    Console.WriteLine("  5  변환 실패 (XML 형식 오류·내부 예외)");
    Console.WriteLine();
    Console.WriteLine("진행 표시:");
    Console.WriteLine("  표준 출력에 PROGRESS:<percent>:<message> 형식으로 emit");
    Console.WriteLine();
    Console.WriteLine("출력 안전성:");
    Console.WriteLine("  임시파일에 쓴 뒤 원자적 rename — 도중 종료 시 반쪽 파일이 남지 않음");
}

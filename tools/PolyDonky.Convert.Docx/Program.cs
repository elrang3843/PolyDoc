using PolyDonky.Codecs.Docx;
using PolyDonky.Core;
using PolyDonky.Iwpf;

// PolyDonky.Convert.Docx — DOCX ↔ IWPF 변환 전용 콘솔 도구.
// CLAUDE.md §3 의 외부 변환 모듈 분리 원칙: 메인 앱은 IWPF/MD/TXT 만 직접 처리하고
// DOCX 는 이 CLI 가 처리한다.
//
// 사용법:
//   PolyDonky.Convert.Docx <input> <output>
//     - input  *.docx,  output *.iwpf  → import (DOCX → IWPF)
//     - input  *.iwpf,  output *.docx  → export (IWPF → DOCX)
//   PolyDonky.Convert.Docx --version
//
// 종료 코드: 0 성공, 2 인자 오류, 3 지원 안 함, 4 입출력 실패, 5 변환 실패.

if (args.Length == 1 && (args[0] == "--version" || args[0] == "-v"))
{
    Console.WriteLine("PolyDonky.Convert.Docx 1.0");
    return 0;
}

if (args.Length != 2)
{
    Console.Error.WriteLine("Usage: PolyDonky.Convert.Docx <input> <output>");
    Console.Error.WriteLine("  Supported pairs: .docx <-> .iwpf");
    return 2;
}

string inPath  = args[0];
string outPath = args[1];

string Ext(string p) => Path.GetExtension(p).TrimStart('.').ToLowerInvariant();
string inExt  = Ext(inPath);
string outExt = Ext(outPath);

if (!File.Exists(inPath))
{
    Console.Error.WriteLine($"입력 파일이 없습니다: {inPath}");
    return 4;
}

try
{
    if (inExt == "docx" && outExt == "iwpf")
    {
        // 버전 정책 — Word 2013 (AppVersion 15.0) 이상만 지원.
        var av = DocxVersionPolicy.ReadAppVersion(inPath);
        if (!DocxVersionPolicy.IsSupported(av))
        {
            Console.Error.WriteLine(
                $"지원하지 않는 버전 — DOCX AppVersion {av:0.0}. " +
                $"PolyDonky 는 Word 2013 (AppVersion {DocxVersionPolicy.MinSupportedAppVersion:0.0}) 이상만 처리합니다.");
            return 6;
        }
        WriteProgress(0, $"DOCX 읽는 중 (AppVersion {av?.ToString("0.0") ?? "?"})");
        PolyDonkyument doc;
        using (var fs = File.OpenRead(inPath))
            doc = new DocxReader().Read(fs);
        WriteProgress(60, "IWPF 로 변환 중");
        using var ofs = File.Create(outPath);
        new IwpfWriter().Write(doc, ofs);
        WriteProgress(100, "완료");
        Console.WriteLine($"OK: {Path.GetFileName(inPath)} → {Path.GetFileName(outPath)}");
        return 0;
    }

    if (inExt == "iwpf" && outExt == "docx")
    {
        WriteProgress(0, "IWPF 읽는 중");
        PolyDonkyument doc;
        using (var fs = File.OpenRead(inPath))
            doc = new IwpfReader().Read(fs);
        WriteProgress(60, "DOCX 로 변환 중");
        using var ofs = File.Create(outPath);
        new DocxWriter().Write(doc, ofs);
        WriteProgress(100, "완료");
        Console.WriteLine($"OK: {Path.GetFileName(inPath)} → {Path.GetFileName(outPath)}");
        return 0;
    }

    Console.Error.WriteLine($"지원하지 않는 변환: .{inExt} → .{outExt}");
    return 3;
}
catch (Exception ex)
{
    Console.Error.WriteLine($"변환 실패: {ex.GetType().Name}: {ex.Message}");
    return 5;
}

static void WriteProgress(int percent, string message)
{
    Console.WriteLine($"PROGRESS:{percent}:{message}");
    Console.Out.Flush();
}

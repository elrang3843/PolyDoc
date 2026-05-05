using PolyDonky.Codecs.Hwpx;
using PolyDonky.Core;
using PolyDonky.Iwpf;

// PolyDonky.Convert.Hwpx — HWPX ↔ IWPF 변환 전용 콘솔 도구.
// CLAUDE.md §3 의 외부 변환 모듈 분리 원칙: 메인 앱은 IWPF/MD/TXT 만 직접 처리하고
// HWPX 는 이 CLI 가 처리한다.
//
// 사용법:
//   PolyDonky.Convert.Hwpx <input> <output>
//     - input  *.hwpx,  output *.iwpf  → import (HWPX → IWPF)
//     - input  *.iwpf,  output *.hwpx  → export (IWPF → HWPX)
//   PolyDonky.Convert.Hwpx --version

if (args.Length == 1 && (args[0] == "--version" || args[0] == "-v"))
{
    Console.WriteLine("PolyDonky.Convert.Hwpx 1.0");
    return 0;
}

if (args.Length != 2)
{
    Console.Error.WriteLine("Usage: PolyDonky.Convert.Hwpx <input> <output>");
    Console.Error.WriteLine("  Supported pairs: .hwpx <-> .iwpf");
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
    if (inExt == "hwpx" && outExt == "iwpf")
    {
        // 버전 정책 — HWPX 1.2 (HWP 2014) 이상만 지원.
        var ver = HwpxVersionPolicy.ReadXmlVersion(inPath);
        if (!HwpxVersionPolicy.IsSupported(ver))
        {
            Console.Error.WriteLine(
                $"지원하지 않는 버전 — HWPX 스키마 {ver}. " +
                $"PolyDonky 는 HWPX {HwpxVersionPolicy.MinSupportedXmlVersion} 이상(HWP 2014 이후)만 처리합니다.");
            return 6;
        }
        WriteProgress(0, $"HWPX 읽는 중 (xmlVersion {ver ?? "?"})");
        PolyDonkyument doc;
        using (var fs = File.OpenRead(inPath))
            doc = new HwpxReader().Read(fs);
        WriteProgress(60, "IWPF 로 변환 중");
        using var ofs = File.Create(outPath);
        new IwpfWriter().Write(doc, ofs);
        WriteProgress(100, "완료");
        Console.WriteLine($"OK: {Path.GetFileName(inPath)} → {Path.GetFileName(outPath)}");
        return 0;
    }

    if (inExt == "iwpf" && outExt == "hwpx")
    {
        WriteProgress(0, "IWPF 읽는 중");
        PolyDonkyument doc;
        using (var fs = File.OpenRead(inPath))
            doc = new IwpfReader().Read(fs);
        WriteProgress(60, "HWPX 로 변환 중");
        using var ofs = File.Create(outPath);
        new HwpxWriter().Write(doc, ofs);
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

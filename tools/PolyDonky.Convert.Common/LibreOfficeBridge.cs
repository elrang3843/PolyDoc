using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace PolyDonky.Convert.Common;

/// <summary>
/// LibreOffice headless 를 이용해 DOC/HWP 파일을 DOCX 로 변환하는 브리지.
/// 파이프라인: DOC/HWP → soffice headless → DOCX → PolyDonky.Codecs.Docx → IWPF
/// </summary>
public static class LibreOfficeBridge
{
    /// <summary>
    /// LibreOffice soffice 실행 파일 경로.
    /// LIBREOFFICE_PATH 환경변수 → 일반 설치 경로 순으로 탐색.
    /// </summary>
    public static string? FindSOffice()
    {
        // 1순위: 환경변수 (메인 앱이 주입)
        // LIBREOFFICE_PATH 는 설치 폴더를 가리킨다 (예: C:\Program Files\LibreOffice).
        // Windows 표준 설치에서 soffice.exe 는 program\ 하위에 있다.
        var envPath = Environment.GetEnvironmentVariable("LIBREOFFICE_PATH");
        if (!string.IsNullOrEmpty(envPath))
        {
            string[] envCandidates =
            [
                Path.Combine(envPath, "program", "soffice.exe"),  // Windows 표준
                Path.Combine(envPath, "program", "soffice"),      // Linux
                Path.Combine(envPath, "bin",     "soffice.exe"),  // 구형/대안
                Path.Combine(envPath, "bin",     "soffice"),      // Linux 대안
                Path.Combine(envPath,            "soffice.exe"),  // envPath 가 program 폴더인 경우
                Path.Combine(envPath,            "soffice"),
            ];
            foreach (var c in envCandidates)
                if (File.Exists(c)) return c;
        }

        // 2순위: 일반 설치 경로
        string[] commonPaths =
        [
            @"C:\Program Files\LibreOffice\program\soffice.exe",
            @"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
            "/usr/bin/soffice",
            "/usr/local/bin/soffice",
            "/opt/libreoffice/program/soffice",
            "/opt/libreoffice/bin/soffice",
        ];

        foreach (var p in commonPaths)
            if (File.Exists(p)) return p;

        return null;
    }

    /// <summary>
    /// LibreOffice headless 로 입력 파일을 DOCX 로 변환.
    /// 출력 DOCX 경로를 반환 (outDir 아래 동일 이름+.docx).
    /// </summary>
    /// <param name="inputPath">변환할 DOC/HWP 파일 경로</param>
    /// <param name="outDir">DOCX 출력 디렉터리</param>
    /// <param name="sOfficePath">soffice 실행 파일 경로 (null이면 자동 탐지)</param>
    /// <param name="progress">진행률 보고 (0-40 범위 사용 권장)</param>
    /// <param name="timeoutMs">타임아웃 밀리초 (기본 120초)</param>
    /// <returns>생성된 DOCX 파일의 전체 경로</returns>
    public static async Task<string> ConvertToDocxAsync(
        string inputPath,
        string outDir,
        string? sOfficePath = null,
        IProgress<(int Percent, string Message)>? progress = null,
        int timeoutMs = 120_000)
    {
        var sofficeBin = sOfficePath ?? FindSOffice()
            ?? throw new InvalidOperationException(
                "LibreOffice 를 찾을 수 없습니다. LibreOffice 를 설치하거나 " +
                "LIBREOFFICE_PATH 환경변수를 설정하세요.");

        Directory.CreateDirectory(outDir);

        progress?.Report((5, $"LibreOffice 실행 중…"));

        // soffice --headless --norestore --env:UserInstallation=... --convert-to docx --outdir <dir> <input>
        var psi = BuildProcessStartInfo(sofficeBin, outDir);
        psi.ArgumentList.Add("--convert-to");
        psi.ArgumentList.Add("docx:MS Word 2007 XML");
        psi.ArgumentList.Add("--outdir");
        psi.ArgumentList.Add(outDir);
        psi.ArgumentList.Add(Path.GetFullPath(inputPath));

        using var proc = Process.Start(psi)
            ?? throw new InvalidOperationException($"soffice 실행 실패: {sofficeBin}");

        var stdoutTask = proc.StandardOutput.ReadToEndAsync();
        var stderrTask = proc.StandardError.ReadToEndAsync();

        using var cts = new System.Threading.CancellationTokenSource(timeoutMs);
        try
        {
            await proc.WaitForExitAsync(cts.Token).ConfigureAwait(false);
        }
        catch (OperationCanceledException)
        {
            try { proc.Kill(entireProcessTree: true); } catch { }
            throw new InvalidOperationException(
                $"LibreOffice 변환 타임아웃 ({timeoutMs}ms 초과)");
        }

        var stdout = await stdoutTask.ConfigureAwait(false);
        var stderr = await stderrTask.ConfigureAwait(false);

        if (proc.ExitCode != 0)
            throw new InvalidOperationException(
                $"LibreOffice 변환 실패 (종료 코드 {proc.ExitCode})\n" +
                $"stdout: {stdout.Trim()}\nstderr: {stderr.Trim()}");

        // LibreOffice 는 출력 파일을 <outDir>/<원본이름>.docx 로 저장
        var baseName   = Path.GetFileNameWithoutExtension(inputPath);
        var docxOutput = Path.Combine(outDir, baseName + ".docx");

        if (!File.Exists(docxOutput))
            throw new InvalidOperationException(
                $"LibreOffice 변환 후 예상 출력 파일 없음: {docxOutput}\n" +
                $"stdout: {stdout.Trim()}");

        progress?.Report((40, "LibreOffice 변환 완료"));
        return docxOutput;
    }

    /// <summary>
    /// LibreOffice headless 로 DOCX 를 DOC 또는 HWP 형식으로 변환.
    /// </summary>
    /// <param name="inputDocxPath">변환할 DOCX 파일 경로</param>
    /// <param name="outDir">출력 디렉터리</param>
    /// <param name="targetExt">목표 확장자 ("doc" 또는 "hwp")</param>
    /// <param name="sOfficePath">soffice 실행 파일 경로 (null이면 자동 탐지)</param>
    /// <param name="progress">진행률 보고</param>
    /// <param name="timeoutMs">타임아웃 밀리초</param>
    /// <returns>생성된 출력 파일의 전체 경로</returns>
    public static async Task<string> ConvertFromDocxAsync(
        string inputDocxPath,
        string outDir,
        string targetExt,
        string? sOfficePath = null,
        IProgress<(int Percent, string Message)>? progress = null,
        int timeoutMs = 120_000)
    {
        var sofficeBin = sOfficePath ?? FindSOffice()
            ?? throw new InvalidOperationException(
                "LibreOffice 를 찾을 수 없습니다.");

        Directory.CreateDirectory(outDir);

        progress?.Report((70, "LibreOffice 로 최종 포맷 변환 중…"));

        // ArgumentList 사용 시 필터 이름에 따옴표를 포함하지 않는다.
        var filterSpec = targetExt.ToLowerInvariant() switch
        {
            "doc" => "doc:MS Word 97",
            "hwp" => "hwp",
            _ => throw new ArgumentException($"지원하지 않는 출력 포맷: {targetExt}"),
        };

        var psi = BuildProcessStartInfo(sofficeBin, outDir);
        psi.ArgumentList.Add("--convert-to");
        psi.ArgumentList.Add(filterSpec);
        psi.ArgumentList.Add("--outdir");
        psi.ArgumentList.Add(outDir);
        psi.ArgumentList.Add(Path.GetFullPath(inputDocxPath));

        using var proc = Process.Start(psi)
            ?? throw new InvalidOperationException($"soffice 실행 실패: {sofficeBin}");

        var stdoutTask = proc.StandardOutput.ReadToEndAsync();
        var stderrTask = proc.StandardError.ReadToEndAsync();

        using var cts = new System.Threading.CancellationTokenSource(timeoutMs);
        try
        {
            await proc.WaitForExitAsync(cts.Token).ConfigureAwait(false);
        }
        catch (OperationCanceledException)
        {
            try { proc.Kill(entireProcessTree: true); } catch { }
            throw new InvalidOperationException(
                $"LibreOffice 변환 타임아웃 ({timeoutMs}ms 초과)");
        }

        var stdout = await stdoutTask.ConfigureAwait(false);
        var stderr = await stderrTask.ConfigureAwait(false);

        if (proc.ExitCode != 0)
            throw new InvalidOperationException(
                $"LibreOffice 변환 실패 (종료 코드 {proc.ExitCode})\n" +
                $"stderr: {stderr.Trim()}");

        var baseName  = Path.GetFileNameWithoutExtension(inputDocxPath);
        var outFile   = Path.Combine(outDir, baseName + "." + targetExt.ToLowerInvariant());
        if (!File.Exists(outFile))
            throw new InvalidOperationException(
                $"LibreOffice 변환 후 예상 출력 파일 없음: {outFile}");

        progress?.Report((95, "LibreOffice 변환 완료"));
        return outFile;
    }

    private static ProcessStartInfo BuildProcessStartInfo(string sofficeBin, string workDir)
    {
        var programDir = Path.GetDirectoryName(sofficeBin) ?? "";

        // Isolated user profile prevents lock-file conflicts when multiple conversions run.
        var userInstallDir = Path.Combine(workDir, "lo-user-" + Guid.NewGuid().ToString("N")[..8]);
        var userInstallUri = "file:///" + userInstallDir.Replace('\\', '/');

        var psi = new ProcessStartInfo
        {
            FileName               = sofficeBin,
            WorkingDirectory       = programDir,
            UseShellExecute        = false,
            RedirectStandardOutput = true,
            RedirectStandardError  = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding  = Encoding.UTF8,
            CreateNoWindow         = true,
        };

        // Prepend LibreOffice's program dir so its Python runtime finds platform libraries.
        if (!string.IsNullOrEmpty(programDir))
        {
            var existing = Environment.GetEnvironmentVariable("PATH") ?? "";
            if (!existing.Contains(programDir, StringComparison.OrdinalIgnoreCase))
                psi.Environment["PATH"] = programDir + Path.PathSeparator + existing;
        }

        psi.ArgumentList.Add("--headless");
        psi.ArgumentList.Add("--norestore");
        psi.ArgumentList.Add($"-env:UserInstallation={userInstallUri}");
        return psi;
    }
}

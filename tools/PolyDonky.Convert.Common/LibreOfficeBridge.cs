using System;
using System.Diagnostics;
using System.IO;
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
        var envPath = Environment.GetEnvironmentVariable("LIBREOFFICE_PATH");
        if (!string.IsNullOrEmpty(envPath))
        {
            string[] envCandidates =
            [
                Path.Combine(envPath, "program", "soffice.exe"),
                Path.Combine(envPath, "program", "soffice"),
                Path.Combine(envPath, "bin",     "soffice.exe"),
                Path.Combine(envPath, "bin",     "soffice"),
                Path.Combine(envPath,            "soffice.exe"),
                Path.Combine(envPath,            "soffice"),
            ];
            foreach (var c in envCandidates)
                if (File.Exists(c)) return c;
        }

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
    /// 출력 DOCX 경로를 반환 (outDir 아래).
    /// </summary>
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

        progress?.Report((5, "LibreOffice 실행 중…"));

        // Non-ASCII filenames can confuse LibreOffice's arg parsing; copy to a safe name.
        var safeExt   = Path.GetExtension(inputPath).ToLowerInvariant();
        var safeInput = Path.Combine(outDir, "lo_input" + safeExt);
        File.Copy(inputPath, safeInput, overwrite: true);

        var psi = BuildProcessStartInfo(sofficeBin, outDir);
        psi.ArgumentList.Add("--convert-to");
        psi.ArgumentList.Add("docx:MS Word 2007 XML");
        psi.ArgumentList.Add("--outdir");
        psi.ArgumentList.Add(outDir);
        psi.ArgumentList.Add(safeInput);

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

        var docxOutput = Path.Combine(outDir, "lo_input.docx");

        if (!File.Exists(docxOutput))
        {
            var existing = string.Join(", ", Directory.GetFiles(outDir).Select(Path.GetFileName));
            throw new InvalidOperationException(
                $"LibreOffice 변환 후 예상 출력 파일 없음: {docxOutput}\n" +
                $"outDir 내 파일: [{existing}]\n" +
                $"stdout: {stdout.Trim()}");
        }

        progress?.Report((40, "LibreOffice 변환 완료"));
        return docxOutput;
    }

    /// <summary>
    /// LibreOffice headless 로 DOCX 를 DOC 또는 HWP 형식으로 변환.
    /// </summary>
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

        var baseName = Path.GetFileNameWithoutExtension(inputDocxPath);
        var outFile  = Path.Combine(outDir, baseName + "." + targetExt.ToLowerInvariant());
        if (!File.Exists(outFile))
            throw new InvalidOperationException(
                $"LibreOffice 변환 후 예상 출력 파일 없음: {outFile}");

        progress?.Report((95, "LibreOffice 변환 완료"));
        return outFile;
    }

    private static ProcessStartInfo BuildProcessStartInfo(string sofficeBin, string workDir)
    {
        var programDir = Path.GetDirectoryName(sofficeBin) ?? "";

        // Create isolated temp profile so LibreOffice doesn't depend on user's profile directory.
        var userInstallDir = Path.Combine(workDir, "lo-user-" + Guid.NewGuid().ToString("N")[..8]);
        Directory.CreateDirectory(userInstallDir);
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

        // Prepend LibreOffice's program dir to PATH, but remove user Python paths that might conflict.
        if (!string.IsNullOrEmpty(programDir))
        {
            var existing = Environment.GetEnvironmentVariable("PATH") ?? "";
            var pathParts = existing.Split(Path.PathSeparator);

            // Filter out Python-related paths (user installations, virtual envs, etc) that conflict
            // with LibreOffice's bundled Python.
            var filtered = pathParts
                .Where(p => !p.Contains("Python", StringComparison.OrdinalIgnoreCase))
                .ToList();

            // Prepend LibreOffice program dir
            filtered.Insert(0, programDir);

            psi.Environment["PATH"] = string.Join(Path.PathSeparator.ToString(), filtered);
        }

        // LibreOffice bundles its own Python (e.g. program\python-core-3.x.y\).
        // Set PYTHONHOME to that directory so Python finds its stdlib.
        if (!string.IsNullOrEmpty(programDir))
        {
            var coreDirs = Directory.GetDirectories(programDir, "python-core-*");
            if (coreDirs.Length > 0)
            {
                psi.Environment["PYTHONHOME"] = coreDirs[0];
                psi.Environment["PYTHONPATH"] = "";
            }
            else
            {
                psi.Environment.Remove("PYTHONHOME");
                psi.Environment.Remove("PYTHONPATH");
            }
        }
        psi.Environment.Remove("PYTHONSTARTUP");

        psi.ArgumentList.Add("--headless");
        psi.ArgumentList.Add("--norestore");
        psi.ArgumentList.Add($"-env:UserInstallation={userInstallUri}");
        return psi;
    }
}

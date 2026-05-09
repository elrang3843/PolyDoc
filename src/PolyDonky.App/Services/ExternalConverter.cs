using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace PolyDonky.App.Services;

/// <summary>
/// CLAUDE.md §3 의 외부 변환 모듈 분리 원칙 — IWPF/MD/TXT 가 아닌 포맷은 별도 CLI 실행 파일을
/// spawn 해 IWPF 로 변환한 뒤 메인 앱이 읽는다(역방향: 저장 시 메인 앱이 IWPF 임시파일을 만들고
/// CLI 가 대상 포맷으로 변환).
///
/// 등록된 변환기:
///   .html / .htm    — PolyDonky.Convert.Html
///   .xml  / .xhtml  — PolyDonky.Convert.Xml
///   .docx           — PolyDonky.Convert.Docx
///   .hwpx           — PolyDonky.Convert.Hwpx
/// </summary>
public static class ExternalConverter
{
    /// <summary>이 확장자가 외부 CLI 변환을 거쳐야 하는지 검사.</summary>
    public static bool IsExternalFormat(string path) => GetConverter(path) is not null;

    /// <summary>지정 확장자에 대응하는 CLI 실행 파일 절대 경로 (없으면 null).</summary>
    public static string? GetConverter(string path)
    {
        var ext = Path.GetExtension(path).TrimStart('.').ToLowerInvariant();
        var name = ext switch
        {
            "html" or "htm"   => "PolyDonky.Convert.Html",
            "xml"  or "xhtml" => "PolyDonky.Convert.Xml",
            "docx"            => "PolyDonky.Convert.Docx",
            "hwpx"            => "PolyDonky.Convert.Hwpx",
            _                 => null,
        };
        if (name is null) return null;

        var baseDir = AppContext.BaseDirectory;
        // 같은 디렉터리에 .exe 또는 .dll(framework-dependent) 가 배치된다고 가정 — 빌드 타깃이 복사함.
        foreach (var candidate in new[]
        {
            Path.Combine(baseDir, name + ".exe"),
            Path.Combine(baseDir, name + ".dll"),
        })
        {
            if (File.Exists(candidate)) return candidate;
        }
        return null;
    }

    /// <summary>
    /// CLI 변환기를 spawn 해 input → output 변환을 수행한다. 백그라운드 스레드에서 호출 권장.
    /// 실패 시 Exception throw — 호출측이 ReportError 등으로 처리.
    /// </summary>
    /// <param name="progress">
    /// CLI 가 stdout 으로 보내는 <c>PROGRESS:&lt;percent&gt;:&lt;message&gt;</c> 줄을 파싱해
    /// (0~100, 메시지) 로 보고. null 이면 진행 정보 무시.
    /// </param>
    public static async Task ConvertAsync(
        string converterPath,
        string inputPath,
        string outputPath,
        IProgress<(int Percent, string Message)>? progress = null)
    {
        ArgumentNullException.ThrowIfNull(converterPath);
        ArgumentNullException.ThrowIfNull(inputPath);
        ArgumentNullException.ThrowIfNull(outputPath);

        bool isDll = converterPath.EndsWith(".dll", StringComparison.OrdinalIgnoreCase);
        var psi = new ProcessStartInfo
        {
            FileName  = isDll ? "dotnet" : converterPath,
            UseShellExecute        = false,
            RedirectStandardOutput = true,
            RedirectStandardError  = true,
            // CLI 가 Console.OutputEncoding=UTF-8 으로 출력하므로 부모도 UTF-8 로 디코딩해야 한다.
            // 미설정 시 Windows cp949 환경에서 한국어 메시지가 mojibake 됨.
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding  = Encoding.UTF8,
            CreateNoWindow         = true,
        };
        if (isDll) psi.ArgumentList.Add(converterPath);
        psi.ArgumentList.Add(inputPath);
        psi.ArgumentList.Add(outputPath);

        using var proc = Process.Start(psi)
            ?? throw new InvalidOperationException($"외부 변환기 실행 실패: {converterPath}");

        // stdout 을 한 줄씩 읽어 PROGRESS: 라인을 파싱하고 나머지는 흡수.
        var stdoutTask = Task.Run(async () =>
        {
            string? line;
            while ((line = await proc.StandardOutput.ReadLineAsync().ConfigureAwait(false)) != null)
            {
                if (progress is null) continue;
                if (!line.StartsWith("PROGRESS:", StringComparison.Ordinal)) continue;

                // PROGRESS:<percent>:<message>
                var rest  = line.AsSpan(9);
                int colon = rest.IndexOf(':');
                if (colon < 0) continue;
                if (!int.TryParse(rest[..colon], out int percent)) continue;
                // CLI 측 버그로 0~100 범위 밖 값이 와도 ProgressBar 오버플로 안 되게 클램프.
                percent = Math.Clamp(percent, 0, 100);
                progress.Report((percent, rest[(colon + 1)..].ToString()));
            }
        });

        var stderrTask = proc.StandardError.ReadToEndAsync();
        await proc.WaitForExitAsync().ConfigureAwait(false);
        await stdoutTask.ConfigureAwait(false);
        var stderr = await stderrTask.ConfigureAwait(false);

        if (proc.ExitCode == ExitCodeUnsupportedVersion)
        {
            // 6 = 지원하지 않는 옛 버전. 호출측이 별도로 안내하도록 전용 예외.
            throw new UnsupportedFormatVersionException(stderr.Trim());
        }
        if (proc.ExitCode != 0)
        {
            throw new InvalidOperationException(
                $"외부 변환 실패 (종료 코드 {proc.ExitCode}): {converterPath}\n{stderr}");
        }
    }

    /// <summary>CLI 가 "지원하지 않는 옛 버전" 으로 거부할 때의 종료 코드.</summary>
    public const int ExitCodeUnsupportedVersion = 6;

    /// <summary>임시 IWPF 파일 경로 — Path.GetTempPath() 아래 고유한 이름.</summary>
    public static string CreateTempIwpfPath()
        => Path.Combine(Path.GetTempPath(), $"polydonky-{Guid.NewGuid():N}.iwpf");
}

/// <summary>외부 CLI 변환기가 지원 범위 밖 버전(예: HWPX 1.0, Word 2010 DOCX) 을 거부했을 때 던진다.</summary>
public sealed class UnsupportedFormatVersionException : Exception
{
    public UnsupportedFormatVersionException(string detail) : base(detail) { }
}

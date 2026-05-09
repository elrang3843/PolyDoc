using System.IO.Compression;
using System.Xml.Linq;

namespace PolyDonky.Codecs.Hwpx;

/// <summary>
/// HWPX 패키지의 버전 정보를 추출해 지원 정책(2014 년 이후 = HWPX 1.2+) 을 적용한다.
///
/// HWPX 의 <c>version.xml</c> 에는 다음 속성이 있다:
///   - <c>xmlVersion</c>  : 스키마 버전 (예: "1.0" / "1.2" / "1.3" / "1.31")
///   - HCFVersion <c>major.minor.micro.buildNumber</c> : 한컴 응용프로그램 버전 (예: 5.1.1.0)
/// 한컴오피스 2014(=HWP 2014) 부터 xmlVersion 이 1.2 로 올라갔다. 그 이전(1.0, 1.1) 은
/// 2010~2013 시기의 파일이며 본 codec 의 매핑이 검증되지 않으므로 거부한다.
/// </summary>
public static class HwpxVersionPolicy
{
    /// <summary>최소 지원 HWPX 스키마 버전 (HWP 2014 = HWPX 1.2 이상).</summary>
    public const string MinSupportedXmlVersion = "1.2";

    /// <summary>패키지의 xmlVersion 을 추출 — 없으면 null. 잘못된 ZIP 이어도 예외 던지지 않음.</summary>
    public static string? ReadXmlVersion(string path)
    {
        try
        {
            using var fs = File.OpenRead(path);
            using var zip = new ZipArchive(fs, ZipArchiveMode.Read, leaveOpen: false);
            var entry = zip.GetEntry("version.xml");
            if (entry is null) return null;
            using var es = entry.Open();
            var xdoc = XDocument.Load(es);
            var root = xdoc.Root;
            return root?.Attribute("xmlVersion")?.Value;
        }
        catch { return null; }
    }

    /// <summary>지원 가능 여부. 버전 정보가 없으면 보수적으로 수락(메타데이터 누락 호환).</summary>
    public static bool IsSupported(string? xmlVersion)
    {
        if (string.IsNullOrWhiteSpace(xmlVersion)) return true;
        return CompareVersion(xmlVersion, MinSupportedXmlVersion) >= 0;
    }

    /// <summary>"1.2.3" 형식 문자열 비교 — major/minor/patch 만 본다.</summary>
    private static int CompareVersion(string a, string b)
    {
        var aParts = a.Split('.');
        var bParts = b.Split('.');
        int n = Math.Max(aParts.Length, bParts.Length);
        for (int i = 0; i < n; i++)
        {
            int ai = i < aParts.Length && int.TryParse(aParts[i], out var av) ? av : 0;
            int bi = i < bParts.Length && int.TryParse(bParts[i], out var bv) ? bv : 0;
            if (ai != bi) return ai.CompareTo(bi);
        }
        return 0;
    }
}

using System.IO.Compression;
using System.Xml.Linq;

namespace PolyDonky.Codecs.Docx;

/// <summary>
/// DOCX 의 버전 정보를 추출해 지원 정책(2012년 이후 = Word 2013 / AppVersion 15.0+) 을 적용한다.
///
/// DOCX 의 <c>docProps/app.xml</c> 에 <c>&lt;AppVersion&gt;</c> 요소가 있다(예: "15.0000" = Word 2013).
/// Word 2010 = 14.0000, 2007 = 12.0000. 2013(=15.0000) 미만은 본 codec 의 호환 검증 범위
/// 밖이므로 거부. 메타데이터 자체가 없는 DOCX(LibreOffice 등) 는 보수적으로 수락한다.
/// </summary>
public static class DocxVersionPolicy
{
    /// <summary>최소 지원 Word AppVersion (Word 2013 = 15.0).</summary>
    public const double MinSupportedAppVersion = 15.0;

    /// <summary>패키지에서 AppVersion 숫자 추출 — 없으면 null.</summary>
    public static double? ReadAppVersion(string path)
    {
        try
        {
            using var fs = File.OpenRead(path);
            using var zip = new ZipArchive(fs, ZipArchiveMode.Read, leaveOpen: false);
            var entry = zip.GetEntry("docProps/app.xml");
            if (entry is null) return null;

            using var es   = entry.Open();
            var xdoc = XDocument.Load(es);
            var root = xdoc.Root;
            if (root is null) return null;

            // app.xml 의 namespace: extended-properties.
            var appVersionEl = root.Elements()
                .FirstOrDefault(e => e.Name.LocalName.Equals("AppVersion", StringComparison.OrdinalIgnoreCase));
            if (appVersionEl is null) return null;
            if (double.TryParse(appVersionEl.Value, System.Globalization.NumberStyles.Any,
                                System.Globalization.CultureInfo.InvariantCulture, out var v))
                return v;
            return null;
        }
        catch { return null; }
    }

    /// <summary>지원 가능 여부. AppVersion 미기재면 보수적으로 수락.</summary>
    public static bool IsSupported(double? appVersion)
    {
        if (appVersion is null) return true;
        return appVersion.Value >= MinSupportedAppVersion;
    }
}

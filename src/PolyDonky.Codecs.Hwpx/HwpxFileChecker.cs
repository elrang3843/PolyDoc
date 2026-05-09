using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

namespace PolyDonky.Codecs.Hwpx;

/// <summary>
/// HWPX 패키지 사전 검증 헬퍼 — reader 호출 전에 친절한 메시지로 거를 수 있는 케이스들.
///
/// 검사 항목:
///   1. <c>mimetype</c> 엔트리 존재·내용 (`application/hwp+zip`) — HWPX 사양상 ZIP 의 첫 엔트리.
///   2. 암호화 / DRM (META-INF/manifest.xml 의 <c>encryption-data</c> 또는 OWPML <c>encryption</c> 노드).
///   3. 핵심 콘텐츠(<c>Contents/header.xml</c>, <c>Contents/section0.xml</c>) 존재.
///   4. <c>version.xml</c> 의 application·appVersion 추출 (UI 안내용).
///
/// 모든 메서드는 예외를 던지지 않고 안전하게 false / null 을 반환한다.
/// </summary>
public static class HwpxFileChecker
{
    public const string ExpectedMimetype = "application/hwp+zip";

    public static bool HasValidMimetype(string path)
    {
        try
        {
            using var fs = File.OpenRead(path);
            using var zip = new ZipArchive(fs, ZipArchiveMode.Read, leaveOpen: false);

            var entry = zip.GetEntry("mimetype");
            if (entry is null)
            {
                // mimetype 엔트리 없어도 본문(header.xml + section*.xml)이 있으면 HWPX 로 간주.
                return HasCoreContentInternal(zip);
            }

            using var es = entry.Open();
            // UTF-8 BOM 자동 감지 — 일부 한컴 HWPX 가 BOM(EF BB BF)을 붙여 ASCII reader 가 비교 실패하는 문제 회피.
            using var sr = new StreamReader(es, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
            var content = sr.ReadToEnd().Trim().Trim('﻿');
            if (content == ExpectedMimetype) return true;
            // 변종 mimetype 이라도 본문 구조가 올바르면 통과.
            return HasCoreContentInternal(zip);
        }
        catch { return false; }
    }

    private static bool HasCoreContentInternal(ZipArchive zip)
    {
        bool hasHeader  = zip.Entries.Any(e =>
            e.FullName.Equals("Contents/header.xml", StringComparison.OrdinalIgnoreCase));
        bool hasSection = zip.Entries.Any(e =>
            e.FullName.StartsWith("Contents/section", StringComparison.OrdinalIgnoreCase) &&
            e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase));
        return hasHeader && hasSection;
    }

    /// <summary>HWPX 가 암호화·DRM 으로 잠겨 있으면 true.</summary>
    public static bool IsEncrypted(string path)
    {
        try
        {
            using var fs = File.OpenRead(path);
            using var zip = new ZipArchive(fs, ZipArchiveMode.Read, leaveOpen: false);

            var manifest = zip.GetEntry("META-INF/manifest.xml")
                        ?? zip.GetEntry("META-INF/Manifest.xml");
            if (manifest is null) return false;

            using var es = manifest.Open();
            var xdoc = XDocument.Load(es);
            // OWPML / OPC 표기 모두 검사.
            return xdoc.Descendants().Any(e =>
                e.Name.LocalName.Equals("encryption-data", StringComparison.OrdinalIgnoreCase) ||
                e.Name.LocalName.Equals("encryption",      StringComparison.OrdinalIgnoreCase));
        }
        catch { return false; }
    }

    /// <summary>핵심 콘텐츠 엔트리(header.xml + 적어도 1개 section)의 존재 여부.</summary>
    public static bool HasCoreContent(string path)
    {
        try
        {
            using var fs = File.OpenRead(path);
            using var zip = new ZipArchive(fs, ZipArchiveMode.Read, leaveOpen: false);
            bool hasHeader  = zip.GetEntry("Contents/header.xml") is not null;
            bool hasSection = zip.Entries.Any(e =>
                e.FullName.StartsWith("Contents/section", StringComparison.Ordinal) &&
                e.FullName.EndsWith(".xml", StringComparison.Ordinal));
            return hasHeader && hasSection;
        }
        catch { return false; }
    }

    /// <summary>
    /// version.xml 에서 작성 앱 정보를 추출. (application, appVersion). 둘 중 하나라도 없으면 null.
    /// 예: ("HwpApp", "9, 0, 0, 1234")
    /// </summary>
    public static (string? application, string? appVersion) ReadAppInfo(string path)
    {
        try
        {
            using var fs = File.OpenRead(path);
            using var zip = new ZipArchive(fs, ZipArchiveMode.Read, leaveOpen: false);
            var entry = zip.GetEntry("version.xml");
            if (entry is null) return (null, null);
            using var es = entry.Open();
            var xdoc = XDocument.Load(es);
            var root = xdoc.Root;
            return (root?.Attribute("application")?.Value,
                    root?.Attribute("appVersion")?.Value);
        }
        catch { return (null, null); }
    }
}

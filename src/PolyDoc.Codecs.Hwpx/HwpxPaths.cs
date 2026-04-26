namespace PolyDoc.Codecs.Hwpx;

/// <summary>
/// HWPX (KS X 6101) 패키지 표준 상수.
/// HWPX 는 OPC/EPUB 와 유사한 ZIP 컨테이너로, 첫 항목이 압축 없이 저장된 mimetype 이어야 한다.
/// </summary>
public static class HwpxPaths
{
    public const string Mimetype = "mimetype";
    public const string MimetypeContent = "application/hwp+zip";

    public const string ContainerXml = "META-INF/container.xml";
    public const string ManifestXml = "META-INF/manifest.xml";

    public const string ContentDir = "Contents/";
    public const string ContentHpf = "Contents/content.hpf";
    public const string HeaderXml = "Contents/header.xml";

    /// <summary>섹션 본문은 section0.xml, section1.xml ... 으로 이어진다.</summary>
    public static string SectionXml(int index) => $"Contents/section{index}.xml";

    public const string SettingsXml = "settings.xml";
    public const string VersionXml = "version.xml";

    public const string BinDataDir = "BinData/";
    public const string PreviewDir = "Preview/";
}

/// <summary>
/// HWPX 의 namespace URI. KS X 6101 표준 / 한컴 hwpml 2011 계열에서 정의된 값을 사용한다.
/// 한컴 오피스가 만든 hwpx 와 호환되도록 prefix 도 관용적인 hh / hp / hs / hc 로 둔다.
/// </summary>
public static class HwpxNamespaces
{
    public const string OpfContainer = "urn:oasis:names:tc:opendocument:xmlns:container";

    public const string Hancom = "http://www.hancom.co.kr/hwpml/2011/";
    public const string Application = Hancom + "application";
    public const string Common = Hancom + "common";
    public const string Head = Hancom + "head";
    public const string Paragraph = Hancom + "paragraph";
    public const string Section = Hancom + "section";
    public const string Hwpml = Hancom + "hwpml";
    public const string Pgcd = Hancom + "pagecontentdef";
    public const string Master = Hancom + "masterpage";

    public const string OpfPackage = "http://www.idpf.org/2007/opf/";
    public const string DcMetadata = "http://purl.org/dc/elements/1.1/";

    public const string PrefixHa = "ha";
    public const string PrefixHc = "hc";
    public const string PrefixHh = "hh";
    public const string PrefixHp = "hp";
    public const string PrefixHs = "hs";
    public const string PrefixHml = "hml";
}

/// <summary>본 codec 이 출력하는 HWPX 의 자체 식별 정보.</summary>
public static class HwpxFormat
{
    public const string Version = "1.31";          // KS X 6101:2017 기준 통상 표기
    public const string ProducedBy = "PolyDoc";
}

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using PolyDonky.Core;

namespace PolyDonky.Convert.Hwp;

/// <summary>
/// IWPF → HWP 변환기. 현재 단계: 텍스트만.
/// HWP는 한글 워드프로세서의 파일 포맷으로, 내부적으로 OLE 구조 또는 ZIP을 사용한다.
/// 이 구현은 단순화된 구조를 생성하여 한글 등 HWP 리더에서 읽을 수 있도록 한다.
/// </summary>
public class HwpWriter
{
    public void Write(PolyDonkyument doc, Stream output)
    {
        var zipStream = new MemoryStream();
        using (var zip = new ZipArchive(zipStream, ZipArchiveMode.Create, leaveOpen: true))
        {
            // HWP 파일은 여러 XML 파일들을 포함한 구조
            // 최소 구조: BinData, content.xml, styles.xml 등

            // 콘텐츠 추출
            var content = ExtractContent(doc);

            // content.xml 생성
            var contentXml = GenerateContentXml(content);
            var contentEntry = zip.CreateEntry("content.xml");
            using (var entryStream = contentEntry.Open())
            using (var writer = new StreamWriter(entryStream, Encoding.UTF8))
            {
                writer.Write(contentXml);
            }

            // styles.xml 생성
            var stylesXml = GenerateStylesXml();
            var stylesEntry = zip.CreateEntry("styles.xml");
            using (var entryStream = stylesEntry.Open())
            using (var writer = new StreamWriter(entryStream, Encoding.UTF8))
            {
                writer.Write(stylesXml);
            }

            // meta.xml 생성
            var metaXml = GenerateMetaXml();
            var metaEntry = zip.CreateEntry("meta.xml");
            using (var entryStream = metaEntry.Open())
            using (var writer = new StreamWriter(entryStream, Encoding.UTF8))
            {
                writer.Write(metaXml);
            }

            // settings.xml 생성
            var settingsXml = GenerateSettingsXml();
            var settingsEntry = zip.CreateEntry("settings.xml");
            using (var entryStream = settingsEntry.Open())
            using (var writer = new StreamWriter(entryStream, Encoding.UTF8))
            {
                writer.Write(settingsXml);
            }
        }

        zipStream.Seek(0, SeekOrigin.Begin);
        zipStream.CopyTo(output);
        zipStream.Dispose();
    }

    private DocumentContent ExtractContent(PolyDonkyument doc)
    {
        var content = new DocumentContent();
        content.Paragraphs = new List<ParagraphInfo>();

        foreach (var section in doc.Sections)
        {
            foreach (var block in section.Blocks)
            {
                if (block is Paragraph para)
                {
                    var parInfo = new ParagraphInfo();
                    parInfo.Text = string.Empty;
                    parInfo.Style = para.Style ?? new ParagraphStyle();

                    foreach (var run in para.Runs)
                    {
                        if (!string.IsNullOrEmpty(run.Text))
                        {
                            parInfo.Text += run.Text;
                        }
                    }

                    if (!string.IsNullOrEmpty(parInfo.Text))
                        content.Paragraphs.Add(parInfo);
                }
            }
        }

        return content;
    }

    private string GenerateContentXml(DocumentContent content)
    {
        var ns = XNamespace.Get("urn:schemas-microsoft-com:office:word");
        var doc = new XDocument(
            new XElement(ns + "document")
        );

        var body = new XElement(ns + "body");
        doc.Root?.Add(body);

        foreach (var para in content.Paragraphs)
        {
            var pElement = new XElement(ns + "p");
            pElement.Add(new XElement(ns + "r", para.Text));
            body.Add(pElement);
        }

        return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n" + doc.ToString(SaveOptions.DisableFormatting);
    }

    private string GenerateStylesXml()
    {
        var ns = XNamespace.Get("urn:schemas-microsoft-com:office:word");
        var doc = new XDocument(
            new XElement(ns + "styles")
        );
        return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n" + doc.ToString(SaveOptions.DisableFormatting);
    }

    private string GenerateMetaXml()
    {
        var ns = XNamespace.Get("urn:schemas-microsoft-com:office:meta");
        var doc = new XDocument(
            new XElement(ns + "document-properties")
        );
        return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n" + doc.ToString(SaveOptions.DisableFormatting);
    }

    private string GenerateSettingsXml()
    {
        var ns = XNamespace.Get("urn:schemas-microsoft-com:office:word");
        var doc = new XDocument(
            new XElement(ns + "settings")
        );
        return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n" + doc.ToString(SaveOptions.DisableFormatting);
    }

    private class DocumentContent
    {
        public List<ParagraphInfo> Paragraphs { get; set; } = new();
    }

    private class ParagraphInfo
    {
        public string Text { get; set; } = string.Empty;
        public ParagraphStyle Style { get; set; } = new();
    }
}

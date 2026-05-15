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
/// HWPX(ZIP 기반) 형식과 호환되는 구조를 생성한다.
/// </summary>
public class HwpWriter
{
    public void Write(PolyDonkyument doc, Stream output)
    {
        using (var zipStream = new MemoryStream())
        {
            using (var zip = new ZipArchive(zipStream, ZipArchiveMode.Create, leaveOpen: true))
            {
                // 콘텐츠 추출
                var paragraphs = ExtractContent(doc);

                // [Content_Types].xml
                AddEntry(zip, "[Content_Types].xml", GenerateContentTypes());

                // _rels/.rels
                zip.CreateEntry("_rels/");
                AddEntry(zip, "_rels/.rels", GenerateRels());

                // word/ directory
                zip.CreateEntry("word/");

                // word/document.xml
                AddEntry(zip, "word/document.xml", GenerateDocumentXml(paragraphs));

                // word/styles.xml
                AddEntry(zip, "word/styles.xml", GenerateStylesXml());

                // word/fontTable.xml
                AddEntry(zip, "word/fontTable.xml", GenerateFontTable());

                // word/numbering.xml
                AddEntry(zip, "word/numbering.xml", GenerateNumbering());

                // docProps/
                zip.CreateEntry("docProps/");

                // docProps/core.xml
                AddEntry(zip, "docProps/core.xml", GenerateCore());

                // docProps/app.xml
                AddEntry(zip, "docProps/app.xml", GenerateApp());
            }

            zipStream.Seek(0, SeekOrigin.Begin);
            zipStream.CopyTo(output);
            zipStream.Dispose();
        }
    }

    private List<string> ExtractContent(PolyDonkyument doc)
    {
        var paragraphs = new List<string>();

        foreach (var section in doc.Sections)
        {
            foreach (var block in section.Blocks)
            {
                if (block is Paragraph para)
                {
                    var text = string.Empty;
                    foreach (var run in para.Runs)
                    {
                        if (!string.IsNullOrEmpty(run.Text))
                            text += run.Text;
                    }

                    if (!string.IsNullOrEmpty(text))
                        paragraphs.Add(text);
                }
            }
        }

        return paragraphs;
    }

    private void AddEntry(ZipArchive zip, string path, string content)
    {
        var entry = zip.CreateEntry(path);
        using (var stream = entry.Open())
        using (var writer = new StreamWriter(stream, Encoding.UTF8))
        {
            writer.Write(content);
        }
    }

    private string GenerateContentTypes()
    {
        var ns = XNamespace.Get("http://schemas.openxmlformats.org/package/2006/content-types");
        var doc = new XDocument(
            new XElement(ns + "Types",
                new XElement(ns + "Default", new XAttribute("Extension", "rels"), new XAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml")),
                new XElement(ns + "Default", new XAttribute("Extension", "xml"), new XAttribute("ContentType", "application/xml")),
                new XElement(ns + "Override", new XAttribute("PartName", "/word/document.xml"), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")),
                new XElement(ns + "Override", new XAttribute("PartName", "/word/styles.xml"), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml")),
                new XElement(ns + "Override", new XAttribute("PartName", "/word/fontTable.xml"), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml")),
                new XElement(ns + "Override", new XAttribute("PartName", "/word/numbering.xml"), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml")),
                new XElement(ns + "Override", new XAttribute("PartName", "/docProps/core.xml"), new XAttribute("ContentType", "application/vnd.openxmlformats-package.core-properties+xml")),
                new XElement(ns + "Override", new XAttribute("PartName", "/docProps/app.xml"), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.extended-properties+xml"))
            )
        );
        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" + doc.ToString();
    }

    private string GenerateRels()
    {
        var ns = XNamespace.Get("http://schemas.openxmlformats.org/package/2006/relationships");
        var doc = new XDocument(
            new XElement(ns + "Relationships",
                new XElement(ns + "Relationship",
                    new XAttribute("Id", "rId1"),
                    new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"),
                    new XAttribute("Target", "word/document.xml")
                ),
                new XElement(ns + "Relationship",
                    new XAttribute("Id", "rId2"),
                    new XAttribute("Type", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"),
                    new XAttribute("Target", "docProps/core.xml")
                ),
                new XElement(ns + "Relationship",
                    new XAttribute("Id", "rId3"),
                    new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"),
                    new XAttribute("Target", "docProps/app.xml")
                )
            )
        );
        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" + doc.ToString();
    }

    private string GenerateDocumentXml(List<string> paragraphs)
    {
        var ns = XNamespace.Get("http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        var doc = new XDocument(
            new XElement(ns + "document",
                new XElement(ns + "body")
            )
        );

        var body = doc.Root?.Element(ns + "body");

        foreach (var text in paragraphs)
        {
            var p = new XElement(ns + "p");
            var pPr = new XElement(ns + "pPr");
            p.Add(pPr);

            var r = new XElement(ns + "r");
            var rPr = new XElement(ns + "rPr");
            r.Add(rPr);

            var t = new XElement(ns + "t");
            t.Value = text;
            r.Add(t);

            p.Add(r);
            body?.Add(p);
        }

        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" + doc.ToString();
    }

    private string GenerateStylesXml()
    {
        var ns = XNamespace.Get("http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        var doc = new XDocument(
            new XElement(ns + "styles")
        );
        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" + doc.ToString();
    }

    private string GenerateFontTable()
    {
        var ns = XNamespace.Get("http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        var doc = new XDocument(
            new XElement(ns + "fontTbl")
        );
        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" + doc.ToString();
    }

    private string GenerateNumbering()
    {
        var ns = XNamespace.Get("http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        var doc = new XDocument(
            new XElement(ns + "numbering")
        );
        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" + doc.ToString();
    }

    private string GenerateCore()
    {
        var ns = XNamespace.Get("http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
        var dcNs = XNamespace.Get("http://purl.org/dc/elements/1.1/");
        var dcTermsNs = XNamespace.Get("http://purl.org/dc/terms/");
        var doc = new XDocument(
            new XElement(ns + "coreProperties",
                new XElement(dcNs + "creator", "PolyDonky"),
                new XElement(dcTermsNs + "created", new XAttribute(XNamespace.Xmlns + "xsi", "http://www.w3.org/2001/XMLSchema-instance"), DateTime.UtcNow.ToString("O"))
            )
        );
        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" + doc.ToString();
    }

    private string GenerateApp()
    {
        var ns = XNamespace.Get("http://schemas.openxmlformats.org/officeDocument/2006/extended-properties");
        var doc = new XDocument(
            new XElement(ns + "Properties",
                new XElement(ns + "TotalTime", 0),
                new XElement(ns + "Pages", 1),
                new XElement(ns + "Words", 0),
                new XElement(ns + "Characters", 0),
                new XElement(ns + "Application", "PolyDonky"),
                new XElement(ns + "AppVersion", "1.0")
            )
        );
        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" + doc.ToString();
    }
}

using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using PolyDoc.Core;

namespace PolyDoc.Iwpf;

/// <summary>
/// PolyDocument 를 IWPF (ZIP+JSON) 패키지로 직렬화한다.
/// 매니페스트는 모든 파트의 SHA-256 해시를 포함하므로 본문 파트를 먼저
/// ZIP 에 기록한 뒤 마지막에 매니페스트를 추가한다.
///
/// ImageBlock 은 바이너리 inline 직렬화 대신 resources/images/ 로 분리해
/// 패키지 무게를 줄이고 매니페스트로 무결성을 검증한다. 같은 SHA-256 의
/// 이미지는 한 번만 저장하고 모든 ImageBlock 이 같은 ResourcePath 를 공유.
/// </summary>
public sealed class IwpfWriter : IDocumentWriter
{
    public string FormatId => "iwpf";

    public void Write(PolyDocument document, Stream output)
    {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(output);

        using var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true);

        var manifest = new IwpfManifest
        {
            Created = document.Metadata.Created,
            Modified = document.Metadata.Modified,
        };

        // 1. 이미지 자원 분리 저장. 같은 해시면 한 파트만 저장하고 ResourcePath 를 공유.
        var stashedImageBytes = SplitImagesAndAssignPaths(document, archive, manifest);

        try
        {
            // 2. content/document.json — ImageBlock.Data 가 비워진 상태로 직렬화.
            var documentBytes = SerializeDocument(document);
            AddPart(archive, manifest, IwpfPaths.DocumentJson, IwpfMediaTypes.Document, documentBytes);

            // 3. content/styles.json
            var stylesBytes = JsonSerializer.SerializeToUtf8Bytes(document.Styles, JsonDefaults.Options);
            AddPart(archive, manifest, IwpfPaths.StylesJson, IwpfMediaTypes.Styles, stylesBytes);

            // 4. provenance/source-map.json (있을 때만)
            if (document.Provenance.NodeAnchors.Count > 0)
            {
                var provenanceBytes = JsonSerializer.SerializeToUtf8Bytes(document.Provenance, JsonDefaults.Options);
                AddPart(archive, manifest, IwpfPaths.ProvenanceJson, IwpfMediaTypes.Provenance, provenanceBytes);
            }

            // 5. manifest.json — 항상 마지막. 본문 파트의 해시가 모두 결정된 뒤에 작성한다.
            var manifestBytes = JsonSerializer.SerializeToUtf8Bytes(manifest, JsonDefaults.Options);
            WriteEntry(archive, IwpfPaths.Manifest, manifestBytes);
        }
        finally
        {
            // 호출자가 같은 PolyDocument 를 계속 사용할 수 있도록 ImageBlock.Data 복원.
            foreach (var (image, bytes) in stashedImageBytes)
            {
                image.Data = bytes;
            }
        }
    }

    /// <summary>
    /// 문서 트리를 훑어 각 ImageBlock 에 ResourcePath 와 Sha256 을 채우고,
    /// dedupe 후 ZIP 에 분리 저장한다. ImageBlock.Data 는 직렬화 직전에 임시로
    /// 비워지며, 호출자에게 복원용 (block, original-bytes) 페어 목록을 돌려준다.
    /// </summary>
    private static List<(ImageBlock image, byte[] originalBytes)> SplitImagesAndAssignPaths(
        PolyDocument document,
        ZipArchive archive,
        IwpfManifest manifest)
    {
        var stashed = new List<(ImageBlock, byte[])>();
        var pathByHash = new Dictionary<string, string>(StringComparer.Ordinal);
        int counter = 0;

        foreach (var section in document.Sections)
        {
            Walk(section.Blocks);
        }
        return stashed;

        void Walk(IList<Block> blocks)
        {
            foreach (var block in blocks)
            {
                switch (block)
                {
                    case ImageBlock image when image.Data.Length > 0:
                        AssignAndWrite(image);
                        break;
                    case Table table:
                        foreach (var row in table.Rows)
                            foreach (var cell in row.Cells)
                                Walk(cell.Blocks);
                        break;
                }
            }
        }

        void AssignAndWrite(ImageBlock image)
        {
            var hash = string.IsNullOrEmpty(image.Sha256)
                ? Sha256Hex(image.Data)
                : image.Sha256;
            image.Sha256 = hash;

            if (!pathByHash.TryGetValue(hash, out var path))
            {
                counter++;
                path = $"{IwpfPaths.ImagesDir}img-{counter:D4}{ExtensionForMediaType(image.MediaType)}";
                pathByHash[hash] = path;
                AddPart(archive, manifest, path, image.MediaType, image.Data);
            }

            image.ResourcePath = path;
            stashed.Add((image, image.Data));
            image.Data = Array.Empty<byte>();
        }
    }

    private static string ExtensionForMediaType(string mediaType) => mediaType switch
    {
        "image/png" => ".png",
        "image/jpeg" => ".jpg",
        "image/gif" => ".gif",
        "image/bmp" => ".bmp",
        "image/tiff" => ".tif",
        "image/svg+xml" => ".svg",
        "image/webp" => ".webp",
        _ => ".bin",
    };

    private static byte[] SerializeDocument(PolyDocument document)
    {
        // 패키지에서는 styles 와 provenance 를 별도 파트로 분리해 저장하므로
        // document.json 본문에는 그 둘을 비워서 직렬화한다.
        var detached = new PolyDocument
        {
            Metadata = document.Metadata,
            Sections = document.Sections,
            Styles = new StyleSheet(),
            Provenance = new Provenance(),
        };
        return JsonSerializer.SerializeToUtf8Bytes(detached, JsonDefaults.Options);
    }

    private static void AddPart(ZipArchive archive, IwpfManifest manifest, string path, string mediaType, byte[] payload)
    {
        WriteEntry(archive, path, payload);
        manifest.Parts[path] = new IwpfManifestEntry
        {
            Path = path,
            MediaType = mediaType,
            Size = payload.LongLength,
            Sha256 = Sha256Hex(payload),
        };
    }

    private static void WriteEntry(ZipArchive archive, string path, byte[] payload)
    {
        var entry = archive.CreateEntry(path, CompressionLevel.Optimal);
        using var entryStream = entry.Open();
        entryStream.Write(payload, 0, payload.Length);
    }

    private static string Sha256Hex(ReadOnlySpan<byte> data)
    {
        Span<byte> hash = stackalloc byte[SHA256.HashSizeInBytes];
        SHA256.HashData(data, hash);
        return Convert.ToHexStringLower(hash);
    }

    /// <summary>편의 메서드 — 텍스트 디버깅용으로 매니페스트만 별도 직렬화.</summary>
    public static string SerializeManifest(IwpfManifest manifest)
        => Encoding.UTF8.GetString(JsonSerializer.SerializeToUtf8Bytes(manifest, JsonDefaults.Options));
}

using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using PolyDonky.Core;

namespace PolyDonky.Iwpf;

/// <summary>
/// PolyDonkyument 를 IWPF (ZIP+JSON) 패키지로 직렬화한다.
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

    /// <summary>암호 보호 모드. 기본값은 <see cref="PasswordMode.None"/> (평문).</summary>
    public PasswordMode PasswordMode { get; init; }

    /// <summary>
    /// Read / Both 모드에서 AES-256-GCM 암호화 키 유도에 사용할 비밀번호(열기 암호).
    /// Write 전용 모드에서는 사용하지 않는다.
    /// </summary>
    public string? Password { get; init; }

    /// <summary>
    /// Write / Both 모드에서 쓰기 잠금 PBKDF2 해시 생성에 사용할 비밀번호(쓰기 암호).
    /// null 이면 <see cref="Password"/> 를 대신 사용한다(Both 동일 암호 시).
    /// </summary>
    public string? WritePassword { get; init; }

    public void Write(PolyDonkyument document, Stream output)
    {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(output);

        switch (PasswordMode)
        {
            case PasswordMode.Read:
            case PasswordMode.Both:
                if (string.IsNullOrEmpty(Password))
                    throw new InvalidOperationException(
                        "Password is required for Read or Both protection mode.");
                WriteEncrypted(document, output, Password, PasswordMode);
                break;

            case PasswordMode.Write:
                WritePlain(document, output, writeLockPassword: WritePassword ?? Password);
                break;

            default: // None
                WritePlain(document, output);
                break;
        }
    }

    private void WritePlain(PolyDonkyument document, Stream output, string? writeLockPassword = null)
    {
        using var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true);

        var manifest = new IwpfManifest
        {
            Created  = document.Metadata.Created,
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

            // 5. security/write-lock.json — 쓰기 보호가 설정된 경우.
            if (!string.IsNullOrEmpty(writeLockPassword))
            {
                manifest.WriteLocked = true;
                var writeLock      = IwpfEncryption.CreateWriteLock(writeLockPassword);
                var writeLockBytes = JsonSerializer.SerializeToUtf8Bytes(writeLock, JsonDefaults.Options);
                AddPart(archive, manifest, IwpfPaths.SecurityWriteLock, IwpfMediaTypes.Json, writeLockBytes);
            }

            // 6. manifest.json — 항상 마지막. 본문 파트의 해시가 모두 결정된 뒤에 작성한다.
            var manifestBytes = JsonSerializer.SerializeToUtf8Bytes(manifest, JsonDefaults.Options);
            WriteEntry(archive, IwpfPaths.Manifest, manifestBytes);
        }
        finally
        {
            // 호출자가 같은 PolyDonkyument 를 계속 사용할 수 있도록 ImageBlock.Data 복원.
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
        PolyDonkyument document,
        ZipArchive archive,
        IwpfManifest manifest)
    {
        var stashed    = new List<(ImageBlock, byte[])>();
        var pathByHash = new Dictionary<string, string>(StringComparer.Ordinal);
        int counter    = 0;

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
                    case TextBoxObject textbox:
                        // 글상자 안의 이미지/표/(중첩 글상자) 도 동일한 분리 저장 경로로 처리.
                        Walk(textbox.Content);
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
        "image/png"     => ".png",
        "image/jpeg"    => ".jpg",
        "image/gif"     => ".gif",
        "image/bmp"     => ".bmp",
        "image/tiff"    => ".tif",
        "image/svg+xml" => ".svg",
        "image/webp"    => ".webp",
        _               => ".bin",
    };

    private static byte[] SerializeDocument(PolyDonkyument document)
    {
        // 패키지에서는 styles 와 provenance 를 별도 파트로 분리해 저장하므로
        // document.json 본문에는 그 둘을 비워서 직렬화한다.
        // Watermark 는 본문 일부로 함께 보존된다.
        var detached = new PolyDonkyument
        {
            Metadata   = document.Metadata,
            Sections   = document.Sections,
            Styles     = new StyleSheet(),
            Provenance = new Provenance(),
            Watermark  = document.Watermark,
        };
        return JsonSerializer.SerializeToUtf8Bytes(detached, JsonDefaults.Options);
    }

    private static void AddPart(ZipArchive archive, IwpfManifest manifest, string path, string mediaType, byte[] payload)
    {
        WriteEntry(archive, path, payload);
        manifest.Parts[path] = new IwpfManifestEntry
        {
            Path      = path,
            MediaType = mediaType,
            Size      = payload.LongLength,
            Sha256    = Sha256Hex(payload),
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

    /// <summary>
    /// 암호화 모드(Read / Both). 평문 IWPF ZIP 을 메모리 내 빌드한 뒤 AES-256-GCM 으로 봉인하고
    /// outer ZIP 에는 manifest.json (encrypted=true), security/envelope.json,
    /// security/payload.bin 만 담는다.
    /// Both 모드에서는 inner ZIP 안에 쓰기 잠금(write-lock.json) 도 함께 포함한다.
    /// </summary>
    private void WriteEncrypted(PolyDonkyument document, Stream output, string password, PasswordMode mode)
    {
        // 1. inner IWPF ZIP 을 메모리에 평문으로 빌드.
        //    Both 모드면 write-lock 도 inner 에 포함한다.
        using var innerStream = new MemoryStream();
        var writeLockPwd = (mode == PasswordMode.Both) ? (WritePassword ?? password) : null;
        WritePlain(document, innerStream, writeLockPassword: writeLockPwd);
        var innerBytes = innerStream.ToArray();

        // 2. AES-256-GCM 암호화.
        var (cipherText, envelope) = IwpfEncryption.Encrypt(innerBytes, password);

        // 3. outer ZIP 작성: encrypted=true 매니페스트 + envelope.json + payload.bin.
        using var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true);

        var outerManifest = new IwpfManifest
        {
            Created   = document.Metadata.Created,
            Modified  = document.Metadata.Modified,
            Encrypted = true,
        };

        var envelopeBytes = JsonSerializer.SerializeToUtf8Bytes(envelope, JsonDefaults.Options);
        AddPart(archive, outerManifest, IwpfPaths.SecurityEnvelope, IwpfMediaTypes.Json,         envelopeBytes);
        AddPart(archive, outerManifest, IwpfPaths.SecurityPayload,  IwpfMediaTypes.OctetStream,   cipherText);

        var manifestBytes = JsonSerializer.SerializeToUtf8Bytes(outerManifest, JsonDefaults.Options);
        WriteEntry(archive, IwpfPaths.Manifest, manifestBytes);
    }
}

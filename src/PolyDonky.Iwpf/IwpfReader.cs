using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using PolyDonky.Core;

namespace PolyDonky.Iwpf;

/// <summary>IWPF 패키지를 읽어 PolyDonkyument 로 복원한다.</summary>
public sealed class IwpfReader : IDocumentReader
{
    public string FormatId => "iwpf";

    /// <summary>true 면 매니페스트의 SHA-256 해시와 실제 페이로드 해시를 비교한다.</summary>
    public bool VerifyHashes { get; init; } = true;

    public PolyDonkyument Read(Stream input) => Read(input, password: null);

    /// <summary>
    /// password 가 null/빈 문자열이면 암호화되지 않은 IWPF 만 읽을 수 있다.
    /// 암호화된 패키지를 비밀번호 없이 읽으려 하면 <see cref="EncryptedIwpfException"/> 을 던지고,
    /// 비밀번호가 틀리면 <see cref="WrongIwpfPasswordException"/> 을 던진다.
    /// </summary>
    public PolyDonkyument Read(Stream input, string? password)
    {
        ArgumentNullException.ThrowIfNull(input);

        // ZIP 은 random access 를 요구하므로 stream 이 seek 불가하면 메모리에 펼친다.
        Stream zipStream = input;
        MemoryStream? buffered = null;
        if (!input.CanSeek)
        {
            buffered = new MemoryStream();
            input.CopyTo(buffered);
            buffered.Position = 0;
            zipStream = buffered;
        }

        try
        {
            using var archive = new ZipArchive(zipStream, ZipArchiveMode.Read, leaveOpen: true);

            var manifest = ReadManifest(archive);

            // 암호화된 envelope 패키지면 inner ZIP 을 복호화한 뒤 재귀적으로 읽는다.
            if (manifest.Encrypted)
            {
                if (string.IsNullOrEmpty(password))
                    throw new EncryptedIwpfException();

                var envelopeBytes = ReadEntry(archive, IwpfPaths.SecurityEnvelope)
                    ?? throw new InvalidDataException($"Encrypted IWPF is missing '{IwpfPaths.SecurityEnvelope}'.");
                var payloadBytes = ReadEntry(archive, IwpfPaths.SecurityPayload)
                    ?? throw new InvalidDataException($"Encrypted IWPF is missing '{IwpfPaths.SecurityPayload}'.");

                var envelope = JsonSerializer.Deserialize<IwpfEnvelope>(envelopeBytes, JsonDefaults.Options)
                    ?? throw new InvalidDataException("Failed to deserialize security/envelope.json.");

                var innerBytes = IwpfEncryption.Decrypt(payloadBytes, envelope, password);
                using var innerStream = new MemoryStream(innerBytes, writable: false);
                // 풀린 inner 는 평문 IWPF 다 — password=null 로 재귀.
                return Read(innerStream, password: null);
            }

            var documentBytes = ReadEntry(archive, IwpfPaths.DocumentJson)
                ?? throw new InvalidDataException($"IWPF package is missing required part '{IwpfPaths.DocumentJson}'.");

            if (VerifyHashes && manifest.Parts.TryGetValue(IwpfPaths.DocumentJson, out var docEntry))
            {
                VerifyHash(IwpfPaths.DocumentJson, documentBytes, docEntry.Sha256);
            }

            var document = JsonSerializer.Deserialize<PolyDonkyument>(documentBytes, JsonDefaults.Options)
                ?? throw new InvalidDataException("Failed to deserialize content/document.json.");

            // 쓰기 잠금 — manifest.WriteLocked 이면 security/write-lock.json 을 읽어
            // Metadata.Custom 에 보관한다. ViewModel 이 이를 꺼내 _writeLock 으로 저장한다.
            if (manifest.WriteLocked)
            {
                var writeLockBytes = ReadEntry(archive, IwpfPaths.SecurityWriteLock);
                if (writeLockBytes is not null)
                {
                    if (VerifyHashes && manifest.Parts.TryGetValue(IwpfPaths.SecurityWriteLock, out var wlEntry))
                    {
                        VerifyHash(IwpfPaths.SecurityWriteLock, writeLockBytes, wlEntry.Sha256);
                    }
                    document.Metadata.Custom["iwpf.writeLock"] = Encoding.UTF8.GetString(writeLockBytes);
                }
            }

            // styles 파트 (있을 때만 머지)
            var stylesBytes = ReadEntry(archive, IwpfPaths.StylesJson);
            if (stylesBytes is not null)
            {
                if (VerifyHashes && manifest.Parts.TryGetValue(IwpfPaths.StylesJson, out var stEntry))
                {
                    VerifyHash(IwpfPaths.StylesJson, stylesBytes, stEntry.Sha256);
                }
                document.Styles = JsonSerializer.Deserialize<StyleSheet>(stylesBytes, JsonDefaults.Options)
                    ?? new StyleSheet();
            }

            // provenance 파트 (선택)
            var provenanceBytes = ReadEntry(archive, IwpfPaths.ProvenanceJson);
            if (provenanceBytes is not null)
            {
                if (VerifyHashes && manifest.Parts.TryGetValue(IwpfPaths.ProvenanceJson, out var prEntry))
                {
                    VerifyHash(IwpfPaths.ProvenanceJson, provenanceBytes, prEntry.Sha256);
                }
                document.Provenance = JsonSerializer.Deserialize<Provenance>(provenanceBytes, JsonDefaults.Options)
                    ?? new Provenance();
            }

            // resources/images/* — ImageBlock.ResourcePath 가 있는 블록의 Data 를 다시 채운다.
            RehydrateImageResources(document, archive, manifest);

            // 옛 빌드의 연속 Y 좌표를 페이지-로컬 좌표로 자동 마이그레이션 (멱등).
            OverlayAnchorMigration.MigrateContinuousAnchors(document);

            return document;
        }
        finally
        {
            buffered?.Dispose();
        }
    }

    private static IwpfManifest ReadManifest(ZipArchive archive)
    {
        var bytes = ReadEntry(archive, IwpfPaths.Manifest)
            ?? throw new InvalidDataException($"IWPF package is missing '{IwpfPaths.Manifest}'.");

        var manifest = JsonSerializer.Deserialize<IwpfManifest>(bytes, JsonDefaults.Options)
            ?? throw new InvalidDataException("Failed to deserialize manifest.json.");

        if (manifest.PackageType != "polydoc.iwpf")
        {
            throw new InvalidDataException(
                $"Unexpected package type '{manifest.PackageType}'. Expected 'polydoc.iwpf'.");
        }

        return manifest;
    }

    private static byte[]? ReadEntry(ZipArchive archive, string path)
    {
        var entry = archive.GetEntry(path);
        if (entry is null)
        {
            return null;
        }
        using var stream = entry.Open();
        using var ms = new MemoryStream(checked((int)entry.Length));
        stream.CopyTo(ms);
        return ms.ToArray();
    }

    private static void VerifyHash(string path, byte[] payload, string expectedHex)
    {
        Span<byte> hash = stackalloc byte[SHA256.HashSizeInBytes];
        SHA256.HashData(payload, hash);
        var actualHex = Convert.ToHexStringLower(hash);
        if (!string.Equals(actualHex, expectedHex, StringComparison.Ordinal))
        {
            throw new InvalidDataException(
                $"IWPF integrity check failed for '{path}'. Expected SHA-256 {expectedHex}, got {actualHex}.");
        }
    }

    private void RehydrateImageResources(PolyDonkyument document, ZipArchive archive, IwpfManifest manifest)
    {
        var cache = new Dictionary<string, byte[]>(StringComparer.Ordinal);

        foreach (var section in document.Sections)
        {
            Walk(section.Blocks);
        }
        return;

        void Walk(IList<Block> blocks)
        {
            foreach (var block in blocks)
            {
                switch (block)
                {
                    case ImageBlock image when !string.IsNullOrEmpty(image.ResourcePath):
                        var loaded = LoadResource(image.ResourcePath!);
                        if (loaded is not null)
                        {
                            image.Data = loaded;
                        }
                        break;
                    case Table table:
                        foreach (var row in table.Rows)
                            foreach (var cell in row.Cells)
                                Walk(cell.Blocks);
                        break;
                    case TextBoxObject textbox:
                        Walk(textbox.Content);
                        break;
                }
            }
        }

        byte[]? LoadResource(string path)
        {
            // 외부 URI 참조(http/https/file/data: …)는 IWPF 내부 파트가 아니므로
            // ZIP 엔트리로 간주하지 않고 ResourcePath 만 보존한 채 데이터는 비워둔다.
            // HTML/Markdown 에서 ingest 된 원격 이미지가 그대로 IWPF 로 저장된 경우를 대비.
            if (IsExternalReference(path))
            {
                return null;
            }

            if (cache.TryGetValue(path, out var cached))
            {
                return cached;
            }
            var payload = ReadEntry(archive, path)
                ?? throw new InvalidDataException(
                    $"IWPF package is missing referenced image part '{path}'.");
            if (VerifyHashes && manifest.Parts.TryGetValue(path, out var entry))
            {
                VerifyHash(path, payload, entry.Sha256);
            }
            cache[path] = payload;
            return payload;
        }
    }

    private static bool IsExternalReference(string path)
    {
        if (path.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }
        // 절대 URI 면 외부 참조 — 단, "C:" 같은 단일 문자 스킴은 거른다 (IWPF 내부 경로는
        // 상대 ZIP 경로라 절대 URI 가 될 수 없다).
        return Uri.TryCreate(path, UriKind.Absolute, out var uri)
            && uri.Scheme.Length > 1;
    }
}

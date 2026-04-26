using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using PolyDoc.Core;

namespace PolyDoc.Iwpf;

/// <summary>IWPF 패키지를 읽어 PolyDocument 로 복원한다.</summary>
public sealed class IwpfReader : IDocumentReader
{
    public string FormatId => "iwpf";

    /// <summary>true 면 매니페스트의 SHA-256 해시와 실제 페이로드 해시를 비교한다.</summary>
    public bool VerifyHashes { get; init; } = true;

    public PolyDocument Read(Stream input) => Read(input, password: null);

    /// <summary>
    /// password 가 null/빈 문자열이면 암호화되지 않은 IWPF 만 읽을 수 있다.
    /// 암호화된 패키지를 비밀번호 없이 읽으려 하면 <see cref="EncryptedIwpfException"/> 을 던지고,
    /// 비밀번호가 틀리면 <see cref="WrongIwpfPasswordException"/> 을 던진다.
    /// </summary>
    public PolyDocument Read(Stream input, string? password)
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

            var document = JsonSerializer.Deserialize<PolyDocument>(documentBytes, JsonDefaults.Options)
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

    private void RehydrateImageResources(PolyDocument document, ZipArchive archive, IwpfManifest manifest)
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
                        image.Data = LoadResource(image.ResourcePath!);
                        break;
                    case Table table:
                        foreach (var row in table.Rows)
                            foreach (var cell in row.Cells)
                                Walk(cell.Blocks);
                        break;
                }
            }
        }

        byte[] LoadResource(string path)
        {
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
}

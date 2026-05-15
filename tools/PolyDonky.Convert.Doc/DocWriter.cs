using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using PolyDonky.Core;

namespace PolyDonky.Convert.Doc;

/// <summary>
/// IWPF → DOC (97-2003) 변환기. 현재 단계: 텍스트만.
/// OLE Compound Document 형식으로 최소 구조를 생성한다.
/// </summary>
public class DocWriter
{
    private const int SectorSize = 512;
    private const int MiniFatSectorSize = 64;
    private const uint EndOfChain = 0xFFFFFFFE;
    private const uint FatSector = 0xFFFFFFFD;

    public void Write(PolyDonkyument doc, Stream output)
    {
        var writer = new OleDocumentWriter(output);

        // 전체 텍스트 추출
        var text = ExtractAllText(doc);

        // 최소 WordDocument 스트림 생성
        var wordDocStream = CreateWordDocumentStream(text);

        // OLE 문서에 스트림 추가
        writer.AddStream("WordDocument", wordDocStream);

        // OLE 문서 작성
        writer.Write();
    }

    private string ExtractAllText(PolyDonkyument doc)
    {
        var sb = new StringBuilder();
        foreach (var section in doc.Sections)
        {
            foreach (var block in section.Blocks)
            {
                if (block is Paragraph para)
                {
                    foreach (var run in para.Runs)
                    {
                        if (!string.IsNullOrEmpty(run.Text))
                            sb.Append(run.Text);
                    }
                    sb.AppendLine();
                }
            }
        }
        return sb.ToString();
    }

    private byte[] CreateWordDocumentStream(string text)
    {
        using var ms = new MemoryStream();
        using var writer = new BinaryWriter(ms, Encoding.UTF8, leaveOpen: true);

        // Word Document Stream header (minimal for 97-2003)
        // 바이트 0-1: 버전 정보
        writer.Write((ushort)0xDB95);  // Word 7.0 format

        // 바이트 2-3: 버전 생성
        writer.Write((ushort)0x0000);

        // 바이트 4-5: 버전 최소
        writer.Write((ushort)0x0000);

        // 바이트 6-7: 문서 타입
        writer.Write((ushort)0x0001);  // Normal document

        // 추가 헤더 정보 (간소화)
        for (int i = 0; i < 20; i++)
            writer.Write((uint)0);

        // 텍스트 콘텐츠를 스트림에 기록
        var textBytes = Encoding.UTF8.GetBytes(text);
        writer.Write(textBytes);

        ms.Flush();
        return ms.ToArray();
    }
}

/// <summary>
/// OLE Compound Document 작성기.
/// 최소 구조만 생성: 루트 엔트리 + 단일 스트림.
/// </summary>
internal class OleDocumentWriter
{
    private const uint EndOfChain = 0xFFFFFFFE;
    private const uint FatSector = 0xFFFFFFFD;
    private const int SectorSize = 512;

    private readonly Stream _output;
    private readonly Dictionary<string, byte[]> _streams = new();

    public OleDocumentWriter(Stream output)
    {
        _output = output;
    }

    public void AddStream(string name, byte[] data)
    {
        _streams[name] = data;
    }

    public void Write()
    {
        // OLE 구조:
        // [Header (512)] [FAT (512)] [Mini FAT (512)] [Root Entry (512)] [Stream data...]

        var header = CreateHeader();
        var fatData = CreateFat();
        var miniFatData = CreateMiniFat();
        var rootEntry = CreateRootEntry();
        var streamData = CombineStreamData();

        _output.Write(header, 0, header.Length);
        _output.Write(fatData, 0, fatData.Length);
        _output.Write(miniFatData, 0, miniFatData.Length);
        _output.Write(rootEntry, 0, rootEntry.Length);
        _output.Write(streamData, 0, streamData.Length);
        _output.Flush();
    }

    private byte[] CreateHeader()
    {
        var header = new byte[512];

        // Signature (8 bytes)
        header[0] = 0xD0;
        header[1] = 0xCF;
        header[2] = 0x11;
        header[3] = 0xE0;
        header[4] = 0xA1;
        header[5] = 0xB1;
        header[6] = 0x1A;
        header[7] = 0xE1;

        // CLSID (16 bytes, all zeros)
        // Already zero-initialized

        // Minor version (2 bytes)
        Array.Copy(BitConverter.GetBytes((ushort)0x003E), 0, header, 24, 2);

        // Major version (2 bytes) - 3 for 512-byte sectors
        Array.Copy(BitConverter.GetBytes((ushort)0x0003), 0, header, 26, 2);

        // Byte order identifier (2 bytes) - little endian
        Array.Copy(BitConverter.GetBytes((ushort)0xFFFE), 0, header, 28, 2);

        // Sector shift (2 bytes) - 9 for 512 bytes (2^9)
        Array.Copy(BitConverter.GetBytes((ushort)0x0009), 0, header, 30, 2);

        // Mini sector shift (2 bytes) - 6 for 64 bytes (2^6)
        Array.Copy(BitConverter.GetBytes((ushort)0x0006), 0, header, 32, 2);

        // Total sectors (4 bytes, reserved)
        Array.Copy(BitConverter.GetBytes((uint)0), 0, header, 48, 4);

        // Number of FAT sectors (4 bytes)
        Array.Copy(BitConverter.GetBytes((uint)1), 0, header, 44, 4);

        // Directory first sector (4 bytes) - sector 2
        Array.Copy(BitConverter.GetBytes((uint)2), 0, header, 48, 4);

        // First mini FAT sector (4 bytes) - sector 1
        Array.Copy(BitConverter.GetBytes((uint)1), 0, header, 60, 4);

        // Number of mini FAT sectors (4 bytes)
        Array.Copy(BitConverter.GetBytes((uint)1), 0, header, 64, 4);

        // First DIFAT sector (4 bytes)
        Array.Copy(BitConverter.GetBytes((uint)EndOfChain), 0, header, 68, 4);

        // Number of DIFAT sectors (4 bytes)
        Array.Copy(BitConverter.GetBytes((uint)0), 0, header, 72, 4);

        // DIFAT array (first 109 FAT sector positions)
        Array.Copy(BitConverter.GetBytes((uint)0), 0, header, 76, 4);
        for (int i = 1; i < 109; i++)
        {
            Array.Copy(BitConverter.GetBytes((uint)EndOfChain), 0, header, 76 + i * 4, 4);
        }

        return header;
    }

    private byte[] CreateFat()
    {
        var fat = new byte[512];
        var fatEntries = new uint[128];

        // Mark FAT sector itself
        fatEntries[0] = FatSector;

        // Mark mini FAT sector
        fatEntries[1] = FatSector;

        // Mark directory sector
        fatEntries[2] = FatSector;

        // Mark stream sectors
        int sectorIndex = 3;
        int totalDataSize = 0;
        foreach (var streamData in _streams.Values)
            totalDataSize += streamData.Length;

        int sectorsNeeded = (totalDataSize + 511) / 512;
        for (int i = 0; i < sectorsNeeded - 1; i++)
            fatEntries[sectorIndex + i] = (uint)(sectorIndex + i + 1);
        if (sectorsNeeded > 0)
            fatEntries[sectorIndex + sectorsNeeded - 1] = EndOfChain;

        // Fill rest with free sectors
        for (int i = sectorIndex + sectorsNeeded; i < fatEntries.Length; i++)
            fatEntries[i] = 0xFFFFFFFE;  // FREESECT

        // Convert to bytes
        Buffer.BlockCopy(fatEntries, 0, fat, 0, fat.Length);
        return fat;
    }

    private byte[] CreateMiniFat()
    {
        // Mini FAT (all free for now)
        return new byte[512];
    }

    private byte[] CreateRootEntry()
    {
        var entries = new byte[512];

        // Root entry (128 bytes)
        var rootName = Encoding.Unicode.GetBytes("Root Entry");
        Array.Copy(rootName, 0, entries, 0, rootName.Length);

        // Name length (2 bytes) - includes null terminator
        var nameLen = (ushort)(rootName.Length + 2);
        Array.Copy(BitConverter.GetBytes(nameLen), 0, entries, 64, 2);

        // Entry type (1 byte) - 5 = root storage
        entries[66] = 5;

        // Color (1 byte) - 1 = black
        entries[67] = 1;

        // First child DID (4 bytes)
        Array.Copy(BitConverter.GetBytes((uint)0xFFFFFFFF), 0, entries, 68, 4);

        // Next sibling DID (4 bytes)
        Array.Copy(BitConverter.GetBytes((uint)0xFFFFFFFF), 0, entries, 72, 4);

        // Prev sibling DID (4 bytes)
        Array.Copy(BitConverter.GetBytes((uint)0xFFFFFFFF), 0, entries, 76, 4);

        // DID of first stream sector (4 bytes)
        if (_streams.Count > 0)
            Array.Copy(BitConverter.GetBytes((uint)1), 0, entries, 116, 4);  // Stream starts at sector 3
        else
            Array.Copy(BitConverter.GetBytes((uint)EndOfChain), 0, entries, 116, 4);

        // Size (4 bytes)
        int totalSize = _streams.Values.Sum(s => s.Length);
        Array.Copy(BitConverter.GetBytes((uint)totalSize), 0, entries, 120, 4);

        // Stream entry (if we have streams)
        if (_streams.Count > 0)
        {
            var streamName = "WordDocument";
            var streamNameBytes = Encoding.Unicode.GetBytes(streamName);
            var streamNameLen = (ushort)(streamNameBytes.Length + 2);

            // Copy stream name
            Array.Copy(streamNameBytes, 0, entries, 128, streamNameBytes.Length);

            // Stream name length
            Array.Copy(BitConverter.GetBytes(streamNameLen), 0, entries, 128 + 64, 2);

            // Stream entry type (1 byte) - 2 = stream
            entries[128 + 66] = 2;

            // Stream color
            entries[128 + 67] = 1;

            // Stream siblings
            Array.Copy(BitConverter.GetBytes((uint)0xFFFFFFFF), 0, entries, 128 + 68, 4);
            Array.Copy(BitConverter.GetBytes((uint)0xFFFFFFFF), 0, entries, 128 + 72, 4);
            Array.Copy(BitConverter.GetBytes((uint)0xFFFFFFFF), 0, entries, 128 + 76, 4);

            // Stream start sector
            Array.Copy(BitConverter.GetBytes((uint)3), 0, entries, 128 + 116, 4);

            // Stream size
            Array.Copy(BitConverter.GetBytes((uint)totalSize), 0, entries, 128 + 120, 4);
        }

        return entries;
    }

    private byte[] CombineStreamData()
    {
        using var ms = new MemoryStream();
        foreach (var streamData in _streams.Values)
        {
            ms.Write(streamData, 0, streamData.Length);

            // Pad to 512-byte boundary
            int padding = 512 - (streamData.Length % 512);
            if (padding < 512)
                ms.Write(new byte[padding], 0, padding);
        }
        return ms.ToArray();
    }
}

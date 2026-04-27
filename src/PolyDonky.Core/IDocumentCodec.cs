namespace PolyDonky.Core;

/// <summary>한 포맷에 대한 reader/writer 결합 인터페이스. 메인 앱이 Format 매트릭스를 동적으로 구성한다.</summary>
public interface IDocumentCodec : IDocumentReader, IDocumentWriter
{
}

public interface IDocumentReader
{
    /// <summary>예: "iwpf", "txt", "md", "docx", "hwpx".</summary>
    string FormatId { get; }

    PolyDonkyument Read(Stream input);

    PolyDonkyument ReadFile(string path)
    {
        using var fs = File.OpenRead(path);
        return Read(fs);
    }
}

public interface IDocumentWriter
{
    string FormatId { get; }

    void Write(PolyDonkyument document, Stream output);

    void WriteFile(PolyDonkyument document, string path)
    {
        ArgumentNullException.ThrowIfNull(document);
        using var fs = File.Create(path);
        Write(document, fs);
    }
}

namespace PolyDoc.Core;

/// <summary>
/// PolyDocument 가 차지하는 데이터 크기를 근사 계산한다.
/// "이 문서가 메모리에서 얼마나 차지하나" 의 사용자 직관에 맞춘 값:
///   - Run.Text 의 문자 byte (UTF-16 = char × 2)
///   - ImageBlock.Data 와 OpaqueBlock.Bytes 의 raw byte 길이
///   - OpaqueBlock.Xml 의 문자 byte (UTF-16)
///   - Run/Block 마다 약간의 오버헤드 가산
///   - Table 은 셀 안의 Block 들로 재귀
/// .NET 객체 그래프 메모리(GC heap)와 정확히 일치하지는 않지만, 사용자가
/// 문서 크기 변화를 직관적으로 인지할 수 있는 안정적 척도.
/// </summary>
public static class DocumentMeasurement
{
    private const int RunOverheadBytes = 64;
    private const int BlockOverheadBytes = 96;

    public static long EstimateBytes(PolyDocument document)
    {
        ArgumentNullException.ThrowIfNull(document);
        long total = 0;
        foreach (var section in document.Sections)
        {
            total += MeasureBlocks(section.Blocks);
        }
        return total;
    }

    private static long MeasureBlocks(IList<Block> blocks)
    {
        long total = 0;
        foreach (var block in blocks)
        {
            total += BlockOverheadBytes;
            switch (block)
            {
                case Paragraph p:
                    foreach (var run in p.Runs)
                    {
                        total += RunOverheadBytes + (long)run.Text.Length * sizeof(char);
                    }
                    break;
                case Table t:
                    foreach (var row in t.Rows)
                    {
                        foreach (var cell in row.Cells)
                        {
                            total += MeasureBlocks(cell.Blocks);
                        }
                    }
                    break;
                case ImageBlock img:
                    total += img.Data.LongLength;
                    break;
                case OpaqueBlock op:
                    total += (long)(op.Xml?.Length ?? 0) * sizeof(char);
                    total += op.Bytes?.LongLength ?? 0;
                    break;
            }
        }
        return total;
    }

    public static string FormatBytes(long bytes)
    {
        if (bytes < 1024) return $"{bytes} B";
        if (bytes < 1024L * 1024) return $"{bytes / 1024.0:F1} KB";
        if (bytes < 1024L * 1024 * 1024) return $"{bytes / 1024.0 / 1024.0:F1} MB";
        return $"{bytes / 1024.0 / 1024.0 / 1024.0:F2} GB";
    }
}

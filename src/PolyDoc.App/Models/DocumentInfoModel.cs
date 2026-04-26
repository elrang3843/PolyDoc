namespace PolyDoc.App.Models;

public sealed class DocumentInfoModel
{
    // 파일 정보
    public string FilePath  { get; init; } = "";
    public string Format    { get; init; } = "";
    public string DataSize  { get; init; } = "";

    // 문서 속성
    public string DocTitle  { get; init; } = "";
    public string Author    { get; init; } = "";
    public string Language  { get; init; } = "";
    public string Created   { get; init; } = "";
    public string Modified  { get; init; } = "";

    // 통계
    public string ParagraphCount { get; init; } = "";
    public string CharCount      { get; init; } = "";
    public string WordCount      { get; init; } = "";
    public string LineCount      { get; init; } = "";
    public string SectionCount   { get; init; } = "";
    public string TableCount     { get; init; } = "";
    public string ImageCount     { get; init; } = "";
}

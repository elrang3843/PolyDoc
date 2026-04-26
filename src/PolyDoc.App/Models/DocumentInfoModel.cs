using CommunityToolkit.Mvvm.ComponentModel;
using PolyDoc.Iwpf;

namespace PolyDoc.App.Models;

/// <summary>
/// 문서 정보 다이얼로그용 view model. 통계/메타데이터의 읽기 표시 + 사용자가 편집 가능한
/// (Author, 암호 보호 모드, 워터마크) 항목을 함께 담는다.
/// </summary>
public sealed partial class DocumentInfoModel : ObservableObject
{
    // ── 파일 정보 (read-only) ────────────────────────────────
    public string FilePath  { get; init; } = "";
    public string Format    { get; init; } = "";
    public string DataSize  { get; init; } = "";

    // ── 문서 속성 ────────────────────────────────────────────
    public string DocTitle  { get; init; } = "";
    public string Language  { get; init; } = "";
    public string Created   { get; init; } = "";
    public string Modified  { get; init; } = "";

    /// <summary>편집 가능한 작성자 이름. 다이얼로그에서 두 방향 바인딩.</summary>
    [ObservableProperty]
    private string _author = "";

    // ── 통계 (read-only) ─────────────────────────────────────
    public string ParagraphCount { get; init; } = "";
    public string CharCount      { get; init; } = "";
    public string WordCount      { get; init; } = "";
    public string LineCount      { get; init; } = "";
    public string SectionCount   { get; init; } = "";
    public string TableCount     { get; init; } = "";
    public string ImageCount     { get; init; } = "";

    // ── 보안 ────────────────────────────────────────────────
    /// <summary>현재 문서의 암호 보호 모드.</summary>
    [ObservableProperty]
    private PasswordMode _passwordMode;

    /// <summary>
    /// PasswordChangeWindow 를 거쳐 사용자가 선택한 새 모드/비밀번호.
    /// 다이얼로그를 거치지 않았으면 <c>PasswordChanged == false</c> — 변경 없음.
    /// Both 모드에서 같은 암호이면 두 값이 동일하고, 다른 암호이면 각각 다른 값이 설정된다.
    /// </summary>
    public bool PasswordChanged       { get; set; }
    public string? NewReadPassword    { get; set; }
    public string? NewWritePassword   { get; set; }

    // ── 워터마크 ────────────────────────────────────────────
    [ObservableProperty]
    private bool _watermarkEnabled;

    [ObservableProperty]
    private string _watermarkText = "";

    [ObservableProperty]
    private string _watermarkColor = "#FF808080";

    [ObservableProperty]
    private int _watermarkFontSize = 48;

    [ObservableProperty]
    private double _watermarkRotation = -45.0;

    [ObservableProperty]
    private double _watermarkOpacity = 0.3;
}

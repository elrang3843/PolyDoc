namespace PolyDonky.Core;

/// <summary>
/// 클립보드·부분 가져오기/내보내기에서 사용하는 IWPF 블록 묶음 포맷.
/// <see cref="Section"/> 에서 페이지 설정을 뺀 경량 컨테이너.
/// 클립보드 포맷명: <c>PolyDonky.IwpfFragment.v1</c>.
/// </summary>
public sealed class IwpfFragment
{
    public string Version { get; init; } = "1.0";
    public List<Block> Blocks { get; init; } = new();
}

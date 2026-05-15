namespace PolyDonky.Convert.Common;

/// <summary>
/// CLI 변환기 종료 코드 상수.
/// </summary>
public static class ConverterExitCodes
{
    /// <summary>성공</summary>
    public const int Ok = 0;

    /// <summary>인자 오류</summary>
    public const int BadArgs = 2;

    /// <summary>지원하지 않는 변환 쌍</summary>
    public const int UnsupportedOp = 3;

    /// <summary>입출력 실패 (파일 없음·권한·디렉터리 없음·디스크 잠금 등)</summary>
    public const int IoError = 4;

    /// <summary>변환 실패 (구조 손상·내부 예외)</summary>
    public const int ConvertError = 5;

    /// <summary>지원하지 않는 옛 버전</summary>
    public const int OldVersion = 6;
}

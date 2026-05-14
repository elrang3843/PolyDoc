namespace PolyDonky.Convert.Common;

/// <summary>
/// CLI 변환기 진행상황 보고.
/// 메인 앱(Services/ExternalConverter)이 stdout을 한 줄씩 읽고
/// "PROGRESS:" 접두로 시작하는 줄을 파싱해 IProgress로 보고한다.
/// </summary>
public static class ConverterProgress
{
    /// <summary>
    /// 진행상황을 stdout으로 보고.
    /// </summary>
    /// <param name="percent">진행률 (0-100)</param>
    /// <param name="message">진행 메시지</param>
    public static void Write(int percent, string message)
    {
        Console.WriteLine($"PROGRESS:{percent}:{message}");
        Console.Out.Flush();
    }
}

using System.Globalization;

namespace PolyDonky.Core;

/// <summary>
/// 문서 단위 변환 유틸리티.
///   EMU  (English Metric Unit) : 1 inch = 914 400 EMU = 25.4 mm → 36 000 EMU/mm
///   Twip (twentieth of a point): 1 inch =   1 440 twips = 25.4 mm → 1440/25.4 twips/mm
///   HwpUnit                    : 1 inch =   7 200 HWPUNIT = 25.4 mm → 7200/25.4 hwpunit/mm
/// </summary>
public static class UnitConverter
{
    // ── EMU ↔ mm ─────────────────────────────────────────────────────────────

    public static double EmuToMm(long emu) => emu / 36000.0;
    public static long MmToEmu(double mm) => (long)Math.Round(mm * 36000.0);

    // ── Twips ↔ mm ───────────────────────────────────────────────────────────

    private const double TwipsPerMm = 1440.0 / 25.4;

    public static double TwipsToMm(double twips) => twips / TwipsPerMm;

    public static int MmToTwipsInt(double mm) => (int)Math.Round(mm * TwipsPerMm);
    public static uint MmToTwipsUInt(double mm) => (uint)Math.Round(mm * TwipsPerMm);
    public static string MmToTwipsString(double mm)
        => MmToTwipsInt(mm).ToString(CultureInfo.InvariantCulture);

    public static double ParseTwipsToMm(string? twipsRaw)
    {
        if (string.IsNullOrEmpty(twipsRaw)) return 0;
        return double.TryParse(twipsRaw, NumberStyles.Float, CultureInfo.InvariantCulture, out var twips)
            ? TwipsToMm(twips)
            : 0;
    }

    // ── HwpUnit ↔ mm ─────────────────────────────────────────────────────────

    private const double HwpUnitPerMm = 7200.0 / 25.4;

    public static double HwpUnitToMm(double hwpUnit) => hwpUnit / HwpUnitPerMm;
    public static long MmToHwpUnit(double mm) => (long)Math.Round(mm * HwpUnitPerMm);
}

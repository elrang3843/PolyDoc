using System;
using System.IO;
using Microsoft.Win32;

namespace PolyDonky.App.Services;

/// <summary>
/// Windows에서 LibreOffice 설치 경로를 자동 탐지하는 유틸리티.
/// </summary>
public static class LibreOfficeLocator
{
    /// <summary>
    /// LibreOffice 설치 경로 자동 탐지.
    /// 레지스트리 → 일반 경로 순서로 검색. 첫 번째 유효 경로 반환.
    /// </summary>
    /// <returns>LibreOffice 설치 폴더 경로 (null이면 미설치)</returns>
    public static string? DetectLibreOfficePath()
    {
        // 1순위: LibreOffice 전용 레지스트리 키
        var regPath = SearchRegistry();
        if (regPath != null) return regPath;

        // 2순위: Windows 설치 정보 (Uninstall 레지스트리) — 사용자 지정 설치 경로도 포함
        var uninstallPath = SearchUninstallRegistry();
        if (uninstallPath != null) return uninstallPath;

        // 3순위: 일반 설치 경로 (기본 폴더)
        var commonPath = SearchCommonPaths();
        if (commonPath != null) return commonPath;

        return null;
    }

    /// <summary>
    /// 주어진 경로에 유효한 LibreOffice soffice.exe 또는 soffice 파일이 있는지 확인.
    /// </summary>
    public static bool ValidatePath(string? path)
    {
        if (string.IsNullOrEmpty(path))
            return false;

        try
        {
            // 탐색 순서: program\ (Windows 표준) → bin\ (일부 배포) → 직접 경로
            string[] candidates =
            [
                Path.Combine(path, "program", "soffice.exe"),  // Windows 표준
                Path.Combine(path, "program", "soffice"),      // Linux
                Path.Combine(path, "bin",     "soffice.exe"),  // 구형/대안
                Path.Combine(path, "bin",     "soffice"),      // Linux 대안
                Path.Combine(path,            "soffice.exe"),  // path 자체가 program 폴더인 경우
                Path.Combine(path,            "soffice"),
            ];
            foreach (var c in candidates)
                if (File.Exists(c)) return true;
            return false;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Windows 레지스트리에서 LibreOffice 설치 경로 검색.
    /// HKEY_LOCAL_MACHINE\SOFTWARE\LibreOffice 조회.
    /// </summary>
    private static string? SearchRegistry()
    {
        try
        {
            // InstallPath 는 "C:\Program Files\LibreOffice\program" 처럼 program 폴더를 가리킬 수 있음.
            // ValidatePath 가 그 경로 자체에서 soffice.exe 를 찾거나, 부모 폴더에서 program\soffice.exe 를 찾음.
            string?[] candidates =
            [
                GetRegValue(@"SOFTWARE\LibreOffice\UNO\InstallPath", ""),
                GetRegValue(@"SOFTWARE\WOW6432Node\LibreOffice\UNO\InstallPath", ""),
                GetRegValue(@"SOFTWARE\LibreOffice\UNO\InstallPath", null),
                GetRegValue(@"SOFTWARE\WOW6432Node\LibreOffice\UNO\InstallPath", null),
            ];
            foreach (var p in candidates)
            {
                if (string.IsNullOrEmpty(p)) continue;
                // 레지스트리가 program 폴더를 가리키는 경우 부모(설치 루트)도 시도
                if (ValidatePath(p)) return p;
                var parent = Path.GetDirectoryName(p);
                if (!string.IsNullOrEmpty(parent) && ValidatePath(parent)) return parent;
            }
            return null;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Windows 프로그램 추가/제거 정보(Uninstall 레지스트리)에서 LibreOffice 설치 위치를 조회.
    /// 사용자가 기본 폴더가 아닌 곳에 설치했어도 찾을 수 있다.
    /// </summary>
    private static string? SearchUninstallRegistry()
    {
        string[] uninstallRoots =
        [
            @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
            @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
        ];

        foreach (var root in uninstallRoots)
        {
            var found = ScanUninstallKey(Registry.LocalMachine, root);
            if (found != null) return found;
        }

        // 현재 사용자 설치 (HKCU)
        var foundUser = ScanUninstallKey(Registry.CurrentUser,
            @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall");
        if (foundUser != null) return foundUser;

        return null;
    }

    private static string? ScanUninstallKey(RegistryKey hive, string subKeyPath)
    {
        try
        {
            using var root = hive.OpenSubKey(subKeyPath);
            if (root == null) return null;

            foreach (var name in root.GetSubKeyNames())
            {
                using var entry = root.OpenSubKey(name);
                if (entry == null) continue;

                var displayName = entry.GetValue("DisplayName") as string ?? "";
                if (!displayName.Contains("LibreOffice", StringComparison.OrdinalIgnoreCase))
                    continue;

                // InstallLocation 값이 설치 폴더를 가리킴
                var installLocation = entry.GetValue("InstallLocation") as string;
                if (!string.IsNullOrEmpty(installLocation) && ValidatePath(installLocation))
                    return installLocation;

                // InstallLocation 이 비어 있으면 UninstallString 에서 경로 추출
                var uninstallStr = entry.GetValue("UninstallString") as string ?? "";
                var exeStart     = uninstallStr.TrimStart('"');
                var exeEnd       = exeStart.IndexOf('"');
                var exePath      = exeEnd > 0 ? exeStart[..exeEnd] : exeStart.Split(' ')[0];
                var dir          = Path.GetDirectoryName(exePath);
                if (!string.IsNullOrEmpty(dir))
                {
                    if (ValidatePath(dir)) return dir;
                    var parent = Path.GetDirectoryName(dir);
                    if (!string.IsNullOrEmpty(parent) && ValidatePath(parent)) return parent;
                }
            }
        }
        catch { /* 레지스트리 접근 실패 무시 */ }
        return null;
    }

    private static string? GetRegValue(string subKey, string? valueName)
    {
        try
        {
            using var key = Registry.LocalMachine.OpenSubKey(subKey);
            return key?.GetValue(valueName ?? "") as string;
        }
        catch { return null; }
    }

    /// <summary>
    /// 일반적인 LibreOffice 설치 경로들을 순회.
    /// </summary>
    private static string? SearchCommonPaths()
    {
        string[] commonPaths =
        {
            @"C:\Program Files\LibreOffice",
            @"C:\Program Files (x86)\LibreOffice",
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),    "LibreOffice"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "LibreOffice"),
            // LocalAppData 하위 사용자 설치 경로 (일부 버전)
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "LibreOffice"),
        };

        foreach (var path in commonPaths)
        {
            try
            {
                if (Directory.Exists(path))
                    return path;
            }
            catch
            {
                // 경로 접근 오류 무시
            }
        }

        return null;
    }
}

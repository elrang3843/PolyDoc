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
        // 1순위: 레지스트리 검색
        var regPath = SearchRegistry();
        if (regPath != null && ValidatePath(regPath))
            return regPath;

        // 2순위: 일반 설치 경로 검색
        var commonPath = SearchCommonPaths();
        if (commonPath != null && ValidatePath(commonPath))
            return commonPath;

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
            // Windows: soffice.exe (bin 폴더)
            var exePath = Path.Combine(path, "bin", "soffice.exe");
            if (File.Exists(exePath))
                return true;

            // Fallback: soffice (Linux 호환성)
            var sOfficePath = Path.Combine(path, "bin", "soffice");
            if (File.Exists(sOfficePath))
                return true;

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
            // 64비트 레지스트리
            var key64 = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\LibreOffice\UNO\InstallPath");
            if (key64?.GetValue("") is string path64 && !string.IsNullOrEmpty(path64))
                return path64;

            // 32비트 레지스트리 (WOW6432Node)
            var key32 = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\WOW6432Node\LibreOffice\UNO\InstallPath");
            if (key32?.GetValue("") is string path32 && !string.IsNullOrEmpty(path32))
                return path32;

            // 다른 레지스트리 위치 시도
            var keyAlt = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\LibreOffice");
            if (keyAlt != null)
            {
                foreach (var valueName in keyAlt.GetValueNames())
                {
                    if (keyAlt.GetValue(valueName) is string altPath && !string.IsNullOrEmpty(altPath))
                        return altPath;
                }
            }

            return null;
        }
        catch
        {
            return null;
        }
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
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "LibreOffice"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "LibreOffice"),
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

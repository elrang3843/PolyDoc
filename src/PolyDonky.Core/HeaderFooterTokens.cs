using System.Globalization;
using System.Text;

namespace PolyDonky.Core;

/// <summary>
/// 머리말·꼬리말 텍스트의 토큰 치환을 담당. 토큰은 <c>{NAME}</c> 형식 — 영문 표준명과
/// 한국어 별칭을 모두 인식한다. 알 수 없는 토큰은 원본 그대로 둔다(편집·라운드트립 안전).
///
/// 지원 토큰
/// <list type="bullet">
///   <item><c>{PAGE}</c> / <c>{페이지}</c> — 현재 페이지 번호</item>
///   <item><c>{NUMPAGES}</c> / <c>{전체페이지}</c> — 총 페이지 수</item>
///   <item><c>{DATE}</c> / <c>{날짜}</c> — yyyy-MM-dd</item>
///   <item><c>{TIME}</c> / <c>{시간}</c> — HH:mm</item>
///   <item><c>{TITLE}</c> / <c>{제목}</c> — 문서 메타데이터 제목</item>
///   <item><c>{AUTHOR}</c> / <c>{저자}</c> — 문서 메타데이터 저자</item>
///   <item><c>{FILENAME}</c> / <c>{파일명}</c> — 확장자 없는 파일명</item>
/// </list>
/// 리터럴 중괄호가 필요하면 <c>\{</c> / <c>\}</c> 로 이스케이프한다.
/// </summary>
public static class HeaderFooterTokens
{
    public readonly struct Context
    {
        public int PageNumber { get; init; }
        public int TotalPages { get; init; }
        public DateTime Now   { get; init; }
        public string? Title  { get; init; }
        public string? Author { get; init; }
        public string? FileName { get; init; }
    }

    /// <summary>
    /// <paramref name="template"/> 안의 토큰을 <paramref name="ctx"/> 값으로 치환해 반환.
    /// null/빈 문자열이면 빈 문자열 반환.
    /// </summary>
    public static string Resolve(string? template, Context ctx)
    {
        if (string.IsNullOrEmpty(template)) return string.Empty;

        var sb = new StringBuilder(template.Length);
        int i = 0;
        while (i < template.Length)
        {
            char ch = template[i];

            // 이스케이프 \{ \}  → 리터럴
            if (ch == '\\' && i + 1 < template.Length &&
                (template[i + 1] == '{' || template[i + 1] == '}'))
            {
                sb.Append(template[i + 1]);
                i += 2;
                continue;
            }

            if (ch == '{')
            {
                int end = template.IndexOf('}', i + 1);
                if (end > i + 1)
                {
                    var name = template.Substring(i + 1, end - i - 1).Trim();
                    if (TryResolveToken(name, ctx, out var value))
                    {
                        sb.Append(value);
                        i = end + 1;
                        continue;
                    }
                }
                // 닫는 중괄호가 없거나 알 수 없는 토큰 — 원본 그대로 보존
            }

            sb.Append(ch);
            i++;
        }
        return sb.ToString();
    }

    private static bool TryResolveToken(string name, Context ctx, out string value)
    {
        switch (name.ToUpperInvariant())
        {
            case "PAGE":
            case "페이지":
                value = ctx.PageNumber.ToString(CultureInfo.InvariantCulture);
                return true;

            case "NUMPAGES":
            case "전체페이지":
                value = ctx.TotalPages.ToString(CultureInfo.InvariantCulture);
                return true;

            case "DATE":
            case "날짜":
                value = ctx.Now.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                return true;

            case "TIME":
            case "시간":
                value = ctx.Now.ToString("HH:mm", CultureInfo.InvariantCulture);
                return true;

            case "TITLE":
            case "제목":
                value = ctx.Title ?? string.Empty;
                return true;

            case "AUTHOR":
            case "저자":
                value = ctx.Author ?? string.Empty;
                return true;

            case "FILENAME":
            case "파일명":
                value = ctx.FileName ?? string.Empty;
                return true;

            default:
                value = string.Empty;
                return false;
        }
    }
}

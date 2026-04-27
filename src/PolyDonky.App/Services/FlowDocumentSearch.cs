using System.Windows.Documents;

namespace PolyDonky.App.Services;

/// <summary>
/// FlowDocument 에서 텍스트를 순방향으로 검색하고 대체하는 헬퍼.
/// WPF TextPointer 기반 — 검색어가 단일 Run 안에 있는 경우 정확히 매치한다.
/// </summary>
internal static class FlowDocumentSearch
{
    /// <summary>
    /// <paramref name="startFrom"/> 이후에서 <paramref name="query"/> 를 처음으로 발견하면
    /// 그 TextRange 를 반환. 없으면 null.
    /// </summary>
    public static TextRange? FindNext(
        FlowDocument doc,
        string query,
        TextPointer? startFrom,
        bool caseSensitive)
    {
        if (string.IsNullOrEmpty(query)) return null;

        var cmp = caseSensitive
            ? StringComparison.Ordinal
            : StringComparison.OrdinalIgnoreCase;

        var pos = startFrom ?? doc.ContentStart;

        while (pos != null && pos.CompareTo(doc.ContentEnd) < 0)
        {
            if (pos.GetPointerContext(LogicalDirection.Forward) == TextPointerContext.Text)
            {
                var run = pos.GetTextInRun(LogicalDirection.Forward);
                var idx = run.IndexOf(query, cmp);
                if (idx >= 0)
                {
                    var matchStart = pos.GetPositionAtOffset(idx);
                    var matchEnd   = matchStart?.GetPositionAtOffset(query.Length);
                    if (matchStart is not null && matchEnd is not null)
                        return new TextRange(matchStart, matchEnd);
                }
                // advance past this text run
                pos = pos.GetPositionAtOffset(run.Length)
                    ?? pos.GetNextContextPosition(LogicalDirection.Forward);
            }
            else
            {
                pos = pos.GetNextContextPosition(LogicalDirection.Forward);
            }
        }
        return null;
    }

    /// <summary>문서 전체에서 모두 바꾸기. 바꾼 횟수를 반환.</summary>
    public static int ReplaceAll(
        FlowDocument doc,
        string query,
        string replacement,
        bool caseSensitive)
    {
        if (string.IsNullOrEmpty(query)) return 0;

        int count = 0;
        TextPointer? pos = doc.ContentStart;
        while (true)
        {
            var found = FindNext(doc, query, pos, caseSensitive);
            if (found is null) break;

            // 다음 검색 시작점 — 교체 후 포인터가 무효화되므로 교체 이전에 저장.
            found.Text = replacement;
            count++;

            // 교체 후 포인터는 found.End 가 바뀌므로 ContentStart 부터 재시도 (중복 방지를 위해 교체분 skip).
            // 단순 구현: 항상 ContentStart 에서 재시작. 성능보다 정확성 우선.
            pos = doc.ContentStart;
            if (string.IsNullOrEmpty(replacement)) continue;

            // 교체문자열 자체가 검색어를 포함하면 무한루프가 될 수 있으므로
            // 교체 후 이전에 찾은 위치의 end 부터 재시작.
            // TextRange.End 는 교체 이후에도 유효한 경우가 많다.
            pos = found.End;
        }
        return count;
    }
}

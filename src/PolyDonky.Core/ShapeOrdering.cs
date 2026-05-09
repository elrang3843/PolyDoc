namespace PolyDonky.Core;

/// <summary>
/// <see cref="ShapeObject"/> 들의 효과적 그리기 순서(z-order)를 계산한다.
///
/// 정렬 정책:
/// <list type="number">
///   <item><c>ZOrder &lt; 0</c> — 자동 그룹보다 항상 뒤. ZOrder 오름차순 + stable doc 순서.</item>
///   <item><c>ZOrder == 0</c> (자동) — 한 도형이 다른 도형의 바운딩 박스를 완전히 포함하면
///   안쪽(포함되는) 도형이 더 뒤에 그려져 결과적으로 위에 표시된다. 포함 관계가 없으면 문서 순서 유지.</item>
///   <item><c>ZOrder &gt; 0</c> — 자동 그룹보다 항상 앞. ZOrder 오름차순 + stable doc 순서.</item>
/// </list>
///
/// 컨테인먼트 검사는 절대 좌표를 가진 도형(<c>WrapMode = InFrontOfText / BehindText</c>) 만 비교한다.
/// 그 외(인라인·텍스트랩)는 절대 좌표가 없어 컨테인먼트가 정의되지 않으므로 자동 보정에서 제외 — 문서 순서를 유지한다.
/// </summary>
public static class ShapeOrdering
{
    /// <summary>입력 도형 시퀀스를 위 정책에 따라 그리기 순서로 재배열한 새 리스트를 반환한다.</summary>
    public static IReadOnlyList<ShapeObject> OrderForRendering(IEnumerable<ShapeObject> shapes)
    {
        ArgumentNullException.ThrowIfNull(shapes);
        var arr = shapes as ShapeObject[] ?? shapes.ToArray();
        if (arr.Length < 2) return arr.ToList();

        var keys = new (int group, int autoDepth)[arr.Length];
        for (int i = 0; i < arr.Length; i++)
        {
            var s = arr[i];
            int group = s.ZOrder;
            int depth = (s.ZOrder == 0) ? ComputeContainmentDepth(s, arr) : 0;
            keys[i] = (group, depth);
        }

        var indices = new int[arr.Length];
        for (int i = 0; i < indices.Length; i++) indices[i] = i;

        // (group, autoDepth, originalIndex) 사전식 정렬 — Array.Sort 는 stable 하지 않지만
        // tie-breaker 로 originalIndex 를 명시해 의도된 순서를 보장.
        Array.Sort(indices, (a, b) =>
        {
            int c = keys[a].group.CompareTo(keys[b].group);
            if (c != 0) return c;
            c = keys[a].autoDepth.CompareTo(keys[b].autoDepth);
            if (c != 0) return c;
            return a.CompareTo(b);
        });

        var result = new ShapeObject[arr.Length];
        for (int i = 0; i < indices.Length; i++) result[i] = arr[indices[i]];
        return result;
    }

    /// <summary>
    /// CSS <c>z-index</c> 와 같은 정수 값을 각 도형에 대해 계산해 반환한다.
    /// 같은 캔버스/스택 컨텍스트 안에서 이 값이 큰 도형이 위에 그려진다.
    /// 의미:
    /// <list type="bullet">
    ///   <item>명시적 <c>ZOrder != 0</c> 도형 — 그 값을 그대로 사용.</item>
    ///   <item><c>ZOrder == 0</c> 자동 그룹 — 컨테인먼트 깊이를 z-index 로 사용
    ///   (다른 자동 그룹 도형이 자기를 포함할수록 깊이 ↑ → 위에 보임).
    ///   포함 관계가 없으면 <c>0</c>.</item>
    /// </list>
    /// 결과 dictionary 에는 자동 그룹이고 깊이가 0 인 도형도 포함되며 값은 <c>0</c> 이다 — 호출 측에서
    /// "<c>0</c> 이면 z-index 출력 생략" 같은 정책을 자유롭게 적용할 수 있다.
    /// </summary>
    public static IReadOnlyDictionary<ShapeObject, int> ComputeZIndexMap(IEnumerable<ShapeObject> shapes)
    {
        ArgumentNullException.ThrowIfNull(shapes);
        var arr = shapes as ShapeObject[] ?? shapes.ToArray();
        var map = new Dictionary<ShapeObject, int>(arr.Length);
        foreach (var s in arr)
        {
            int z = (s.ZOrder != 0) ? s.ZOrder : ComputeContainmentDepth(s, arr);
            map[s] = z;
        }
        return map;
    }

    private static int ComputeContainmentDepth(ShapeObject s, ShapeObject[] all)
    {
        if (!TryGetAbsoluteBBox(s, out var sBox)) return 0;
        int depth = 0;
        foreach (var other in all)
        {
            if (ReferenceEquals(other, s)) continue;
            if (other.ZOrder != 0) continue;                      // 자동 그룹끼리만 비교
            if (!TryGetAbsoluteBBox(other, out var oBox)) continue;
            if (BBoxStrictlyContains(oBox, sBox)) depth++;
        }
        return depth;
    }

    /// <summary>절대 좌표가 정의된 도형(오버레이) 만 bbox 반환. 그 외엔 false.</summary>
    private static bool TryGetAbsoluteBBox(ShapeObject s, out BBox bbox)
    {
        if (s.WrapMode is ImageWrapMode.InFrontOfText or ImageWrapMode.BehindText
            && s.WidthMm > 0 && s.HeightMm > 0)
        {
            bbox = new BBox(s.OverlayXMm, s.OverlayYMm, s.WidthMm, s.HeightMm);
            return true;
        }
        bbox = default;
        return false;
    }

    /// <summary><paramref name="outer"/> 가 <paramref name="inner"/> 를 진성(strict) 으로 포함하는지.
    /// 동일 bbox 는 포함 관계 아님 — 컨테인먼트 깊이 무한 루프 방지.</summary>
    private static bool BBoxStrictlyContains(in BBox outer, in BBox inner)
    {
        const double eps = 0.01;
        bool contains = inner.X >= outer.X - eps
                     && inner.Y >= outer.Y - eps
                     && inner.X + inner.W <= outer.X + outer.W + eps
                     && inner.Y + inner.H <= outer.Y + outer.H + eps;
        if (!contains) return false;
        // strict: 면적이 동일하면 컨테인먼트로 보지 않는다.
        return outer.W * outer.H > inner.W * inner.H + eps;
    }

    private readonly record struct BBox(double X, double Y, double W, double H);
}

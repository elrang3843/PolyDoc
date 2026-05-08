using PolyDonky.Core;

namespace PolyDonky.Core.Tests;

public class ShapeOrderingTests
{
    private static ShapeObject Overlay(int order, double x, double y, double w, double h, string? id = null)
        => new()
        {
            Id = id, ZOrder = order,
            WrapMode = ImageWrapMode.InFrontOfText,
            OverlayXMm = x, OverlayYMm = y, WidthMm = w, HeightMm = h,
        };

    private static ShapeObject Inline(int order = 0, string? id = null)
        => new()
        {
            Id = id, ZOrder = order,
            WrapMode = ImageWrapMode.Inline,
            WidthMm = 30, HeightMm = 20,
        };

    [Fact]
    public void Empty_Returns_Empty()
    {
        Assert.Empty(ShapeOrdering.OrderForRendering(Array.Empty<ShapeObject>()));
    }

    [Fact]
    public void Single_Returns_Same()
    {
        var s = Inline();
        var ordered = ShapeOrdering.OrderForRendering(new[] { s });
        Assert.Single(ordered);
        Assert.Same(s, ordered[0]);
    }

    [Fact]
    public void NoZOrder_NoContainment_PreservesDocOrder()
    {
        // 두 오버레이 도형이 겹치지 않을 때는 문서 순서 유지.
        var a = Overlay(0,  0,  0, 50, 50, "A");
        var b = Overlay(0, 60, 60, 50, 50, "B");
        var ordered = ShapeOrdering.OrderForRendering(new[] { a, b });
        Assert.Equal(new[] { "A", "B" }, ordered.Select(s => s.Id).ToArray());
    }

    [Fact]
    public void NoZOrder_OuterContainsInner_InnerDrawnLast()
    {
        // 큰 도형이 작은 도형을 포함 → 작은 도형이 뒤에 그려져 위에 보임.
        var outer = Overlay(0, 0, 0, 100, 100, "outer");
        var inner = Overlay(0, 20, 20, 30, 30, "inner");
        // 문서 순서: inner 가 먼저, outer 가 나중. 자동 보정으로 inner 가 뒤로 이동.
        var ordered = ShapeOrdering.OrderForRendering(new[] { inner, outer });
        Assert.Equal(new[] { "outer", "inner" }, ordered.Select(s => s.Id).ToArray());
    }

    [Fact]
    public void NoZOrder_TripleNesting_DeeperShapesDrawnLater()
    {
        var big    = Overlay(0,  0,  0, 100, 100, "big");
        var mid    = Overlay(0, 10, 10,  60,  60, "mid");
        var small  = Overlay(0, 20, 20,  20,  20, "small");
        // 문서 순서를 일부러 뒤집어 입력.
        var ordered = ShapeOrdering.OrderForRendering(new[] { small, mid, big });
        Assert.Equal(new[] { "big", "mid", "small" }, ordered.Select(s => s.Id).ToArray());
    }

    [Fact]
    public void ExplicitNegativeZOrder_AlwaysBehindAuto()
    {
        var bg     = Overlay(-10, 0, 0, 50, 50, "bg");
        var auto   = Overlay(  0, 0, 0, 50, 50, "auto");
        var ordered = ShapeOrdering.OrderForRendering(new[] { auto, bg });
        Assert.Equal(new[] { "bg", "auto" }, ordered.Select(s => s.Id).ToArray());
    }

    [Fact]
    public void ExplicitPositiveZOrder_AlwaysAboveAuto()
    {
        var auto   = Overlay(0, 0, 0, 50, 50, "auto");
        var fg     = Overlay(5, 0, 0, 50, 50, "fg");
        var ordered = ShapeOrdering.OrderForRendering(new[] { fg, auto });
        Assert.Equal(new[] { "auto", "fg" }, ordered.Select(s => s.Id).ToArray());
    }

    [Fact]
    public void MixedZOrder_SortedAscending()
    {
        var a = Overlay(-3, 0, 0, 10, 10, "a");
        var b = Overlay(-1, 0, 0, 10, 10, "b");
        var c = Overlay( 0, 0, 0, 10, 10, "c");
        var d = Overlay( 2, 0, 0, 10, 10, "d");
        var e = Overlay( 7, 0, 0, 10, 10, "e");
        var ordered = ShapeOrdering.OrderForRendering(new[] { e, c, a, d, b });
        Assert.Equal(new[] { "a", "b", "c", "d", "e" }, ordered.Select(s => s.Id).ToArray());
    }

    [Fact]
    public void SameZOrder_PreservesDocOrder()
    {
        var a = Overlay(2, 0, 0, 10, 10, "a");
        var b = Overlay(2, 0, 0, 10, 10, "b");
        var c = Overlay(2, 0, 0, 10, 10, "c");
        var ordered = ShapeOrdering.OrderForRendering(new[] { a, b, c });
        Assert.Equal(new[] { "a", "b", "c" }, ordered.Select(s => s.Id).ToArray());
    }

    [Fact]
    public void InlineShapes_NoBBoxComparison_DocOrderOnly()
    {
        // 인라인 도형은 절대 좌표가 없어 컨테인먼트 비교 제외 — 문서 순서 유지.
        var big   = Inline(0, "big");   big.WidthMm = 100; big.HeightMm = 100;
        var small = Inline(0, "small"); small.WidthMm =  30; small.HeightMm =  30;
        var ordered = ShapeOrdering.OrderForRendering(new[] { big, small });
        Assert.Equal(new[] { "big", "small" }, ordered.Select(s => s.Id).ToArray());
    }

    [Fact]
    public void IdenticalBBox_NotConsideredContainment()
    {
        // 동일한 bbox 두 도형 → 컨테인먼트 아님 → 문서 순서 유지.
        var a = Overlay(0, 10, 10, 50, 50, "a");
        var b = Overlay(0, 10, 10, 50, 50, "b");
        var ordered = ShapeOrdering.OrderForRendering(new[] { a, b });
        Assert.Equal(new[] { "a", "b" }, ordered.Select(s => s.Id).ToArray());
    }

    [Fact]
    public void OverlappingButNotContaining_PreservesDocOrder()
    {
        // 겹치지만 포함 관계 아님 → 자동 보정 안 일어남 → 문서 순서.
        var a = Overlay(0,  0,  0, 50, 50, "a");
        var b = Overlay(0, 30, 30, 50, 50, "b");
        var ordered = ShapeOrdering.OrderForRendering(new[] { a, b });
        Assert.Equal(new[] { "a", "b" }, ordered.Select(s => s.Id).ToArray());
    }

    [Fact]
    public void DefaultZOrder_IsZero()
    {
        Assert.Equal(0, new ShapeObject().ZOrder);
    }
}

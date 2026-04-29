using System;
using System.Linq;
using System.Threading;
using System.Windows;
using PolyDonky.App.Pagination;
using PolyDonky.Core;

namespace PolyDonky.App.Tests;

// ── 모델 구조 테스트 (WPF·STA 불필요) ─────────────────────────────────────

public class PerPageDocumentSliceModelTests
{
    [Fact]
    public void PerPageDocumentSlice_DefaultBodyBlocks_IsEmpty()
    {
        var slice = new PerPageDocumentSlice();
        Assert.Empty(slice.BodyBlocks);
    }

    [Fact]
    public void PerPageDocumentSlice_PageIndex_DefaultsToZero()
    {
        var slice = new PerPageDocumentSlice();
        Assert.Equal(0, slice.PageIndex);
    }

    [Fact]
    public void PerPageDocumentSlice_BodyDimensions_DefaultToZero()
    {
        var slice = new PerPageDocumentSlice();
        Assert.Equal(0.0, slice.BodyWidthDip);
        Assert.Equal(0.0, slice.BodyHeightDip);
    }
}

// ── PerPageDocumentSplitter 통합 테스트 (STA 필수) ──────────────────────────

public class PerPageDocumentSplitterTests
{
    private static void RunOnSta(Action action)
    {
        Exception? caught = null;
        var t = new Thread(() =>
        {
            try { action(); }
            catch (Exception ex) { caught = ex; }
        });
        t.SetApartmentState(ApartmentState.STA);
        t.Start();
        t.Join();
        if (caught is not null)
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(caught).Throw();
    }

    // ── 기본 구조 ──────────────────────────────────────────────────────────

    [Fact]
    public void Split_EmptyDocument_ReturnsAtLeastOneSlice()
    {
        RunOnSta(() =>
        {
            var paginated = FlowDocumentPaginationAdapter.Paginate(PolyDonkyument.Empty());
            var slices    = PerPageDocumentSplitter.Split(paginated);

            Assert.True(slices.Count >= 1);
        });
    }

    [Fact]
    public void Split_SliceCount_MatchesPaginatedPageCount()
    {
        RunOnSta(() =>
        {
            var doc    = PolyDonkyument.Empty();
            var paginated = FlowDocumentPaginationAdapter.Paginate(doc);
            var slices    = PerPageDocumentSplitter.Split(paginated);

            Assert.Equal(paginated.PageCount, slices.Count);
        });
    }

    [Fact]
    public void Split_SlicePageIndices_AreConsecutiveFromZero()
    {
        RunOnSta(() =>
        {
            var doc       = BuildMultiParagraphDocument(3);
            var paginated = FlowDocumentPaginationAdapter.Paginate(doc);
            var slices    = PerPageDocumentSplitter.Split(paginated);

            for (int i = 0; i < slices.Count; i++)
                Assert.Equal(i, slices[i].PageIndex);
        });
    }

    [Fact]
    public void Split_EachSlice_HasNonNullFlowDocument()
    {
        RunOnSta(() =>
        {
            var doc       = BuildMultiParagraphDocument(2);
            var paginated = FlowDocumentPaginationAdapter.Paginate(doc);
            var slices    = PerPageDocumentSplitter.Split(paginated);

            foreach (var slice in slices)
                Assert.NotNull(slice.FlowDocument);
        });
    }

    [Fact]
    public void Split_BodyDimensions_ArePositive()
    {
        RunOnSta(() =>
        {
            var doc       = PolyDonkyument.Empty();
            var paginated = FlowDocumentPaginationAdapter.Paginate(doc);
            var slices    = PerPageDocumentSplitter.Split(paginated);

            foreach (var slice in slices)
            {
                Assert.True(slice.BodyWidthDip  > 0, $"BodyWidthDip={slice.BodyWidthDip}");
                Assert.True(slice.BodyHeightDip > 0, $"BodyHeightDip={slice.BodyHeightDip}");
            }
        });
    }

    [Fact]
    public void Split_PageSettings_MatchesPaginatedDocument()
    {
        RunOnSta(() =>
        {
            var page = new PageSettings { SizeKind = PaperSizeKind.A4 };
            var doc  = new PolyDonkyument();
            doc.Sections.Add(new Section { Page = page });

            var paginated = FlowDocumentPaginationAdapter.Paginate(doc);
            var slices    = PerPageDocumentSplitter.Split(paginated);

            foreach (var slice in slices)
                Assert.Same(page, slice.PageSettings);
        });
    }

    // ── 블록 배정 ─────────────────────────────────────────────────────────

    [Fact]
    public void Split_SingleParagraph_AppearsInFirstSlice()
    {
        RunOnSta(() =>
        {
            var para = new Paragraph();
            para.AddText("테스트 단락");
            var doc = WrapInDocument(para);

            var paginated = FlowDocumentPaginationAdapter.Paginate(doc);
            var slices    = PerPageDocumentSplitter.Split(paginated);

            var allBlocks = slices.SelectMany(s => s.BodyBlocks).ToList();
            Assert.Contains(allBlocks, bop => ReferenceEquals(bop.Source, para));
        });
    }

    [Fact]
    public void Split_TotalBodyBlockCount_MatchesPaginatedTotal()
    {
        RunOnSta(() =>
        {
            var doc       = BuildMultiParagraphDocument(4);
            var paginated = FlowDocumentPaginationAdapter.Paginate(doc);
            var slices    = PerPageDocumentSplitter.Split(paginated);

            int paginatedTotal = paginated.Pages.Sum(p => p.BodyBlocks.Count);
            int sliceTotal     = slices.Sum(s => s.BodyBlocks.Count);

            Assert.Equal(paginatedTotal, sliceTotal);
        });
    }

    [Fact]
    public void Split_SliceBodyBlocks_MatchPaginatedPageBodyBlocks()
    {
        RunOnSta(() =>
        {
            var doc       = BuildMultiParagraphDocument(3);
            var paginated = FlowDocumentPaginationAdapter.Paginate(doc);
            var slices    = PerPageDocumentSplitter.Split(paginated);

            for (int i = 0; i < paginated.PageCount; i++)
            {
                var expected = paginated.Pages[i].BodyBlocks;
                var actual   = slices[i].BodyBlocks;
                Assert.Equal(expected.Count, actual.Count);
            }
        });
    }

    // ── FlowDocument 내용 검증 ──────────────────────────────────────────────

    [Fact]
    public void Split_SliceFlowDocument_PageWidthIsPositive()
    {
        RunOnSta(() =>
        {
            var doc       = PolyDonkyument.Empty();
            var paginated = FlowDocumentPaginationAdapter.Paginate(doc);
            var slices    = PerPageDocumentSplitter.Split(paginated);

            foreach (var slice in slices)
                Assert.True(slice.FlowDocument.PageWidth > 0,
                    $"PageWidth={slice.FlowDocument.PageWidth}");
        });
    }

    [Fact]
    public void Split_NullDocument_Throws()
    {
        Assert.Throws<ArgumentNullException>(() =>
            PerPageDocumentSplitter.Split(null!));
    }

    // ── 헬퍼 ─────────────────────────────────────────────────────────────

    private static PolyDonkyument WrapInDocument(params Block[] blocks)
    {
        var doc     = new PolyDonkyument();
        var section = new Section();
        foreach (var b in blocks) section.Blocks.Add(b);
        doc.Sections.Add(section);
        return doc;
    }

    private static PolyDonkyument BuildMultiParagraphDocument(int count)
    {
        var doc     = new PolyDonkyument();
        var section = new Section();
        for (int i = 0; i < count; i++)
        {
            var p = new Paragraph();
            p.AddText($"단락 {i + 1}");
            section.Blocks.Add(p);
        }
        doc.Sections.Add(section);
        return doc;
    }
}

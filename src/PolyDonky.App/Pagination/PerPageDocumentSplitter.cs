using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using PolyDonky.App.Services;
using PolyDonky.Core;

// WPF 타입과의 이름 충돌 회피
using CoreBlock          = PolyDonky.Core.Block;
using CoreParagraph      = PolyDonky.Core.Paragraph;
using CoreContainerBlock = PolyDonky.Core.ContainerBlock;
using CoreTable          = PolyDonky.Core.Table;
using CoreTextBox        = PolyDonky.Core.TextBoxObject;
using WpfRun             = System.Windows.Documents.Run;
using WpfParagraph       = System.Windows.Documents.Paragraph;

namespace PolyDonky.App.Pagination;

/// <summary>
/// <see cref="PaginatedDocument"/> 의 블록 배정 정보를 기반으로 페이지별·단별
/// <see cref="PerPageDocumentSlice"/> 를 생성한다.
/// <para>
/// <b>STA 스레드에서 호출해야 한다</b> — 내부에서 WPF <c>FlowDocument</c> 를 생성한다.
/// </para>
/// <para>
/// 다단 문서의 경우 페이지당 <c>ColumnCount</c> 개의 슬라이스를 생성한다.
/// 슬라이스 순서: (page 0, col 0), (page 0, col 1), …, (page 1, col 0), …
/// 단일 단이면 기존과 동일하게 페이지당 1개.
/// </para>
/// </summary>
public static class PerPageDocumentSplitter
{
    /// <summary>
    /// <paramref name="paginated"/> 에서 페이지·단별 슬라이스 목록을 생성한다.
    /// </summary>
    /// <param name="paginated">분할할 페이지네이션 결과.</param>
    /// <param name="outlineStyles">
    /// 개요 서식 재정의. <c>null</c> 이면 원본 문서의 서식을 사용하고
    /// 그것도 없으면 기본값을 쓴다.
    /// </param>
    public static IReadOnlyList<PerPageDocumentSlice> Split(
        PaginatedDocument paginated,
        OutlineStyleSet?  outlineStyles = null)
    {
        ArgumentNullException.ThrowIfNull(paginated);

        var styles = outlineStyles
            ?? paginated.Source.OutlineStyles
            ?? OutlineStyleSet.CreateDefault();

        // 각주/미주 번호 맵 — 문서 전체 순서 기준으로 번호 부여.
        var fnNums = paginated.Source.Footnotes.Count > 0
            ? paginated.Source.Footnotes.Select((f, i) => (f.Id, Num: i + 1))
                .ToDictionary(x => x.Id, x => x.Num)
            : null;
        var enNums = paginated.Source.Endnotes.Count > 0
            ? paginated.Source.Endnotes.Select((e, i) => (e.Id, Num: i + 1))
                .ToDictionary(x => x.Id, x => x.Num)
            : null;

        // 원본 문서의 ContainerBlock 계층 복원에 사용할 부모 맵.
        var parentMap      = BuildParentMap(paginated.Source);
        var outlineNumbers = FlowDocumentBuilder.ComputeOutlineNumbers(paginated.Source, styles);

        // 블록 → 섹션 인덱스 맵 — 각 슬라이스의 SectionIndex 결정에 사용.
        var blockSectionMap = BuildBlockToSectionIndexMap(paginated.Source);

        var slices = new List<PerPageDocumentSlice>(paginated.PageCount);

        for (int pageIdx = 0; pageIdx < paginated.PageCount; pageIdx++)
        {
            var pp   = paginated.Pages[pageIdx];
            var page = paginated.GetPageSettings(pageIdx);
            var geo  = new PageGeometry(page);

            int    colCount = geo.ColumnCount;
            double bodyH    = Math.Max(1.0, geo.PageHeightDip - geo.PadTopDip - geo.PadBottomDip);

            // 이 페이지의 첫 본문 블록이 속한 섹션 인덱스
            int sectionIdx = 0;
            foreach (var bop in pp.BodyBlocks)
            {
                if (blockSectionMap.TryGetValue(bop.Source, out var si))
                { sectionIdx = si; break; }
            }

            for (int col = 0; col < colCount; col++)
            {
                double colWidth = col < geo.ColWidthsDip.Length ? geo.ColWidthsDip[col] : geo.ColWidthDip;

                var colBlocks  = pp.BodyBlocks.Where(b => b.ColumnIndex == col).ToList();
                var rawBlocks  = colBlocks.Select(b => b.Source).ToList();

                // ForcePageBreakBefore 는 섹션 경계 마커일 뿐 — per-page RTB 의 첫 블록에
                // WPF BreakPageBefore 를 적용하면 RTB 상단에 공백이 생기므로 잠시 억제한다.
                CoreParagraph? firstBreakPara = rawBlocks.Count > 0
                    && rawBlocks[0] is CoreParagraph fbp && fbp.Style.ForcePageBreakBefore
                    ? fbp : null;
                if (firstBreakPara is not null)
                    firstBreakPara.Style.ForcePageBreakBefore = false;

                // ContainerBlock 자식들을 다시 부모로 감싸 배경·보더·마진을 복원한다.
                var coreBlocks = parentMap.Count > 0
                    ? ReassembleContainerBlocks(rawBlocks, parentMap)
                    : rawBlocks;
                var fd = FlowDocumentBuilder.BuildFromBlocks(coreBlocks, page, styles, outlineNumbers, fnNums, enNums);

                if (firstBreakPara is not null)
                    firstBreakPara.Style.ForcePageBreakBefore = true;

                // 이 페이지/단에 나타나는 각주 수집.
                var pageFootnotes = CollectPageFootnotes(coreBlocks, paginated.Source);

                // 각주 영역 높이 측정 — 각주가 있으면 본문 RTB 높이를 줄인다.
                // 각주는 본문 RTB 바로 아래, 꼬리말 위에 배치된다.
                double footnoteAreaH = pageFootnotes.Count > 0
                    ? MeasureFootnoteAreaHeight(pageFootnotes, fnNums, colWidth)
                    : 0.0;

                double effectiveBodyH = Math.Max(20.0, bodyH - footnoteAreaH);

                // per-column RTB는 단 폭만 담당; 여백·단 오프셋은 PerPageEditorHost 가 위치로 처리.
                fd.PageWidth   = colWidth;
                fd.PagePadding = new Thickness(0);

                slices.Add(new PerPageDocumentSlice
                {
                    PageIndex           = pageIdx,
                    ColumnIndex         = col,
                    ColumnCount         = colCount,
                    SectionIndex        = sectionIdx,
                    XOffsetDip          = geo.ColumnXOffsetDip(col),
                    PageSettings        = page,
                    BodyBlocks          = colBlocks,
                    FlowDocument        = fd,
                    BodyWidthDip        = colWidth,
                    BodyHeightDip       = effectiveBodyH,
                    FootnoteAreaHeightDip = footnoteAreaH,
                    PageFootnotes       = pageFootnotes,
                });
            }
        }

        // 미주가 있으면 문서 끝에 미주 전용 페이지 슬라이스를 추가한다.
        if (paginated.Source.Endnotes.Count > 0)
        {
            var lastPage = paginated.PageCount > 0
                ? paginated.GetPageSettings(paginated.PageCount - 1)
                : new PageSettings();
            var endnotePage = BuildEndnotePage(
                paginated.Source.Endnotes,
                enNums,
                lastPage,
                pageIndex: paginated.PageCount);
            slices.Add(endnotePage);
        }

        return slices;
    }

    // ── 각주 수집 ─────────────────────────────────────────────────────────────

    /// <summary>블록 목록을 재귀 순회해 이 페이지에 나타나는 각주 목록을 반환한다.</summary>
    private static IReadOnlyList<FootnoteEntry> CollectPageFootnotes(
        List<CoreBlock> pageBlocks,
        PolyDonkyument doc)
    {
        if (doc.Footnotes.Count == 0) return Array.Empty<FootnoteEntry>();

        var fnIds = new Dictionary<string, int>();
        CollectFootnoteIds(pageBlocks, fnIds);

        return fnIds.OrderBy(x => x.Value)
            .Select(x => doc.Footnotes.FirstOrDefault(f => f.Id == x.Key))
            .Where(f => f is not null)
            .Cast<FootnoteEntry>()
            .ToList();
    }

    private static void CollectFootnoteIds(IList<CoreBlock> blocks, Dictionary<string, int> fnIds, int seed = 0)
    {
        int order = seed;
        foreach (var block in blocks)
        {
            if (block is CoreParagraph para)
            {
                foreach (var run in para.Runs)
                    if (!string.IsNullOrEmpty(run.FootnoteId) && !fnIds.ContainsKey(run.FootnoteId))
                        fnIds[run.FootnoteId] = order++;
            }
            else if (block is CoreTable table)
            {
                foreach (var row in table.Rows)
                    foreach (var cell in row.Cells)
                        CollectFootnoteIds(cell.Blocks, fnIds, order);
            }
            else if (block is CoreTextBox textbox)
                CollectFootnoteIds(textbox.Content, fnIds, order);
            else if (block is CoreContainerBlock container)
                CollectFootnoteIds(container.Children, fnIds, order);
        }
    }

    // ── 각주 영역 높이 측정 ───────────────────────────────────────────────────

    /// <summary>
    /// 각주 FlowDocument 를 오프스크린으로 Measure() 해 정확한 각주 영역 높이를 반환한다.
    /// STA 스레드 전용.
    /// </summary>
    private static double MeasureFootnoteAreaHeight(
        IReadOnlyList<FootnoteEntry> footnotes,
        IReadOnlyDictionary<string, int>? fnNums,
        double colWidth)
    {
        const double SeparatorH = 1 + 4 + 4; // 1px 선 + 위아래 여백

        var fd = BuildFootnoteFlowDocument(footnotes, fnNums, colWidth);

        var tempRtb = new RichTextBox
        {
            Document                      = fd,
            Width                         = colWidth,
            Padding                       = new Thickness(0),
            BorderThickness               = new Thickness(0),
            VerticalScrollBarVisibility   = ScrollBarVisibility.Disabled,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
        };
        tempRtb.Measure(new Size(colWidth, double.PositiveInfinity));

        return SeparatorH + tempRtb.DesiredSize.Height;
    }

    /// <summary>각주 표시용 FlowDocument — 오프스크린 측정과 실제 렌더링에 공용으로 사용.</summary>
    internal static FlowDocument BuildFootnoteFlowDocument(
        IReadOnlyList<FootnoteEntry> footnotes,
        IReadOnlyDictionary<string, int>? fnNums,
        double width)
    {
        var fd = new FlowDocument
        {
            PageWidth   = width,
            PagePadding = new Thickness(0),
        };

        int autoIdx = 0;
        foreach (var note in footnotes)
        {
            autoIdx++;
            int num = (fnNums != null && fnNums.TryGetValue(note.Id, out var n)) ? n : autoIdx;

            foreach (var coreBlock in note.Blocks.OfType<CoreParagraph>())
            {
                var wpfPara = new WpfParagraph
                {
                    FontSize = 10,
                    Margin   = new Thickness(0, 0, 0, 2),
                };
                wpfPara.Inlines.Add(new WpfRun($"{num} ") { FontWeight = FontWeights.Bold });
                foreach (var coreRun in coreBlock.Runs)
                    wpfPara.Inlines.Add(new WpfRun(coreRun.Text ?? string.Empty));
                fd.Blocks.Add(wpfPara);
            }
        }
        return fd;
    }

    // ── 미주 전용 페이지 ──────────────────────────────────────────────────────

    private static PerPageDocumentSlice BuildEndnotePage(
        IList<FootnoteEntry> endnotes,
        IReadOnlyDictionary<string, int>? enNums,
        PageSettings page,
        int pageIndex)
    {
        var geo      = new PageGeometry(page);
        double bodyW = geo.ColWidthsDip[0];
        double bodyH = Math.Max(1.0, geo.PageHeightDip - geo.PadTopDip - geo.PadBottomDip);

        var fd = new FlowDocument
        {
            PageWidth   = bodyW,
            PagePadding = new Thickness(0),
        };

        // 미주 페이지 제목
        var title = new WpfParagraph
        {
            FontSize   = 14,
            FontWeight = FontWeights.Bold,
            Margin     = new Thickness(0, 0, 0, 6),
        };
        title.Inlines.Add(new WpfRun("미주"));
        fd.Blocks.Add(title);

        // 구분선 역할의 빈 단락 (하단 보더)
        var hr = new WpfParagraph
        {
            Margin          = new Thickness(0, 0, 0, 8),
            BorderBrush     = new SolidColorBrush(System.Windows.Media.Color.FromRgb(180, 180, 180)),
            BorderThickness = new Thickness(0, 0, 0, 1),
        };
        fd.Blocks.Add(hr);

        // 미주 본문
        int autoIdx = 0;
        foreach (var note in endnotes)
        {
            autoIdx++;
            int num = (enNums != null && enNums.TryGetValue(note.Id, out var n)) ? n : autoIdx;

            foreach (var coreBlock in note.Blocks.OfType<CoreParagraph>())
            {
                var wpfPara = new WpfParagraph
                {
                    FontSize = 11,
                    Margin   = new Thickness(0, 0, 0, 6),
                };
                wpfPara.Inlines.Add(new WpfRun($"{num}. ") { FontWeight = FontWeights.Bold });
                foreach (var coreRun in coreBlock.Runs)
                    wpfPara.Inlines.Add(new WpfRun(coreRun.Text ?? string.Empty));
                fd.Blocks.Add(wpfPara);
            }
        }

        return new PerPageDocumentSlice
        {
            PageIndex    = pageIndex,
            ColumnIndex  = 0,
            ColumnCount  = 1,
            SectionIndex = 0,
            XOffsetDip   = 0,
            PageSettings = page,
            BodyBlocks   = Array.Empty<BlockOnPage>(),
            FlowDocument = fd,
            BodyWidthDip = bodyW,
            BodyHeightDip = bodyH,
            IsEndnotePage = true,
        };
    }

    // ── 문서 모델 헬퍼 ────────────────────────────────────────────────────────

    /// <summary>문서의 모든 섹션을 순회해 최상위 블록 → 섹션 인덱스 맵을 만든다.</summary>
    private static Dictionary<CoreBlock, int> BuildBlockToSectionIndexMap(PolyDonkyument doc)
    {
        int totalBlocks = doc.Sections.Sum(s => s.Blocks.Count);
        var map = new Dictionary<CoreBlock, int>(totalBlocks, ReferenceEqualityComparer.Instance);
        for (int si = 0; si < doc.Sections.Count; si++)
            foreach (var block in doc.Sections[si].Blocks)
                map[block] = si;
        return map;
    }

    /// <summary>원본 문서를 순회해 ContainerBlock 의 직접 자식 → 부모 ContainerBlock 맵을 만든다.</summary>
    private static Dictionary<CoreBlock, CoreContainerBlock> BuildParentMap(PolyDonkyument doc)
    {
        int totalBlocks = doc.Sections.Sum(s => s.Blocks.Count);
        var map = new Dictionary<CoreBlock, CoreContainerBlock>(totalBlocks, ReferenceEqualityComparer.Instance);
        foreach (var section in doc.Sections)
            CollectParents(section.Blocks, map);
        return map;
    }

    private static void CollectParents(IEnumerable<CoreBlock> blocks, Dictionary<CoreBlock, CoreContainerBlock> map)
    {
        foreach (var block in blocks)
        {
            if (block is not CoreContainerBlock container) continue;
            foreach (var child in container.Children)
                map[child] = container;
            CollectParents(container.Children, map);
        }
    }

    /// <summary>
    /// 평탄화된 블록 목록에서 같은 ContainerBlock 부모를 가진 연속 블록을
    /// 새 ContainerBlock 으로 다시 감싼다 — 배경·보더·패딩 복원.
    /// </summary>
    private static List<CoreBlock> ReassembleContainerBlocks(
        List<CoreBlock> flat,
        Dictionary<CoreBlock, CoreContainerBlock> parentMap)
    {
        if (flat.Count == 0) return flat;

        var result = new List<CoreBlock>(flat.Count);
        int i = 0;
        while (i < flat.Count)
        {
            var block = flat[i];
            if (!parentMap.TryGetValue(block, out var parent))
            {
                result.Add(block);
                i++;
                continue;
            }

            int j = i + 1;
            while (j < flat.Count &&
                   parentMap.TryGetValue(flat[j], out var p) &&
                   ReferenceEquals(p, parent))
                j++;

            var wrapped = new CoreContainerBlock
            {
                BorderTopPt       = parent.BorderTopPt,       BorderTopColor    = parent.BorderTopColor,
                BorderRightPt     = parent.BorderRightPt,     BorderRightColor  = parent.BorderRightColor,
                BorderBottomPt    = parent.BorderBottomPt,    BorderBottomColor = parent.BorderBottomColor,
                BorderLeftPt      = parent.BorderLeftPt,      BorderLeftColor   = parent.BorderLeftColor,
                BackgroundColor   = parent.BackgroundColor,
                PaddingTopMm      = parent.PaddingTopMm,      PaddingRightMm    = parent.PaddingRightMm,
                PaddingBottomMm   = parent.PaddingBottomMm,   PaddingLeftMm     = parent.PaddingLeftMm,
                MarginTopMm       = parent.MarginTopMm,       MarginBottomMm    = parent.MarginBottomMm,
                WidthMm           = parent.WidthMm,           HAlign            = parent.HAlign,
                ClassNames        = parent.ClassNames,        Role              = parent.Role,
                Children          = flat.GetRange(i, j - i),
            };
            result.Add(wrapped);
            i = j;
        }
        return result;
    }
}

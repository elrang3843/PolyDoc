using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;

namespace PolyDonky.App.Pagination;

/// <summary>
/// 오프스크린 측정과 per-page 렌더 높이 사이의 오차로 발생하는 페이지 오버플로를
/// 슬라이스 수준에서 후처리로 보정한다.
/// <para>
/// <b>STA 스레드 전용</b> — WPF RichTextBox / FlowDocument 에 의존.
/// </para>
/// <para>
/// 동작 원리:
/// <list type="bullet">
///   <item>각 슬라이스의 FlowDocument 를 오프스크린 RTB 에 붙여 <c>Measure(∞)</c> 로 실제 높이 계산.</item>
///   <item>실제 높이가 <c>BodyHeightDip</c> 를 초과하면 마지막 WPF Block 을 다음 슬라이스 앞으로 이동.</item>
///   <item>측정 후 FlowDocument 를 temp RTB 에서 분리해 다음 사용(재측정·real RTB)에 대비.</item>
/// </list>
/// </para>
/// </summary>
internal static class SliceRefiner
{
    // PerPageEditorHost.ClipRenderingTolerance (2 DIP) + 서브픽셀 여유.
    // 이 값 이하의 초과분은 측정 오차로 간주해 보정하지 않는다.
    private const double OverflowTolerance = 3.0;

    /// <summary>
    /// 각 슬라이스를 오프스크린 RichTextBox 로 측정한 실제 콘텐츠 높이(DIP) 배열을 반환한다.
    /// 측정 후 FlowDocument 는 temp RTB 에서 분리되어 재사용 가능한 상태가 된다.
    /// </summary>
    public static double[] MeasureContentHeights(IReadOnlyList<PerPageDocumentSlice> slices)
    {
        var heights = new double[slices.Count];
        for (int i = 0; i < slices.Count; i++)
        {
            var slice = slices[i];
            var fd    = slice.FlowDocument;

            // FlowDocumentPaginationAdapter 와 동일한 off-screen 측정 패턴:
            // Measure(width, ∞) → Arrange(DesiredSize) → UpdateLayout().
            // DesiredSize.Height = 실제 콘텐츠 전체 높이.
            var tempRtb = new RichTextBox
            {
                Document          = fd,
                Padding           = new Thickness(0),
                BorderThickness   = new Thickness(0),
                VerticalAlignment = VerticalAlignment.Top,
            };

            tempRtb.Measure(new Size(slice.BodyWidthDip, double.PositiveInfinity));
            tempRtb.Arrange(new Rect(tempRtb.DesiredSize));
            tempRtb.UpdateLayout();

            heights[i] = tempRtb.DesiredSize.Height;

            // FlowDocument 분리 — 다음 사용(재측정 또는 SetupPages)을 위해 소유권 해제.
            tempRtb.Document = new FlowDocument();
        }
        return heights;
    }

    /// <summary>
    /// 오버플로 슬라이스의 마지막 WPF Block 을 다음 슬라이스 앞으로 이동한다.
    /// </summary>
    /// <param name="slices">검사·수정할 슬라이스 목록.</param>
    /// <param name="measuredHeights"><see cref="MeasureContentHeights"/> 결과.</param>
    /// <returns>하나 이상의 슬라이스가 수정되었으면 <c>true</c>.</returns>
    public static bool RefineOnce(
        IReadOnlyList<PerPageDocumentSlice> slices,
        double[]                            measuredHeights)
    {
        bool changed = false;

        for (int i = 0; i < slices.Count - 1; i++)
        {
            if (measuredHeights[i] <= slices[i].BodyHeightDip + OverflowTolerance)
                continue;

            var srcFd = slices[i].FlowDocument;

            // 1개 블록만 남았으면 이동 불가 — 분할 없이 이동하면 해당 슬라이스가 빈 페이지가 됨.
            if (srcFd.Blocks.Count <= 1) continue;

            var lastBlock = srcFd.Blocks.LastBlock;
            if (lastBlock is null) continue;

            // 테이블 블록의 Tag (Core.Table 참조) 를 백업해 이동 후 복원한다.
            // WPF 의 FlowDocument 이동 시 일부 메타데이터가 손실될 수 있으므로 명시적으로 보존.
            object? savedTag = null;
            if (lastBlock is Table wpfTable)
                savedTag = wpfTable.Tag;

            srcFd.Blocks.Remove(lastBlock);

            // 테이블의 Tag 복원.
            if (lastBlock is Table table && savedTag is not null)
                table.Tag = savedTag;

            var dstFd = slices[i + 1].FlowDocument;
            if (dstFd.Blocks.FirstBlock is { } firstDst)
                dstFd.Blocks.InsertBefore(firstDst, lastBlock);
            else
                dstFd.Blocks.Add(lastBlock);

            changed = true;
        }

        return changed;
    }
}

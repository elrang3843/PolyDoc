using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using PolyDonky.App.Services;
using PolyDonky.Core;
using WpfDocs = System.Windows.Documents;

namespace PolyDonky.App.Views;

/// <summary>
/// 글상자 내부 다단 레이아웃 헬퍼. 텍스트 박스의 <see cref="TextBoxObject.Content"/>
/// 블록들을 N 개 단(slot)으로 분배하고 단별 FlowDocument 를 만든다.
/// 본문 다단(<see cref="Pagination.FlowDocumentPaginationAdapter"/>) 과 같은 원리지만
/// "단일 페이지 × 다단" 으로 단순화 (페이지 개념 없음).
/// </summary>
public static class TextBoxColumnLayout
{
    /// <summary>
    /// 빈 단 placeholder 단락의 Id 마커. SyncMultiColContentToModel 에서 빈 placeholder 만
    /// 골라 제거하기 위함 — 사용자 입력으로 채워지면 Id 가 클리어된다.
    /// </summary>
    public const string PlaceholderId = "§tbph";

    /// <summary>분배 결과 한 단의 정보.</summary>
    public sealed class ColumnSlice
    {
        public required int                  ColumnIndex   { get; init; }
        public required double               WidthDip      { get; init; }
        public required double               HeightDip     { get; init; }
        public required double               XOffsetDip    { get; init; }
        public required IReadOnlyList<Block> Blocks        { get; init; }
        public required WpfDocs.FlowDocument FlowDocument  { get; init; }
    }

    /// <summary>
    /// <paramref name="content"/> 를 <paramref name="columnCount"/> 개 단으로 분배한다.
    /// 본문 다단의 <c>MapBodyBlocksToPages</c> + <c>SplitCoreParagraph</c> 와 동일한 원리:
    /// 단 폭을 측정 폭으로 잡은 오프스크린 RTB 에서 블록의 Y 좌표를 잰 뒤
    /// <c>slotIdx = floor(topY / colHeight)</c> 로 단 인덱스를 정한다.
    /// 경계를 가로지르는 단락은 줄 단위로 분할(§f0/§f1 접미사) — 본문 다단과 동일.
    /// </summary>
    public static IReadOnlyList<ColumnSlice> Distribute(
        IList<Block> content,
        int          columnCount,
        double       innerWidthDip,
        double       innerHeightDip,
        double       columnGapDip,
        IList<double>? columnWidthsDip = null)
    {
        ArgumentNullException.ThrowIfNull(content);
        if (columnCount < 1) columnCount = 1;
        if (innerWidthDip  < 1) innerWidthDip  = 1;
        if (innerHeightDip < 1) innerHeightDip = 1;

        // 단 너비 결정 (균등 또는 사용자 지정)
        double[] colWidths;
        if (columnCount > 1
            && columnWidthsDip is { Count: > 0 } wList
            && wList.Count == columnCount)
        {
            colWidths = wList.Select(w => Math.Max(10.0, w)).ToArray();
        }
        else if (columnCount > 1)
        {
            double equalW = Math.Max(10.0,
                (innerWidthDip - columnGapDip * (columnCount - 1)) / columnCount);
            colWidths = Enumerable.Repeat(equalW, columnCount).ToArray();
        }
        else
        {
            colWidths = new[] { innerWidthDip };
        }

        double[] colXOffsets = new double[columnCount];
        double offset = 0;
        for (int i = 0; i < columnCount; i++)
        {
            colXOffsets[i] = offset;
            offset += colWidths[i] + (columnCount > 1 ? columnGapDip : 0);
        }

        // 단일 단이거나 콘텐츠가 없으면 단순 처리
        if (columnCount == 1)
        {
            var fdSingle = FlowDocumentBuilder.BuildFromBlocks(content, page: null);
            fdSingle.PageWidth   = colWidths[0];
            fdSingle.PagePadding = new Thickness(0);
            return new[]
            {
                new ColumnSlice
                {
                    ColumnIndex  = 0,
                    WidthDip     = colWidths[0],
                    HeightDip    = innerHeightDip,
                    XOffsetDip   = 0,
                    Blocks       = content.ToArray(),
                    FlowDocument = fdSingle,
                }
            };
        }

        // 오프스크린 RTB 측정 — 단 폭 = 첫 단 너비 (서로 다른 너비 시 근사치).
        // 본문 다단의 측정 방식과 동일.
        var measureFd = FlowDocumentBuilder.BuildFromBlocks(content, page: null);
        measureFd.PageWidth   = colWidths[0];
        measureFd.PagePadding = new Thickness(0);

        var measureRtb = new RichTextBox
        {
            Document          = measureFd,
            Padding           = new Thickness(0),
            BorderThickness   = new Thickness(0),
            VerticalAlignment = VerticalAlignment.Top,
        };
        measureRtb.Measure(new Size(colWidths[0], double.PositiveInfinity));
        measureRtb.Arrange(new Rect(measureRtb.DesiredSize));
        measureRtb.UpdateLayout();

        // 본문과 동일한 mm→DIP 반올림 오차 흡수용 허용치
        const double BoundaryTol = 2.0;

        var slotBlocks = new List<List<Block>>();
        for (int i = 0; i < columnCount; i++) slotBlocks.Add(new List<Block>());
        var slotFill = new Dictionary<int, double>();

        foreach (var wpfBlock in FlattenBlocks(measureFd.Blocks))
        {
            if (wpfBlock.Tag is not Block coreBlock) continue;

            double topY    = TryGetTopY(wpfBlock);
            double bottomY = TryGetBottomY(wpfBlock);
            if (double.IsNaN(topY))
            {
                slotBlocks[0].Add(coreBlock);
                continue;
            }

            double blockH  = (!double.IsNaN(bottomY) && bottomY > topY) ? (bottomY - topY) : 0.0;
            int    slotTop = Math.Max(0, (int)(topY / innerHeightDip));

            // 누적 분할 — 한 단을 초과하는 단락도 단 경계마다 잘라서 §f0/§f1/§f2/...
            // 평탄한 fragment 들로 만든다. 본문 다단의 단일-분할 로직 + 다단계 반복 확장.
            // 입력 폭주(BG 우선순위 starvation)로 reflow 가 한 번에 몰릴 때 단락 높이가
            // innerHeightDip 를 넘어 분할이 막히는 문제 해결.
            if (coreBlock is Paragraph corePara
                && corePara.Style.ListMarker == null
                && blockH > 0
                && wpfBlock is WpfDocs.Paragraph wpfPara
                && slotTop < columnCount - 1)
            {
                // groupId — 모든 fragment 가 공유. 원본 Id 가 §f 를 포함하면 그 앞부분을 재사용,
                // 아니면 원본 Id 또는 새 GUID. 평탄 인덱스로 §f0/§f1/§f2 부여.
                string groupId;
                if (corePara.Id is { } origId)
                {
                    int sepIdx = origId.LastIndexOf("§f", StringComparison.Ordinal);
                    groupId = sepIdx >= 0 ? origId[..sepIdx] : origId;
                }
                else
                {
                    groupId = "§g" + Guid.NewGuid().ToString("N")[..8];
                }

                var    remainingPara = corePara;
                int    remainingChars = corePara.Runs.Sum(r => r.Text.Length);
                int    absOffsetSoFar = 0;
                double remainingTopY = topY;
                int    curSlot       = slotTop;
                int    fragIdx       = 0;
                bool   didSplit      = false;

                while (curSlot < columnCount - 1 && remainingChars > 0)
                {
                    double slotBoundaryY = (curSlot + 1) * innerHeightDip;
                    double remainingH    = (!double.IsNaN(bottomY) && bottomY > remainingTopY)
                                            ? bottomY - remainingTopY : 0.0;
                    double slotRemain    = innerHeightDip - slotFill.GetValueOrDefault(curSlot, 0.0);

                    bool crosses   = !double.IsNaN(bottomY) && bottomY > slotBoundaryY + BoundaryTol;
                    bool overflows = remainingH > slotRemain + BoundaryTol;
                    if (!crosses && !overflows) break;

                    double splitY = crosses
                        ? slotBoundaryY
                        : remainingTopY + Math.Max(0.0, slotRemain);

                    int splitAbs = Pagination.FlowDocumentPaginationAdapter
                        .FindSplitCharOffsetPublic(wpfPara, splitY);
                    int splitRel = splitAbs - absOffsetSoFar;
                    if (splitRel <= 0 || splitRel >= remainingChars) break;

                    var (first, second) = Pagination.FlowDocumentPaginationAdapter
                        .SplitCoreParagraphPublic(remainingPara, splitRel);

                    // 평탄 인덱스로 재명명 — SplitCoreParagraph 가 만든 중첩 §f0§f1 무시.
                    first.Id = groupId + "§f" + fragIdx;
                    fragIdx++;

                    double frag1H = Math.Max(0.0, splitY - remainingTopY);
                    slotFill[curSlot] = Math.Min(innerHeightDip,
                        slotFill.GetValueOrDefault(curSlot, 0.0) + frag1H);
                    slotBlocks[curSlot].Add(first);

                    remainingPara  = second;
                    remainingChars -= splitRel;
                    absOffsetSoFar = splitAbs;
                    remainingTopY  = splitY;
                    curSlot++;
                    didSplit = true;
                }

                if (didSplit)
                {
                    // 마지막 fragment 를 curSlot 에 — curSlot 이 마지막 단이면 거기 클립.
                    if (remainingChars > 0)
                    {
                        if (curSlot >= columnCount) curSlot = columnCount - 1;
                        remainingPara.Id = groupId + "§f" + fragIdx;
                        double frag2H = (!double.IsNaN(bottomY) && bottomY > remainingTopY)
                                            ? bottomY - remainingTopY : 0.0;
                        slotFill[curSlot] = Math.Min(innerHeightDip,
                            slotFill.GetValueOrDefault(curSlot, 0.0) + Math.Min(frag2H, innerHeightDip));
                        slotBlocks[curSlot].Add(remainingPara);
                    }
                    continue;
                }
            }

            // 단일 블록 배정
            if (!double.IsNaN(bottomY)
                && bottomY > (slotTop + 1) * innerHeightDip + BoundaryTol
                && blockH < innerHeightDip)
            {
                slotTop += 1;
            }
            if (blockH > 0 && blockH < innerHeightDip)
            {
                while (slotTop < columnCount - 1
                    && slotFill.GetValueOrDefault(slotTop, 0.0) + blockH > innerHeightDip + BoundaryTol)
                    slotTop += 1;
            }
            if (slotTop >= columnCount) slotTop = columnCount - 1;

            if (blockH > 0)
                slotFill[slotTop] = Math.Min(innerHeightDip,
                    slotFill.GetValueOrDefault(slotTop, 0.0) + Math.Min(blockH, innerHeightDip));
            slotBlocks[slotTop].Add(coreBlock);
        }

        // 측정용 RTB 분리
        measureRtb.Document = new WpfDocs.FlowDocument();

        // 단별 FlowDocument 생성
        var slices = new ColumnSlice[columnCount];
        for (int col = 0; col < columnCount; col++)
        {
            var blocks = slotBlocks[col];
            // 빈 단도 편집 가능하도록 placeholder 단락 1개 추가. PlaceholderId 마커로
            // SyncMultiColContentToModel 이 사용자 미입력 placeholder 를 골라 제거할 수 있게 함.
            if (blocks.Count == 0) blocks.Add(new Paragraph { Id = PlaceholderId });
            var fd = FlowDocumentBuilder.BuildFromBlocks(blocks, page: null);
            fd.PageWidth   = colWidths[col];
            fd.PagePadding = new Thickness(0);
            slices[col] = new ColumnSlice
            {
                ColumnIndex  = col,
                WidthDip     = colWidths[col],
                HeightDip    = innerHeightDip,
                XOffsetDip   = colXOffsets[col],
                Blocks       = blocks.ToArray(),
                FlowDocument = fd,
            };
        }
        return slices;
    }

    private static IEnumerable<WpfDocs.Block> FlattenBlocks(WpfDocs.BlockCollection blocks)
    {
        foreach (var b in blocks)
        {
            yield return b;
            if (b is WpfDocs.List list)
                foreach (var li in list.ListItems)
                    foreach (var nested in FlattenBlocks(li.Blocks))
                        yield return nested;
        }
    }

    private static double TryGetTopY(WpfDocs.Block block)
    {
        try
        {
            var r = block.ContentStart.GetCharacterRect(WpfDocs.LogicalDirection.Forward);
            if (r == Rect.Empty || double.IsInfinity(r.Y) || double.IsNaN(r.Y)) return double.NaN;
            return r.Y;
        }
        catch { return double.NaN; }
    }

    private static double TryGetBottomY(WpfDocs.Block block)
    {
        try
        {
            var r = block.ContentEnd.GetCharacterRect(WpfDocs.LogicalDirection.Backward);
            if (r == Rect.Empty || double.IsNaN(r.Bottom) || double.IsInfinity(r.Bottom))
                return double.NaN;
            return r.Bottom;
        }
        catch { return double.NaN; }
    }
}

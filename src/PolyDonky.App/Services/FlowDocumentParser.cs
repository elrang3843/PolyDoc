using System;
using System.Linq;
using System.Text;
using System.Windows;
using PolyDonky.Core;
using Wpf = System.Windows.Documents;
using WpfMedia = System.Windows.Media;

namespace PolyDonky.App.Services;

/// <summary>
/// RichTextBox 의 FlowDocument 를 PolyDonkyument 로 역변환한다.
/// FlowDocumentBuilder 가 보관해 둔 Tag(원본 PolyDonky 노드)와 머지해
/// 한글 조판 / Provenance 등 비-FlowDocument 속성을 비파괴 보존한다.
/// </summary>
public static class FlowDocumentParser
{
    public static PolyDonkyument Parse(Wpf.FlowDocument fd, PolyDonkyument? originalForMerge = null)
    {
        ArgumentNullException.ThrowIfNull(fd);

        var doc = new PolyDonkyument();
        if (originalForMerge is not null)
        {
            // 메타데이터·스타일·provenance·워터마크·개요 스타일은 본문 흐름에 들어가지 않는
            // 문서 수준 상태이므로 원본을 그대로 인계해야 save/load 시 보존된다.
            doc.Metadata = originalForMerge.Metadata;
            doc.Styles = originalForMerge.Styles;
            doc.Provenance = originalForMerge.Provenance;
            doc.Watermark = originalForMerge.Watermark;
            doc.OutlineStyles = originalForMerge.OutlineStyles;
            doc.Footnotes = originalForMerge.Footnotes;
            doc.Endnotes  = originalForMerge.Endnotes;
        }

        var section = new Section();
        if (originalForMerge?.Sections.FirstOrDefault() is { } origSection)
        {
            section.Page = origSection.Page;
        }
        doc.Sections.Add(section);

        ParseBlocks(section, fd.Blocks);

        // 글상자(TextBoxObject) 는 FlowDocument 안에 anchor paragraph 가 없는
        // 자유 부유 블록이라 ParseBlocks 로 복원되지 않는다. 원본 모델에서 직접 인계.
        if (originalForMerge?.Sections.FirstOrDefault() is { } origForTb)
        {
            foreach (var b in origForTb.Blocks)
            {
                if (b is TextBoxObject) section.Blocks.Add(b);
            }
        }
        return doc;
    }

    private static void ParseBlocks(Section section, IEnumerable<Wpf.Block> blocks)
    {
        ParseInto(section.Blocks, blocks);
    }

    /// <summary>
    /// 단일 WPF Block 을 Core.Block 으로 변환해 반환한다 (선택 영역 직렬화용 진입점).
    /// Tag 가 살아있으면 그대로 사용하고, 붙여넣기 등으로 Tag=null 인 경우 시각 트리에서 재구성.
    /// 매칭되는 변환이 없으면 null.
    /// </summary>
    public static Block? ParseSingleBlock(Wpf.Block wpfBlock)
    {
        ArgumentNullException.ThrowIfNull(wpfBlock);
        var bucket = new List<Block>();
        ParseInto(bucket, new[] { wpfBlock });
        return bucket.FirstOrDefault();
    }

    /// <summary>
    /// 단락을 [clipStart, clipEnd] 선택 범위로 잘라 파싱한다.
    /// Run 은 범위 내 글자만 추출하고, InlineUIContainer / Floater 등 내장 객체는
    /// 범위 내에 있으면 회피 없이 그대로 포함한다.
    /// </summary>
    public static Paragraph ParseParagraphClipped(
        Wpf.Paragraph wpfPara,
        Wpf.TextPointer clipStart,
        Wpf.TextPointer clipEnd)
    {
        var p = wpfPara.Tag is Paragraph orig ? CloneShallow(orig) : new Paragraph();
        p.Runs.Clear();
        p.Style.ListMarker = null;

        p.Style.Alignment = wpfPara.TextAlignment switch
        {
            TextAlignment.Center  => Alignment.Center,
            TextAlignment.Right   => Alignment.Right,
            TextAlignment.Justify => Alignment.Justify,
            _                     => Alignment.Left,
        };

        if (wpfPara.Tag is Paragraph tagged && tagged.Style.Outline > OutlineLevel.Body)
            p.Style.Outline = tagged.Style.Outline;
        else
            p.Style.Outline = InferHeadingFromFontSize(wpfPara.FontSize);

        if (wpfPara.Tag is Paragraph baseP)
        {
            p.Style.IndentFirstLineMm  = baseP.Style.IndentFirstLineMm;
            p.Style.IndentLeftMm       = baseP.Style.IndentLeftMm;
            p.Style.IndentRightMm      = baseP.Style.IndentRightMm;
            p.Style.SpaceBeforePt      = baseP.Style.SpaceBeforePt;
            p.Style.SpaceAfterPt       = baseP.Style.SpaceAfterPt;
            p.Style.LineHeightFactor   = baseP.Style.LineHeightFactor;
        }

        foreach (var inline in wpfPara.Inlines)
            ParseInlineClipped(p, inline, new RunStyle(), clipStart, clipEnd);

        if (p.Runs.Count == 0) p.AddText(string.Empty);
        return p;
    }

    private static void ParseInlineClipped(
        Paragraph p, Wpf.Inline inline, RunStyle baseStyle,
        Wpf.TextPointer clipStart, Wpf.TextPointer clipEnd)
    {
        // 이 Inline 이 선택 범위 밖이면 건너뜀
        if (inline.ContentEnd.CompareTo(clipStart) <= 0) return;
        if (inline.ContentStart.CompareTo(clipEnd) >= 0) return;

        switch (inline)
        {
            case Wpf.Run r:
            {
                var effStart = r.ContentStart.CompareTo(clipStart) < 0 ? clipStart : r.ContentStart;
                var effEnd   = r.ContentEnd.CompareTo(clipEnd)     > 0 ? clipEnd   : r.ContentEnd;
                var text = new Wpf.TextRange(effStart, effEnd).Text;
                if (string.IsNullOrEmpty(text)) return;
                var seed = r.Tag is Run original ? Clone(original.Style) : Clone(baseStyle);
                ExtractRunStyle(r, seed);
                p.AddText(text, seed);
                break;
            }
            case Wpf.Span span:
            {
                // per-char Span (자간·글자폭) — 선택 범위 교차 텍스트만 잘라냄
                if (span.Tag is Run spanRun && TryMergePerCharSpan(span, spanRun, out _))
                {
                    var effStart = span.ContentStart.CompareTo(clipStart) < 0 ? clipStart : span.ContentStart;
                    var effEnd   = span.ContentEnd.CompareTo(clipEnd)     > 0 ? clipEnd   : span.ContentEnd;
                    var clipped  = new Wpf.TextRange(effStart, effEnd).Text;
                    if (!string.IsNullOrEmpty(clipped)) p.AddText(clipped, Clone(spanRun.Style));
                    break;
                }
                var spanStyle = Clone(baseStyle);
                MergeInlineProperties(spanStyle, span);
                foreach (var child in span.Inlines)
                    ParseInlineClipped(p, child, spanStyle, clipStart, clipEnd);
                break;
            }
            case Wpf.LineBreak:
                p.AddText("\n", Clone(baseStyle));
                break;
            default:
                // InlineUIContainer(이미지·수식·이모지), Floater(WrapLeft/WrapRight 이미지·도형) 등
                // 선택 범위 내에 있으면 회피 없이 그대로 포함
                ParseInline(p, inline, baseStyle);
                break;
        }
    }

    private static void ParseInto(IList<Block> target, IEnumerable<Wpf.Block> blocks)
    {
        foreach (var block in blocks)
        {
            // 합성 페이지-갭 패딩 단락은 시각 장치이므로 모델로 절대 흘러가면 안 된다.
            if (PageBreakPadder.IsPagePadding(block)) continue;

            switch (block)
            {
                // Tag 가 OpaqueBlock 인 paragraph 는 보존 섬 — 원본 그대로 회수.
                case Wpf.Paragraph { Tag: OpaqueBlock opaque }:
                    target.Add(opaque);
                    break;

                // 래핑 모드(WrapLeft/WrapRight) 그림은 Floater 가 든 Paragraph 로 빌드됨 — Tag 로 회수.
                // ShapeObject 래핑 모드도 동일. 반드시 일반 'case Wpf.Paragraph' 보다 먼저 위치해야 한다.
                case Wpf.Paragraph wrappedImagePara when wrappedImagePara.Tag is ImageBlock wrappedImage:
                    target.Add(wrappedImage);
                    break;

                case Wpf.Paragraph wrappedShapePara when wrappedShapePara.Tag is ShapeObject wrappedShape:
                    target.Add(wrappedShape);
                    break;

                // 오버레이(InFrontOfText/BehindText/Fixed) 모드 표 앵커 단락 — Tag 로 원본 Table 을 회수.
                case Wpf.Paragraph wrappedTablePara when wrappedTablePara.Tag is Table wrappedTable:
                    target.Add(wrappedTable);
                    break;

                // Fallback: 붙여넣기로 Tag 가 사라진 AsText/WrapLeft/WrapRight 이미지 단락.
                // WPF XamlPackage 클립보드 포맷은 BitmapSource 를 보존하므로 시각 트리에서 ImageBlock 을 재구성한다.
                // 반드시 일반 'case Wpf.Paragraph' 보다 먼저 위치해야 한다.
                case Wpf.Paragraph pastedImgPara
                    when pastedImgPara.Tag is null && TryRecoverImageFromPara(pastedImgPara, out var pastedWrapImg):
                    target.Add(pastedWrapImg!);
                    break;

                case Wpf.Paragraph wpfPara:
                    target.Add(ParseParagraph(wpfPara, listMarker: null));
                    break;

                case Wpf.List list:
                {
                    var kind = IsBulletMarker(list.MarkerStyle) ? ListKind.Bullet : ListKind.OrderedDecimal;
                    var counter = 0;
                    foreach (var item in list.ListItems)
                    {
                        counter++;
                        var marker = new ListMarker
                        {
                            Kind = kind,
                            OrderedNumber = kind == ListKind.OrderedDecimal
                                ? Math.Max(list.StartIndex, 1) + counter - 1
                                : null,
                        };
                        foreach (var inner in item.Blocks)
                        {
                            if (inner is Wpf.Paragraph pp)
                            {
                                target.Add(ParseParagraph(pp, marker));
                            }
                        }
                    }
                    break;
                }

                case Wpf.Table wpfTable:
                    target.Add(ParseTable(wpfTable));
                    break;

                case Wpf.BlockUIContainer container when container.Tag is ImageBlock image:
                    target.Add(image);
                    break;

                case Wpf.BlockUIContainer container when container.Tag is ShapeObject shapeInline:
                    target.Add(shapeInline);
                    break;

                // Fallback: 붙여넣기로 Tag 가 사라진 Inline 모드 이미지.
                // BitmapSource 는 XamlPackage 포맷에 보존되므로 시각 트리에서 ImageBlock 을 재구성한다.
                case Wpf.BlockUIContainer pastedBuc
                    when pastedBuc.Tag is null && TryRecoverImageFromBUC(pastedBuc, out var pastedInlineImg):
                    target.Add(pastedInlineImg!);
                    break;

                case Wpf.Section nested:
                    ParseInto(target, nested.Blocks);
                    break;
            }
        }
    }

    private static Table ParseTable(Wpf.Table wpfTable)
    {
        // 표는 사용자가 셀 안 텍스트만 편집할 수 있다고 가정. 구조(행/열/병합)와 컬럼 너비는 Tag 의 원본을
        // 베이스로 두고, 셀 본문만 파싱 결과로 갱신해 비파괴 라운드트립을 보장한다.
        var seed = wpfTable.Tag is Table original
            ? new Table
            {
                Id      = original.Id,
                Status  = original.Status,
                WrapMode  = original.WrapMode,
                HAlign    = original.HAlign,
                OverlayXMm = original.OverlayXMm,
                OverlayYMm = original.OverlayYMm,
                BackgroundColor              = original.BackgroundColor,
                DefaultCellPaddingTopMm      = original.DefaultCellPaddingTopMm,
                DefaultCellPaddingBottomMm   = original.DefaultCellPaddingBottomMm,
                DefaultCellPaddingLeftMm     = original.DefaultCellPaddingLeftMm,
                DefaultCellPaddingRightMm    = original.DefaultCellPaddingRightMm,
                OuterMarginTopMm             = original.OuterMarginTopMm,
                OuterMarginBottomMm          = original.OuterMarginBottomMm,
                OuterMarginLeftMm            = original.OuterMarginLeftMm,
                OuterMarginRightMm           = original.OuterMarginRightMm,
                BorderThicknessPt            = original.BorderThicknessPt,
                BorderColor                  = original.BorderColor,
                Columns = new List<TableColumn>(original.Columns.Select(c => new TableColumn { WidthMm = c.WidthMm })),
            }
            : new Table();

        var rowGroup = wpfTable.RowGroups.FirstOrDefault();
        if (rowGroup is null)
        {
            return seed;
        }

        for (int rowIndex = 0; rowIndex < rowGroup.Rows.Count; rowIndex++)
        {
            var wpfRow = rowGroup.Rows[rowIndex];
            var origRow = (wpfTable.Tag is Table o && rowIndex < o.Rows.Count) ? o.Rows[rowIndex] : null;
            var row = new TableRow { HeightMm = origRow?.HeightMm ?? 0, IsHeader = origRow?.IsHeader ?? false };

            for (int cellIndex = 0; cellIndex < wpfRow.Cells.Count; cellIndex++)
            {
                var wpfCell = wpfRow.Cells[cellIndex];
                var origCell = (origRow is not null && cellIndex < origRow.Cells.Count) ? origRow.Cells[cellIndex] : null;
                var cell = new TableCell
                {
                    ColumnSpan        = wpfCell.ColumnSpan > 0 ? wpfCell.ColumnSpan : (origCell?.ColumnSpan ?? 1),
                    RowSpan           = wpfCell.RowSpan    > 0 ? wpfCell.RowSpan    : (origCell?.RowSpan    ?? 1),
                    WidthMm           = origCell?.WidthMm           ?? 0,
                    TextAlign         = origCell?.TextAlign         ?? CellTextAlign.Left,
                    PaddingTopMm      = origCell?.PaddingTopMm      ?? 0,
                    PaddingBottomMm   = origCell?.PaddingBottomMm   ?? 0,
                    PaddingLeftMm     = origCell?.PaddingLeftMm     ?? 0,
                    PaddingRightMm    = origCell?.PaddingRightMm    ?? 0,
                    BorderThicknessPt = origCell?.BorderThicknessPt ?? 0,
                    BorderColor       = origCell?.BorderColor,
                    BackgroundColor   = origCell?.BackgroundColor,
                };
                ParseInto(cell.Blocks, wpfCell.Blocks);
                if (cell.Blocks.Count == 0)
                {
                    cell.Blocks.Add(Paragraph.Of(string.Empty));
                }
                row.Cells.Add(cell);
            }

            seed.Rows.Add(row);
        }

        return seed;
    }

    private static bool IsBulletMarker(TextMarkerStyle style)
        => style is TextMarkerStyle.Disc
            or TextMarkerStyle.Circle
            or TextMarkerStyle.Square
            or TextMarkerStyle.Box
            or TextMarkerStyle.None;

    private static Paragraph ParseParagraph(Wpf.Paragraph wpfPara, ListMarker? listMarker)
    {
        // Tag 에 원본 Paragraph 가 있으면 머지 베이스로 시작.
        var p = wpfPara.Tag is Paragraph orig ? CloneShallow(orig) : new Paragraph();
        p.Runs.Clear();
        p.Style.ListMarker = listMarker;

        p.Style.Alignment = wpfPara.TextAlignment switch
        {
            TextAlignment.Center => Alignment.Center,
            TextAlignment.Right => Alignment.Right,
            TextAlignment.Justify => Alignment.Justify,
            _ => Alignment.Left,
        };

        // Heading 추정: 단락 Tag 에 원본 Outline 이 있으면 우선, 없으면 FontSize 로 단순 추정.
        if (wpfPara.Tag is Paragraph tagged && tagged.Style.Outline > OutlineLevel.Body)
        {
            p.Style.Outline = tagged.Style.Outline;
        }
        else
        {
            p.Style.Outline = InferHeadingFromFontSize(wpfPara.FontSize);
        }

        p.Style.ForcePageBreakBefore = wpfPara.BreakPageBefore;

        // 들여쓰기·간격은 사용자가 직접 변경할 가능성 높지만, 본 사이클에서는 단순 보존만.
        if (wpfPara.Tag is Paragraph baseP)
        {
            p.Style.IndentFirstLineMm = baseP.Style.IndentFirstLineMm;
            p.Style.IndentLeftMm = baseP.Style.IndentLeftMm;
            p.Style.IndentRightMm = baseP.Style.IndentRightMm;
            p.Style.SpaceBeforePt = baseP.Style.SpaceBeforePt;
            p.Style.SpaceAfterPt = baseP.Style.SpaceAfterPt;
            p.Style.LineHeightFactor = baseP.Style.LineHeightFactor;
        }

        foreach (var inline in wpfPara.Inlines)
        {
            ParseInline(p, inline, baseStyle: new RunStyle());
        }

        if (p.Runs.Count == 0)
        {
            p.AddText(string.Empty);
        }
        return p;
    }

    private static OutlineLevel InferHeadingFromFontSize(double fontSizeDip)
    {
        // FlowDocumentBuilder 가 사용한 헤더별 크기를 역추정.
        var pt = FlowDocumentBuilder.DipToPt(fontSizeDip);
        return pt switch
        {
            >= 23 => OutlineLevel.H1,
            >= 19 => OutlineLevel.H2,
            >= 16 => OutlineLevel.H3,
            >= 14 => OutlineLevel.H4,
            >= 12.5 => OutlineLevel.H5,
            >= 11.5 => OutlineLevel.H6,
            _ => OutlineLevel.Body,
        };
    }

    private static void ParseInline(Paragraph p, Wpf.Inline inline, RunStyle baseStyle)
    {
        switch (inline)
        {
            case Wpf.Run r:
            {
                // 각주/미주 참조 런 — Tag 의 원본 Run 을 그대로 회수 (FootnoteId/EndnoteId 보존).
                if (r.Tag is Run origRef && (origRef.FootnoteId is not null || origRef.EndnoteId is not null))
                {
                    p.Runs.Add(origRef.Clone());
                    break;
                }
                // Tag 에 원본 Run 이 있으면 그 RunStyle 을 base 로 두고 변경된 속성만 덮어쓴다.
                var seed = r.Tag is Run original ? Clone(original.Style) : Clone(baseStyle);
                ExtractRunStyle(r, seed);
                p.AddText(r.Text, seed);
                break;
            }
            case Wpf.InlineUIContainer iuc:
            {
                // FlowDocumentBuilder 가 만든 컨테이너.
                // Tag 에 원본 PolyDonky Run 이 있으면 직접 회수. 없으면 시각 트리에서 추출.
                if (iuc.Tag is Run origRun)
                {
                    // Run.Clone() 으로 모든 필드 복사 — 이전 자체 구현은 Url 누락.
                    p.Runs.Add(origRun.Clone());
                }
                else if (iuc.Child is System.Windows.Controls.StackPanel panel)
                {
                    // 자간 StackPanel: 첫 TextBlock 에서 스타일·자간 읽기
                    var style = new RunStyle { WidthPercent = 100 };
                    var sb = new System.Text.StringBuilder();
                    System.Windows.Controls.TextBlock? firstTb = null;
                    foreach (var child in panel.Children)
                    {
                        if (child is System.Windows.Controls.TextBlock ctb)
                        {
                            sb.Append(ctb.Text);
                            firstTb ??= ctb;
                        }
                    }
                    if (firstTb != null)
                    {
                        style.LetterSpacingPx = firstTb.Margin.Right;
                        if (firstTb.LayoutTransform is WpfMedia.ScaleTransform st)
                            style.WidthPercent = st.ScaleX * 100.0;
                        ExtractStyleFromTextBlock(firstTb, style);
                    }
                    p.AddText(sb.ToString(), style);
                }
                else if (iuc.Child is System.Windows.Controls.TextBlock tb)
                {
                    var style = new RunStyle { WidthPercent = 100 };
                    if (tb.LayoutTransform is WpfMedia.ScaleTransform st)
                        style.WidthPercent = st.ScaleX * 100.0;
                    ExtractStyleFromTextBlock(tb, style);
                    p.AddText(tb.Text, style);
                }
                break;
            }
            case Wpf.LineBreak:
                p.AddText("\n", Clone(baseStyle));
                break;
            case Wpf.Span span:
            {
                // 글자폭·자간 시각화 Span — Tag 가 PolyDonky.Run 이고 자식이 모두 같은 Tag 의 per-char IUC 면
                // 한 덩어리로 머지해 원본 Run 의 비-FlowDocument 속성(WidthPercent / LetterSpacingPx) 보존.
                if (span.Tag is Run spanRun && TryMergePerCharSpan(span, spanRun, out var mergedText))
                {
                    p.AddText(mergedText, Clone(spanRun.Style));
                    break;
                }
                var spanStyle = Clone(baseStyle);
                MergeInlineProperties(spanStyle, span);
                foreach (var child in span.Inlines)
                {
                    ParseInline(p, child, spanStyle);
                }
                break;
            }
        }
    }

    /// <summary>
    /// per-char IUC Span 의 자식이 모두 같은 Run Tag 를 공유하는 IUC 면 텍스트를 모아 반환.
    /// 사용자 편집(중간에 Run 삽입, 일부 IUC 제거 등)이 있으면 false — fallback 으로 자식별 파싱.
    /// </summary>
    private static bool TryMergePerCharSpan(Wpf.Span span, Run spanRun, out string text)
    {
        var sb = new StringBuilder();
        foreach (var child in span.Inlines)
        {
            if (child is Wpf.InlineUIContainer iuc
                && ReferenceEquals(iuc.Tag, spanRun)
                && iuc.Child is System.Windows.Controls.TextBlock ctb)
            {
                sb.Append(ctb.Text);
            }
            else
            {
                text = string.Empty;
                return false;
            }
        }
        text = sb.ToString();
        return text.Length > 0;
    }

    private static void ExtractRunStyle(Wpf.Run wpfRun, RunStyle s)
    {
        MergeInlineProperties(s, wpfRun);

        if (wpfRun.BaselineAlignment == BaselineAlignment.Superscript)
        {
            s.Superscript = true;
            s.Subscript = false;
        }
        else if (wpfRun.BaselineAlignment == BaselineAlignment.Subscript)
        {
            s.Subscript = true;
            s.Superscript = false;
        }

        // 자간·글자폭 — Tag 에 원본 Run 이 있으면 그 값을 보존 (MergeInlineProperties 가 덮어쓰지 않으므로 여기서 복원).
        if (wpfRun.Tag is Run origRun)
        {
            if (Math.Abs(origRun.Style.LetterSpacingPx) > 0.01)
                s.LetterSpacingPx = origRun.Style.LetterSpacingPx;
            if (Math.Abs(origRun.Style.WidthPercent - 100) > 0.5)
                s.WidthPercent = origRun.Style.WidthPercent;
        }
    }

    private static void ExtractStyleFromTextBlock(System.Windows.Controls.TextBlock tb, RunStyle s)
    {
        s.FontSizePt = FlowDocumentBuilder.DipToPt(tb.FontSize);
        s.Bold = tb.FontWeight.ToOpenTypeWeight() >= FontWeights.Bold.ToOpenTypeWeight();
        s.Italic = tb.FontStyle == FontStyles.Italic;
        if (tb.FontFamily != null) s.FontFamily = tb.FontFamily.Source;
        if (tb.Foreground is WpfMedia.SolidColorBrush fg)
            s.Foreground = new Color(fg.Color.R, fg.Color.G, fg.Color.B, fg.Color.A);
        if (tb.Background is WpfMedia.SolidColorBrush bg)
            s.Background = new Color(bg.Color.R, bg.Color.G, bg.Color.B, bg.Color.A);
        if (tb.TextDecorations is { Count: > 0 } decos)
        {
            foreach (var d in decos)
            {
                if (d.Location == TextDecorationLocation.Underline) s.Underline = true;
                else if (d.Location == TextDecorationLocation.Strikethrough) s.Strikethrough = true;
                else if (d.Location == TextDecorationLocation.OverLine) s.Overline = true;
            }
        }
    }

    private static void MergeInlineProperties(RunStyle s, Wpf.Inline inline)
    {
        var family = inline.FontFamily?.Source;
        if (!string.IsNullOrEmpty(family))
        {
            s.FontFamily = family;
        }

        // FontSize 는 항상 inheritance 결과로 값이 있다. dip → pt 변환.
        s.FontSizePt = FlowDocumentBuilder.DipToPt(inline.FontSize);

        s.Bold = inline.FontWeight.ToOpenTypeWeight() >= FontWeights.Bold.ToOpenTypeWeight();
        s.Italic = inline.FontStyle == FontStyles.Italic;

        s.Underline = false;
        s.Strikethrough = false;
        s.Overline = false;
        if (inline.TextDecorations is { Count: > 0 } decos)
        {
            foreach (var deco in decos)
            {
                if (deco.Location == TextDecorationLocation.Underline)
                {
                    s.Underline = true;
                }
                else if (deco.Location == TextDecorationLocation.Strikethrough)
                {
                    s.Strikethrough = true;
                }
                else if (deco.Location == TextDecorationLocation.OverLine)
                {
                    s.Overline = true;
                }
            }
        }

        if (inline.Foreground is WpfMedia.SolidColorBrush fg)
        {
            s.Foreground = new Color(fg.Color.R, fg.Color.G, fg.Color.B, fg.Color.A);
        }
        if (inline.Background is WpfMedia.SolidColorBrush bg)
        {
            s.Background = new Color(bg.Color.R, bg.Color.G, bg.Color.B, bg.Color.A);
        }
    }

    // ── 붙여넣기 후 Tag 가 없는 이미지 재구성 헬퍼 ─────────────────────────

    /// <summary>
    /// Tag=null 인 BlockUIContainer(Inline 모드 이미지)에서 ImageBlock 을 재구성한다.
    /// WPF XamlPackage 클립보드 포맷은 BitmapSource 를 온전히 보존하므로 크기·정렬·테두리를 복구할 수 있다.
    /// </summary>
    private static bool TryRecoverImageFromBUC(Wpf.BlockUIContainer buc, out ImageBlock? img)
    {
        img = null;
        var (wpfImg, border) = FindImageInVisual(buc.Child);
        if (wpfImg?.Source is not System.Windows.Media.Imaging.BitmapSource) return false;

        var ha = (border?.HorizontalAlignment ?? wpfImg.HorizontalAlignment) switch
        {
            System.Windows.HorizontalAlignment.Center => ImageHAlign.Center,
            System.Windows.HorizontalAlignment.Right  => ImageHAlign.Right,
            _                                          => ImageHAlign.Left,
        };
        img = BuildRecoveredImageBlock(wpfImg, border, buc.Margin, ImageWrapMode.Inline, ha);
        return true;
    }

    /// <summary>
    /// Tag=null 인 Paragraph 에서 AsText / WrapLeft / WrapRight 이미지 단락을 재구성한다.
    /// </summary>
    private static bool TryRecoverImageFromPara(Wpf.Paragraph para, out ImageBlock? img)
    {
        img = null;

        // AsText 모드: 단 하나의 InlineUIContainer(Image 또는 Border>Image)
        if (para.Inlines.Count == 1
            && para.Inlines.FirstOrDefault() is Wpf.InlineUIContainer iuc)
        {
            var (wpfImg, border) = FindImageInVisual(iuc.Child);
            if (wpfImg?.Source is System.Windows.Media.Imaging.BitmapSource)
            {
                var ha = (border?.HorizontalAlignment ?? wpfImg.HorizontalAlignment) switch
                {
                    System.Windows.HorizontalAlignment.Center => ImageHAlign.Center,
                    System.Windows.HorizontalAlignment.Right  => ImageHAlign.Right,
                    _                                          => ImageHAlign.Left,
                };
                img = BuildRecoveredImageBlock(wpfImg, border, para.Margin, ImageWrapMode.AsText, ha);
                return true;
            }
        }

        // WrapLeft / WrapRight 모드: Floater > BlockUIContainer > Image
        var floater = para.Inlines.OfType<Wpf.Floater>().FirstOrDefault();
        if (floater is null) return false;
        var innerBuc = floater.Blocks.OfType<Wpf.BlockUIContainer>().FirstOrDefault();
        if (innerBuc is null) return false;
        {
            var (wpfImg, border) = FindImageInVisual(innerBuc.Child);
            if (wpfImg?.Source is not System.Windows.Media.Imaging.BitmapSource) return false;
            var mode = floater.HorizontalAlignment == System.Windows.HorizontalAlignment.Right
                       ? ImageWrapMode.WrapRight
                       : ImageWrapMode.WrapLeft;
            var ha = mode == ImageWrapMode.WrapRight ? ImageHAlign.Right : ImageHAlign.Left;
            img = BuildRecoveredImageBlock(wpfImg, border, para.Margin, mode, ha);
            return true;
        }
    }

    private static (System.Windows.Controls.Image? img, System.Windows.Controls.Border? border)
        FindImageInVisual(System.Windows.UIElement? child) => child switch
    {
        System.Windows.Controls.Image i => (i, null),
        System.Windows.Controls.Border b when b.Child is System.Windows.Controls.Image bi => (bi, b),
        _ => (null, null),
    };

    private static ImageBlock BuildRecoveredImageBlock(
        System.Windows.Controls.Image wpfImg,
        System.Windows.Controls.Border? border,
        System.Windows.Thickness margin,
        ImageWrapMode mode,
        ImageHAlign hAlign)
    {
        var ib = new ImageBlock { WrapMode = mode, HAlign = hAlign };

        if (!double.IsNaN(wpfImg.Width)  && wpfImg.Width  > 0)
            ib.WidthMm  = FlowDocumentBuilder.DipToMm(wpfImg.Width);
        if (!double.IsNaN(wpfImg.Height) && wpfImg.Height > 0)
            ib.HeightMm = FlowDocumentBuilder.DipToMm(wpfImg.Height);

        if (margin.Top    > 0) ib.MarginTopMm    = FlowDocumentBuilder.DipToMm(margin.Top);
        if (margin.Bottom > 0) ib.MarginBottomMm = FlowDocumentBuilder.DipToMm(margin.Bottom);

        if (border is not null && border.BorderThickness.Left > 0)
        {
            ib.BorderThicknessPt = FlowDocumentBuilder.DipToPt(border.BorderThickness.Left);
            if (border.BorderBrush is WpfMedia.SolidColorBrush scb)
                ib.BorderColor = $"#{scb.Color.R:X2}{scb.Color.G:X2}{scb.Color.B:X2}";
        }

        if (wpfImg.Source is System.Windows.Media.Imaging.BitmapSource bmp)
        {
            try
            {
                using var ms = new System.IO.MemoryStream();
                var encoder = new System.Windows.Media.Imaging.PngBitmapEncoder();
                encoder.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(bmp));
                encoder.Save(ms);
                ib.Data      = ms.ToArray();
                ib.MediaType = "image/png";
            }
            catch { /* 인코딩 실패 시 Data 비워둠 — [이미지 누락] 플레이스홀더로 표시됨 */ }
        }

        return ib;
    }

    // Core 의 정식 RunStyle.Clone() 사용 — 자체 구현 시 새 필드(Url 등) 누락 위험.
    private static RunStyle Clone(RunStyle s) => s.Clone();
    private static Paragraph CloneShallow(Paragraph original)
    {
        var p = new Paragraph
        {
            Id = original.Id,
            Status = original.Status,
            StyleId = original.StyleId,
        };
        p.Style.Alignment = original.Style.Alignment;
        p.Style.LineHeightFactor = original.Style.LineHeightFactor;
        p.Style.SpaceBeforePt = original.Style.SpaceBeforePt;
        p.Style.SpaceAfterPt = original.Style.SpaceAfterPt;
        p.Style.IndentFirstLineMm = original.Style.IndentFirstLineMm;
        p.Style.IndentLeftMm = original.Style.IndentLeftMm;
        p.Style.IndentRightMm = original.Style.IndentRightMm;
        p.Style.Outline = original.Style.Outline;
        p.Style.ListMarker = original.Style.ListMarker is { } m
            ? new ListMarker { Kind = m.Kind, Level = m.Level, OrderedNumber = m.OrderedNumber }
            : null;
        return p;
    }
}

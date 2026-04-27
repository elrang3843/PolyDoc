using System;
using System.Linq;
using System.Text;
using System.Windows;
using PolyDoc.Core;
using Wpf = System.Windows.Documents;
using WpfMedia = System.Windows.Media;

namespace PolyDoc.App.Services;

/// <summary>
/// RichTextBox 의 FlowDocument 를 PolyDocument 로 역변환한다.
/// FlowDocumentBuilder 가 보관해 둔 Tag(원본 PolyDoc 노드)와 머지해
/// 한글 조판 / Provenance 등 비-FlowDocument 속성을 비파괴 보존한다.
/// </summary>
public static class FlowDocumentParser
{
    public static PolyDocument Parse(Wpf.FlowDocument fd, PolyDocument? originalForMerge = null)
    {
        ArgumentNullException.ThrowIfNull(fd);

        var doc = new PolyDocument();
        if (originalForMerge is not null)
        {
            // 메타데이터·스타일·provenance 는 원본을 인계 (편집 후에도 유지).
            doc.Metadata = originalForMerge.Metadata;
            doc.Styles = originalForMerge.Styles;
            doc.Provenance = originalForMerge.Provenance;
        }

        var section = new Section();
        if (originalForMerge?.Sections.FirstOrDefault() is { } origSection)
        {
            section.Page = origSection.Page;
        }
        doc.Sections.Add(section);

        ParseBlocks(section, fd.Blocks);
        return doc;
    }

    private static void ParseBlocks(Section section, IEnumerable<Wpf.Block> blocks)
    {
        ParseInto(section.Blocks, blocks);
    }

    private static void ParseInto(IList<Block> target, IEnumerable<Wpf.Block> blocks)
    {
        foreach (var block in blocks)
        {
            switch (block)
            {
                // Tag 가 OpaqueBlock 인 paragraph 는 보존 섬 — 원본 그대로 회수.
                case Wpf.Paragraph { Tag: OpaqueBlock opaque }:
                    target.Add(opaque);
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
                Id = original.Id,
                Status = original.Status,
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
            var row = new TableRow { HeightMm = origRow?.HeightMm ?? 0 };

            for (int cellIndex = 0; cellIndex < wpfRow.Cells.Count; cellIndex++)
            {
                var wpfCell = wpfRow.Cells[cellIndex];
                var origCell = (origRow is not null && cellIndex < origRow.Cells.Count) ? origRow.Cells[cellIndex] : null;
                var cell = new TableCell
                {
                    ColumnSpan = wpfCell.ColumnSpan > 0 ? wpfCell.ColumnSpan : (origCell?.ColumnSpan ?? 1),
                    RowSpan = wpfCell.RowSpan > 0 ? wpfCell.RowSpan : (origCell?.RowSpan ?? 1),
                    WidthMm = origCell?.WidthMm ?? 0,
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
                // Tag 에 원본 Run 이 있으면 그 RunStyle 을 base 로 두고 변경된 속성만 덮어쓴다.
                var seed = r.Tag is Run original ? Clone(original.Style) : Clone(baseStyle);
                ExtractRunStyle(r, seed);
                p.AddText(r.Text, seed);
                break;
            }
            case Wpf.InlineUIContainer iuc:
            {
                // FlowDocumentBuilder 가 만든 컨테이너.
                // Tag 에 원본 PolyDoc Run 이 있으면 직접 회수. 없으면 시각 트리에서 추출.
                if (iuc.Tag is Run origRun)
                {
                    // 수식 Run 은 LatexSource / IsDisplayEquation 보존
                    p.Runs.Add(new Run
                    {
                        Text              = origRun.Text,
                        Style             = Clone(origRun.Style),
                        LatexSource       = origRun.LatexSource,
                        IsDisplayEquation = origRun.IsDisplayEquation,
                    });
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
                // 글자폭·자간 시각화 Span — Tag 가 PolyDoc.Run 이고 자식이 모두 같은 Tag 의 per-char IUC 면
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

    private static RunStyle Clone(RunStyle s) => new()
    {
        FontFamily = s.FontFamily,
        FontSizePt = s.FontSizePt,
        Bold = s.Bold,
        Italic = s.Italic,
        Underline = s.Underline,
        Strikethrough = s.Strikethrough,
        Overline = s.Overline,
        Superscript = s.Superscript,
        Subscript = s.Subscript,
        Foreground = s.Foreground,
        Background = s.Background,
        WidthPercent = s.WidthPercent,
        LetterSpacingPx = s.LetterSpacingPx,
    };

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

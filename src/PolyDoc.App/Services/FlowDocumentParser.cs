using System;
using System.Linq;
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
        foreach (var block in blocks)
        {
            switch (block)
            {
                case Wpf.Paragraph wpfPara:
                    section.Blocks.Add(ParseParagraph(wpfPara, listMarker: null));
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
                                section.Blocks.Add(ParseParagraph(pp, marker));
                            }
                        }
                    }
                    break;
                }

                case Wpf.Section nested:
                    ParseBlocks(section, nested.Blocks);
                    break;
            }
        }
    }

    private static bool IsBulletMarker(Wpf.TextMarkerStyle style)
        => style is Wpf.TextMarkerStyle.Disc
            or Wpf.TextMarkerStyle.Circle
            or Wpf.TextMarkerStyle.Square
            or Wpf.TextMarkerStyle.Box
            or Wpf.TextMarkerStyle.None;

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
            case Wpf.LineBreak:
                p.AddText("\n", Clone(baseStyle));
                break;
            case Wpf.Span span:
            {
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

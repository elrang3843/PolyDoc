using System;
using System.Linq;
using System.Windows;
using PolyDoc.Core;
using Wpf = System.Windows.Documents;
using WpfMedia = System.Windows.Media;

namespace PolyDoc.App.Services;

/// <summary>
/// PolyDocument 를 WPF FlowDocument 로 변환한다.
/// RichTextBox.Document 에 직접 할당해 사용자가 서식 그대로 보고 편집할 수 있게 한다.
///
/// FlowDocument 가 표현 못 하는 PolyDoc 속성(예: 한글 조판의 장평·자간, Provenance 등)
/// 은 이 변환에서 누락된다. Save 시 원본 PolyDocument 를 ViewModel 이 보관하고
/// FlowDocumentParser 로 변경분만 갱신하는 식으로 보존한다.
/// </summary>
public static class FlowDocumentBuilder
{
    private const double DipsPerInch = 96.0;
    private const double PointsPerInch = 72.0;
    private const double MmPerInch = 25.4;

    public static double PtToDip(double pt) => pt * (DipsPerInch / PointsPerInch);
    public static double DipToPt(double dip) => dip * (PointsPerInch / DipsPerInch);
    public static double MmToDip(double mm) => mm * (DipsPerInch / MmPerInch);

    public static Wpf.FlowDocument Build(PolyDocument document)
    {
        ArgumentNullException.ThrowIfNull(document);

        var fd = new Wpf.FlowDocument
        {
            FontFamily = new WpfMedia.FontFamily("맑은 고딕, Malgun Gothic, Segoe UI"),
            FontSize = PtToDip(11),
            PagePadding = new Thickness(48),
        };

        foreach (var section in document.Sections)
        {
            BuildSection(fd, section);
        }
        return fd;
    }

    private static void BuildSection(Wpf.FlowDocument fd, Section section)
    {
        Wpf.List? currentList = null;
        ListKind? currentKind = null;

        foreach (var block in section.Blocks)
        {
            if (block is not Paragraph p)
            {
                continue;
            }

            if (p.Style.ListMarker is { } marker)
            {
                if (currentList is null || currentKind != marker.Kind)
                {
                    currentList = new Wpf.List
                    {
                        MarkerStyle = marker.Kind == ListKind.Bullet
                            ? Wpf.TextMarkerStyle.Disc
                            : Wpf.TextMarkerStyle.Decimal,
                    };
                    if (marker.Kind != ListKind.Bullet && marker.OrderedNumber is { } start && start >= 1)
                    {
                        currentList.StartIndex = start;
                    }
                    fd.Blocks.Add(currentList);
                    currentKind = marker.Kind;
                }

                currentList.ListItems.Add(new Wpf.ListItem(BuildParagraph(p)));
            }
            else
            {
                currentList = null;
                currentKind = null;
                fd.Blocks.Add(BuildParagraph(p));
            }
        }
    }

    private static Wpf.Paragraph BuildParagraph(Paragraph p)
    {
        var wpfPara = new Wpf.Paragraph();
        ApplyParagraphStyle(wpfPara, p.Style);
        foreach (var run in p.Runs)
        {
            wpfPara.Inlines.Add(BuildRun(run));
        }
        // 원본 PolyDoc.Paragraph 를 Tag 에 보관 — Parser 가 머지할 때 비-FlowDocument 속성 복원에 사용.
        wpfPara.Tag = p;
        return wpfPara;
    }

    private static void ApplyParagraphStyle(Wpf.Paragraph wpfPara, ParagraphStyle style)
    {
        wpfPara.TextAlignment = style.Alignment switch
        {
            Alignment.Center => TextAlignment.Center,
            Alignment.Right => TextAlignment.Right,
            Alignment.Justify or Alignment.Distributed => TextAlignment.Justify,
            _ => TextAlignment.Left,
        };

        // Heading 시각화 — 레벨별로 단락 폰트 크기·굵기를 조정. Run 단의 명시 크기는 BuildRun 에서 다시 덮어쓴다.
        if (style.Outline > OutlineLevel.Body)
        {
            wpfPara.FontSize = style.Outline switch
            {
                OutlineLevel.H1 => PtToDip(24),
                OutlineLevel.H2 => PtToDip(20),
                OutlineLevel.H3 => PtToDip(17),
                OutlineLevel.H4 => PtToDip(15),
                OutlineLevel.H5 => PtToDip(13),
                OutlineLevel.H6 => PtToDip(12),
                _ => PtToDip(11),
            };
            wpfPara.FontWeight = FontWeights.SemiBold;
        }

        var top = style.SpaceBeforePt > 0 ? PtToDip(style.SpaceBeforePt) : 0.0;
        var bottom = style.SpaceAfterPt > 0 ? PtToDip(style.SpaceAfterPt) : 0.0;
        var left = style.IndentLeftMm > 0 ? MmToDip(style.IndentLeftMm) : 0.0;
        var right = style.IndentRightMm > 0 ? MmToDip(style.IndentRightMm) : 0.0;
        if (top > 0 || bottom > 0 || left > 0 || right > 0)
        {
            wpfPara.Margin = new Thickness(left, top, right, bottom);
        }

        if (Math.Abs(style.IndentFirstLineMm) > 0.001)
        {
            wpfPara.TextIndent = MmToDip(style.IndentFirstLineMm);
        }

        // LineHeight 는 절대 DIP. 1.2 (기본) 면 명시 안 해 자연스러운 동작에 맡김.
        if (Math.Abs(style.LineHeightFactor - 1.2) > 0.01)
        {
            wpfPara.LineHeight = wpfPara.FontSize * style.LineHeightFactor;
        }
    }

    private static Wpf.Run BuildRun(Run run)
    {
        var wpfRun = new Wpf.Run(run.Text);
        var s = run.Style;

        if (!string.IsNullOrEmpty(s.FontFamily))
        {
            wpfRun.FontFamily = new WpfMedia.FontFamily(s.FontFamily);
        }
        if (Math.Abs(s.FontSizePt - 11) > 0.001)
        {
            wpfRun.FontSize = PtToDip(s.FontSizePt);
        }
        if (s.Bold)
        {
            wpfRun.FontWeight = FontWeights.Bold;
        }
        if (s.Italic)
        {
            wpfRun.FontStyle = FontStyles.Italic;
        }

        var decorations = new TextDecorationCollection();
        if (s.Underline)
        {
            foreach (var d in TextDecorations.Underline) decorations.Add(d);
        }
        if (s.Strikethrough)
        {
            foreach (var d in TextDecorations.Strikethrough) decorations.Add(d);
        }
        if (s.Overline)
        {
            foreach (var d in TextDecorations.OverLine) decorations.Add(d);
        }
        if (decorations.Count > 0)
        {
            wpfRun.TextDecorations = decorations;
        }

        if (s.Foreground is { } fg)
        {
            wpfRun.Foreground = new WpfMedia.SolidColorBrush(WpfMedia.Color.FromArgb(fg.A, fg.R, fg.G, fg.B));
        }
        if (s.Background is { } bg)
        {
            wpfRun.Background = new WpfMedia.SolidColorBrush(WpfMedia.Color.FromArgb(bg.A, bg.R, bg.G, bg.B));
        }
        if (s.Superscript)
        {
            wpfRun.BaselineAlignment = BaselineAlignment.Superscript;
        }
        else if (s.Subscript)
        {
            wpfRun.BaselineAlignment = BaselineAlignment.Subscript;
        }

        // 원본 Run 도 Tag 에 보관 (Parser 머지용).
        wpfRun.Tag = run;
        return wpfRun;
    }
}

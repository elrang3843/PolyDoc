using System;
using System.IO;
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

        var outlineStyles = document.OutlineStyles ?? OutlineStyleSet.CreateDefault();

        // 첫 번째 섹션의 PageSettings 를 FlowDocument 기본으로 사용
        var page = document.Sections.FirstOrDefault()?.Page ?? new PageSettings();

        double wDip = MmToDip(page.EffectiveWidthMm);

        var fd = new Wpf.FlowDocument
        {
            FontFamily  = new WpfMedia.FontFamily("맑은 고딕, Malgun Gothic, Segoe UI"),
            FontSize    = PtToDip(11),
            PageWidth   = wDip,
            PagePadding = new Thickness(0),
        };

        // 글자 방향 — 가로쓰기는 FlowDirection 으로 처리.
        // 세로쓰기는 WPF FlowDocument 가 native 지원하지 않아 다음 사이클에서 커스텀 레이아웃으로 도입 예정.
        // 일단 모델은 보존되며, 가로쓰기 RTL/LTR 만 시각적으로 적용된다.
        if (page.TextOrientation == TextOrientation.Horizontal)
        {
            fd.FlowDirection = page.TextProgression == TextProgression.Leftward
                ? System.Windows.FlowDirection.RightToLeft
                : System.Windows.FlowDirection.LeftToRight;
        }

        // 용지 배경색
        if (!string.IsNullOrEmpty(page.PaperColor))
        {
            try
            {
                var c = (WpfMedia.Color)WpfMedia.ColorConverter.ConvertFromString(page.PaperColor)!;
                fd.Background = new WpfMedia.SolidColorBrush(c);
            }
            catch { /* 파싱 실패 시 기본 배경 유지 */ }
        }

        // 다단 — FlowDocument.ColumnWidth 로 단 너비 지정
        // (RichTextBox 에서는 시각적 효과가 제한적이나 PageViewer/Print 에서 적용됨)
        if (page.ColumnCount > 1)
        {
            double gapDip     = MmToDip(page.ColumnGapMm);
            double contentDip = wDip - MmToDip(page.MarginLeftMm) - MmToDip(page.MarginRightMm);
            fd.ColumnWidth = Math.Max(10, (contentDip - gapDip * (page.ColumnCount - 1)) / page.ColumnCount);
            fd.ColumnGap   = gapDip;
        }

        foreach (var section in document.Sections)
        {
            BuildSection(fd, section, outlineStyles);
        }

        // RTL 모드에서는 paragraph 시작에 U+202E (RLO, Right-to-Left Override) 를
        // 박아두어 한글/라틴 같은 Bidi-약방향 문자도 강제로 RTL flow 로 표시되게 한다.
        // FlowDirection.RightToLeft 만으로는 우측 정렬만 될 뿐 글자 입력 순서가
        // 좌→우로 유지되어 사용자가 기대하는 "왼쪽 진행" 동작이 되지 않는다.
        if (fd.FlowDirection == System.Windows.FlowDirection.RightToLeft)
        {
            InjectRtlOverrideMarks(fd);
        }
        return fd;
    }

    /// <summary>RLO 마커 (U+202E). paragraph 시작에 두면 그 paragraph 내 모든 글자가
    /// 강제로 RTL 시각 흐름으로 override 된다 — 새 글자가 기존 글자의 왼쪽에 붙음.</summary>
    public const string RtlOverrideMark = "\u202E";

    private static void InjectRtlOverrideMarks(Wpf.FlowDocument fd)
    {
        foreach (var p in EnumerateParagraphs(fd.Blocks))
        {
            EnsureRloAtParagraphStart(p);
        }
    }

    private static System.Collections.Generic.IEnumerable<Wpf.Paragraph> EnumerateParagraphs(
        Wpf.BlockCollection blocks)
    {
        foreach (var b in blocks)
        {
            switch (b)
            {
                case Wpf.Paragraph p:
                    yield return p;
                    break;
                case Wpf.Section s:
                    foreach (var nested in EnumerateParagraphs(s.Blocks)) yield return nested;
                    break;
                case Wpf.List l:
                    foreach (var item in l.ListItems)
                        foreach (var nested in EnumerateParagraphs(item.Blocks)) yield return nested;
                    break;
                case Wpf.Table t:
                    foreach (var rg in t.RowGroups)
                        foreach (var row in rg.Rows)
                            foreach (var cell in row.Cells)
                                foreach (var nested in EnumerateParagraphs(cell.Blocks)) yield return nested;
                    break;
            }
        }
    }

    private static void EnsureRloAtParagraphStart(Wpf.Paragraph p)
    {
        var first = p.Inlines.FirstInline;
        if (first is Wpf.Run r)
        {
            if (r.Text.Length > 0 && r.Text[0] == '\u202E') return;
            r.Text = RtlOverrideMark + r.Text;
        }
        else if (first == null)
        {
            p.Inlines.Add(new Wpf.Run(RtlOverrideMark));
        }
        else
        {
            p.Inlines.InsertBefore(first, new Wpf.Run(RtlOverrideMark));
        }
    }

    private static void BuildSection(Wpf.FlowDocument fd, Section section, OutlineStyleSet outlineStyles)
    {
        AppendBlocks(fd.Blocks, section.Blocks, outlineStyles);
    }

    /// <summary>FlowDocument 또는 셀(TableCell) 양쪽에서 공유하는 블록 추가 로직.</summary>
    private static void AppendBlocks(System.Collections.IList target, IList<Block> blocks,
        OutlineStyleSet? outlineStyles = null)
    {
        Wpf.List? currentList = null;
        ListKind? currentKind = null;

        foreach (var block in blocks)
        {
            switch (block)
            {
                case Paragraph p when p.Style.ListMarker is { } marker:
                    if (currentList is null || currentKind != marker.Kind)
                    {
                        currentList = new Wpf.List
                        {
                            MarkerStyle = marker.Kind == ListKind.Bullet
                                ? TextMarkerStyle.Disc
                                : TextMarkerStyle.Decimal,
                        };
                        if (marker.Kind != ListKind.Bullet && marker.OrderedNumber is { } start && start >= 1)
                        {
                            currentList.StartIndex = start;
                        }
                        target.Add(currentList);
                        currentKind = marker.Kind;
                    }
                    currentList.ListItems.Add(new Wpf.ListItem(BuildParagraph(p, outlineStyles)));
                    break;

                case Paragraph p:
                    currentList = null;
                    currentKind = null;
                    target.Add(BuildParagraph(p, outlineStyles));
                    break;

                case Table t:
                    currentList = null;
                    currentKind = null;
                    target.Add(BuildTable(t, outlineStyles));
                    break;

                case ImageBlock image:
                    currentList = null;
                    currentKind = null;
                    target.Add(BuildImage(image));
                    break;

                case OpaqueBlock opaque:
                    currentList = null;
                    currentKind = null;
                    target.Add(BuildOpaquePlaceholder(opaque));
                    break;
            }
        }
    }

    private static Wpf.Table BuildTable(Table table, OutlineStyleSet? outlineStyles = null)
    {
        var wtable = new Wpf.Table { CellSpacing = 0 };
        foreach (var col in table.Columns)
        {
            var width = col.WidthMm > 0
                ? new GridLength(MmToDip(col.WidthMm))
                : GridLength.Auto;
            wtable.Columns.Add(new Wpf.TableColumn { Width = width });
        }

        var rowGroup = new Wpf.TableRowGroup();
        wtable.RowGroups.Add(rowGroup);

        foreach (var row in table.Rows)
        {
            var wrow = new Wpf.TableRow();
            foreach (var cell in row.Cells)
            {
                var wcell = new Wpf.TableCell
                {
                    BorderBrush = WpfMedia.Brushes.LightGray,
                    BorderThickness = new Thickness(1),
                    Padding = new Thickness(4),
                    ColumnSpan = Math.Max(cell.ColumnSpan, 1),
                    RowSpan = Math.Max(cell.RowSpan, 1),
                };
                AppendBlocks(wcell.Blocks, cell.Blocks, outlineStyles);
                if (wcell.Blocks.Count == 0)
                {
                    wcell.Blocks.Add(new Wpf.Paragraph(new Wpf.Run(string.Empty)));
                }
                wrow.Cells.Add(wcell);
            }
            rowGroup.Rows.Add(wrow);
        }

        wtable.Tag = table;
        return wtable;
    }

    private static Wpf.BlockUIContainer BuildImage(ImageBlock image)
    {
        var container = new Wpf.BlockUIContainer { Tag = image };

        if (image.Data.Length == 0)
        {
            container.Child = new System.Windows.Controls.TextBlock
            {
                Text = $"[이미지 누락 — {image.MediaType}]",
                Foreground = WpfMedia.Brushes.Gray,
                FontStyle = FontStyles.Italic,
            };
            return container;
        }

        var bitmap = new WpfMedia.Imaging.BitmapImage();
        bitmap.BeginInit();
        bitmap.CacheOption = WpfMedia.Imaging.BitmapCacheOption.OnLoad;
        bitmap.StreamSource = new MemoryStream(image.Data, writable: false);
        bitmap.EndInit();
        bitmap.Freeze();

        var control = new System.Windows.Controls.Image
        {
            Source = bitmap,
            Stretch = WpfMedia.Stretch.Uniform,
            HorizontalAlignment = HorizontalAlignment.Left,
        };
        if (image.WidthMm > 0)
        {
            control.Width = MmToDip(image.WidthMm);
        }
        if (image.HeightMm > 0)
        {
            control.Height = MmToDip(image.HeightMm);
        }
        if (!string.IsNullOrEmpty(image.Description))
        {
            control.ToolTip = image.Description;
        }

        container.Child = control;
        return container;
    }

    private static Wpf.Paragraph BuildOpaquePlaceholder(OpaqueBlock opaque)
    {
        // 보존 섬은 편집 불가 placeholder 로 시각화. Parser 가 Tag 에서 원본을 그대로 회수한다.
        var paragraph = new Wpf.Paragraph
        {
            Background = WpfMedia.Brushes.WhiteSmoke,
            Foreground = WpfMedia.Brushes.DimGray,
            FontStyle = FontStyles.Italic,
            BorderBrush = WpfMedia.Brushes.LightGray,
            BorderThickness = new Thickness(1),
            Padding = new Thickness(8, 4, 8, 4),
            Tag = opaque,
        };
        paragraph.Inlines.Add(new Wpf.Run(opaque.DisplayLabel));
        return paragraph;
    }

    private static Wpf.Paragraph BuildParagraph(Paragraph p, OutlineStyleSet? outlineStyles = null)
    {
        var wpfPara = new Wpf.Paragraph();
        ApplyParagraphStyle(wpfPara, p.Style, outlineStyles);
        foreach (var run in p.Runs)
        {
            wpfPara.Inlines.Add(BuildInline(run));
        }
        // 원본 PolyDoc.Paragraph 를 Tag 에 보관 — Parser 가 머지할 때 비-FlowDocument 속성 복원에 사용.
        wpfPara.Tag = p;
        return wpfPara;
    }

    private static void ApplyParagraphStyle(Wpf.Paragraph wpfPara, ParagraphStyle style,
        OutlineStyleSet? outlineStyles = null)
    {
        wpfPara.TextAlignment = style.Alignment switch
        {
            Alignment.Center => TextAlignment.Center,
            Alignment.Right => TextAlignment.Right,
            Alignment.Justify or Alignment.Distributed => TextAlignment.Justify,
            _ => TextAlignment.Left,
        };

        // 개요 수준이 있으면 OutlineStyleSet 에서 글자 크기·굵기 읽기 (없으면 내장 기본값).
        if (style.Outline > OutlineLevel.Body)
        {
            var ls = outlineStyles?.GetLevel(style.Outline) ?? OutlineStyleSet.DefaultForLevel(style.Outline);
            var charStyle = ls.Char;
            wpfPara.FontSize   = PtToDip(charStyle.FontSizePt > 0 ? charStyle.FontSizePt : 11);
            wpfPara.FontWeight = charStyle.Bold ? FontWeights.Bold : FontWeights.SemiBold;
            if (!string.IsNullOrEmpty(charStyle.FontFamily))
                wpfPara.FontFamily = new WpfMedia.FontFamily(charStyle.FontFamily);
            if (charStyle.Italic)
                wpfPara.FontStyle = FontStyles.Italic;
            if (charStyle.Foreground is { } fg)
                wpfPara.Foreground = new WpfMedia.SolidColorBrush(
                    WpfMedia.Color.FromArgb(fg.A, fg.R, fg.G, fg.B));
            if (ls.BackgroundColor is { } bgHex)
            {
                try
                {
                    var bgc = (WpfMedia.Color)WpfMedia.ColorConverter.ConvertFromString(bgHex);
                    wpfPara.Background = new WpfMedia.SolidColorBrush(bgc);
                }
                catch { }
            }
            if (ls.Border.ShowTop || ls.Border.ShowBottom)
            {
                WpfMedia.SolidColorBrush borderBrush;
                if (!string.IsNullOrEmpty(ls.Border.Color))
                {
                    try { borderBrush = new WpfMedia.SolidColorBrush((WpfMedia.Color)WpfMedia.ColorConverter.ConvertFromString(ls.Border.Color)); }
                    catch { borderBrush = WpfMedia.Brushes.DimGray; }
                }
                else
                {
                    borderBrush = WpfMedia.Brushes.DimGray;
                }
                wpfPara.BorderBrush     = borderBrush;
                wpfPara.BorderThickness = new Thickness(0,
                    ls.Border.ShowTop    ? 1 : 0, 0,
                    ls.Border.ShowBottom ? 1 : 0);
            }
            // Para 공간 설정은 OutlineStyle 의 Para 를 우선하되, ParagraphStyle 직접 값이 0이 아니면 덮어씀
            var paraStyle = ls.Para;
            var top    = style.SpaceBeforePt > 0 ? PtToDip(style.SpaceBeforePt)
                        : paraStyle.SpaceBeforePt > 0 ? PtToDip(paraStyle.SpaceBeforePt) : 0.0;
            var bottom = style.SpaceAfterPt  > 0 ? PtToDip(style.SpaceAfterPt)
                        : paraStyle.SpaceAfterPt  > 0 ? PtToDip(paraStyle.SpaceAfterPt)  : 0.0;
            var left   = style.IndentLeftMm  > 0 ? MmToDip(style.IndentLeftMm)  : 0.0;
            var right  = style.IndentRightMm > 0 ? MmToDip(style.IndentRightMm) : 0.0;
            if (top > 0 || bottom > 0 || left > 0 || right > 0)
                wpfPara.Margin = new Thickness(left, top, right, bottom);

            var lhf = style.LineHeightFactor != 1.2 ? style.LineHeightFactor : paraStyle.LineHeightFactor;
            if (Math.Abs(lhf - 1.2) > 0.01)
                wpfPara.LineHeight = wpfPara.FontSize * lhf;

            if (Math.Abs(style.IndentFirstLineMm) > 0.001)
                wpfPara.TextIndent = MmToDip(style.IndentFirstLineMm);
            return;
        }

        // 본문 (Body) 처리 — OutlineStyle 의 본문 스타일도 적용
        if (outlineStyles != null)
        {
            var bodyLs = outlineStyles.GetLevel(OutlineLevel.Body);
            var bc = bodyLs.Char;
            if (bc.FontSizePt > 0 && Math.Abs(bc.FontSizePt - 11) > 0.01)
                wpfPara.FontSize = PtToDip(bc.FontSizePt);
            if (!string.IsNullOrEmpty(bc.FontFamily))
                wpfPara.FontFamily = new WpfMedia.FontFamily(bc.FontFamily);
            var bpLhf = bodyLs.Para.LineHeightFactor;
            if (Math.Abs(bpLhf - 1.2) > 0.01 && Math.Abs(style.LineHeightFactor - 1.2) < 0.01)
                wpfPara.LineHeight = wpfPara.FontSize * bpLhf;
        }

        var sTop = style.SpaceBeforePt > 0 ? PtToDip(style.SpaceBeforePt) : 0.0;
        var sBottom = style.SpaceAfterPt > 0 ? PtToDip(style.SpaceAfterPt) : 0.0;
        var sLeft = style.IndentLeftMm > 0 ? MmToDip(style.IndentLeftMm) : 0.0;
        var sRight = style.IndentRightMm > 0 ? MmToDip(style.IndentRightMm) : 0.0;
        if (sTop > 0 || sBottom > 0 || sLeft > 0 || sRight > 0)
            wpfPara.Margin = new Thickness(sLeft, sTop, sRight, sBottom);

        if (Math.Abs(style.IndentFirstLineMm) > 0.001)
            wpfPara.TextIndent = MmToDip(style.IndentFirstLineMm);

        if (Math.Abs(style.LineHeightFactor - 1.2) > 0.01)
            wpfPara.LineHeight = wpfPara.FontSize * style.LineHeightFactor;
    }

    /// <summary>글자폭 != 100% 또는 자간 != 0 이면 Span(per-char InlineUIContainer 들), 그 외에는 Run 반환.</summary>
    public static Wpf.Inline BuildInline(Run run)
    {
        var s = run.Style;
        if (NeedsContainer(s))
            return BuildScaledContainer(run);

        var wpfRun = new Wpf.Run(run.Text);

        if (!string.IsNullOrEmpty(s.FontFamily))
            wpfRun.FontFamily = new WpfMedia.FontFamily(s.FontFamily);
        if (Math.Abs(s.FontSizePt - 11) > 0.001)
            wpfRun.FontSize = PtToDip(s.FontSizePt);
        if (s.Bold)
            wpfRun.FontWeight = FontWeights.Bold;
        if (s.Italic)
            wpfRun.FontStyle = FontStyles.Italic;

        var decorations = new TextDecorationCollection();
        if (s.Underline) foreach (var d in TextDecorations.Underline) decorations.Add(d);
        if (s.Strikethrough) foreach (var d in TextDecorations.Strikethrough) decorations.Add(d);
        if (s.Overline) foreach (var d in TextDecorations.OverLine) decorations.Add(d);
        if (decorations.Count > 0)
            wpfRun.TextDecorations = decorations;

        if (s.Foreground is { } fg)
            wpfRun.Foreground = new WpfMedia.SolidColorBrush(WpfMedia.Color.FromArgb(fg.A, fg.R, fg.G, fg.B));
        if (s.Background is { } bg)
            wpfRun.Background = new WpfMedia.SolidColorBrush(WpfMedia.Color.FromArgb(bg.A, bg.R, bg.G, bg.B));

        if (s.Superscript)
            wpfRun.BaselineAlignment = BaselineAlignment.Superscript;
        else if (s.Subscript)
            wpfRun.BaselineAlignment = BaselineAlignment.Subscript;

        wpfRun.Tag = run;
        return wpfRun;
    }

    private static bool NeedsContainer(RunStyle s)
        => Math.Abs(s.WidthPercent - 100) > 0.5 || Math.Abs(s.LetterSpacingPx) > 0.01;

    /// <summary>
    /// 글자폭·자간을 시각화. WPF 의 Run 은 LayoutTransform/RenderTransform 을 직접 지원하지 않으므로
    /// InlineUIContainer 가 필요하다. 단, 한 Run 전체를 하나의 IUC 로 감싸면 atomic 요소가 되어
    /// 선택이 통째로 묶여 캐럿이 안으로 못 들어가는 UX 문제가 생긴다.
    /// 그래서 문자별로 IUC 를 분리하고 같은 부모 Span 아래에 묶어, WPF 가 IUC 사이에
    /// 캐럿 위치·줄바꿈·문자 단위 선택을 정상적으로 처리하게 한다.
    /// Span.Tag, 각 IUC.Tag 모두 원본 PolyDoc.Run 을 가리켜 라운드트립 머지의 단서가 된다.
    /// </summary>
    public static Wpf.Span BuildScaledContainer(Run run)
    {
        var s = run.Style;
        var fontSize = PtToDip(s.FontSizePt > 0 ? s.FontSizePt : 11);
        var span = new Wpf.Span { Tag = run };

        var text = run.Text.Length > 0 ? run.Text : " ";
        bool hasSpacing = Math.Abs(s.LetterSpacingPx) > 0.01;
        for (int i = 0; i < text.Length; i++)
        {
            var tb = BuildCharTextBlock(text[i].ToString(), s, fontSize);
            // 마지막 문자 뒤 자간은 영역 끝의 군더더기 — 제거.
            if (hasSpacing && i == text.Length - 1)
                tb.Margin = new Thickness(0);
            span.Inlines.Add(new Wpf.InlineUIContainer(tb)
            {
                BaselineAlignment = BaselineAlignment.Baseline,
                Tag = run,
            });
        }
        return span;
    }

    private static System.Windows.Controls.TextBlock BuildCharTextBlock(string ch, RunStyle s, double fontSize)
    {
        var tb = new System.Windows.Controls.TextBlock
        {
            Text = ch,
            FontSize = fontSize,
            Margin = new Thickness(0, 0, s.LetterSpacingPx, 0),
        };
        if (Math.Abs(s.WidthPercent - 100) > 0.5)
            tb.LayoutTransform = new WpfMedia.ScaleTransform(s.WidthPercent / 100.0, 1.0);
        ApplyStyleToTextBlock(tb, s);
        return tb;
    }

    private static void ApplyStyleToTextBlock(System.Windows.Controls.TextBlock tb, RunStyle s)
    {
        if (!string.IsNullOrEmpty(s.FontFamily)) tb.FontFamily = new WpfMedia.FontFamily(s.FontFamily);
        if (s.Bold) tb.FontWeight = FontWeights.Bold;
        if (s.Italic) tb.FontStyle = FontStyles.Italic;

        var decos = new TextDecorationCollection();
        if (s.Underline) foreach (var d in TextDecorations.Underline) decos.Add(d);
        if (s.Strikethrough) foreach (var d in TextDecorations.Strikethrough) decos.Add(d);
        if (s.Overline) foreach (var d in TextDecorations.OverLine) decos.Add(d);
        if (decos.Count > 0) tb.TextDecorations = decos;

        if (s.Foreground is { } fg)
            tb.Foreground = new WpfMedia.SolidColorBrush(WpfMedia.Color.FromArgb(fg.A, fg.R, fg.G, fg.B));
        if (s.Background is { } bg)
            tb.Background = new WpfMedia.SolidColorBrush(WpfMedia.Color.FromArgb(bg.A, bg.R, bg.G, bg.B));
    }
}

using System;
using System.IO;
using System.Linq;
using System.Windows;
using PolyDonky.Core;
using Wpf = System.Windows.Documents;
using WpfMedia = System.Windows.Media;

namespace PolyDonky.App.Services;

/// <summary>
/// PolyDonkyument 를 WPF FlowDocument 로 변환한다.
/// RichTextBox.Document 에 직접 할당해 사용자가 서식 그대로 보고 편집할 수 있게 한다.
///
/// FlowDocument 가 표현 못 하는 PolyDonky 속성(예: 한글 조판의 장평·자간, Provenance 등)
/// 은 이 변환에서 누락된다. Save 시 원본 PolyDonkyument 를 ViewModel 이 보관하고
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
    public static double DipToMm(double dip) => dip * (MmPerInch / DipsPerInch);

    public static Wpf.FlowDocument Build(PolyDonkyument document)
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

        // 글자 방향(세로쓰기 / 왼쪽으로 진행 등)은 추후 지원 예정.
        // 현재는 항상 LTR 가로쓰기로 표시하며, 모델의 TextOrientation/TextProgression 값은 보존만 한다.

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

        return fd;
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

    /// <summary>
    /// ImageBlock 을 WPF Block 으로 빌드한다.
    /// - WrapMode = Inline                → BlockUIContainer (자체 줄 차지, 가로 정렬만 적용)
    /// - WrapMode = WrapLeft              → Paragraph + Floater(왼쪽), 텍스트가 오른쪽으로 흐름
    /// - WrapMode = WrapRight             → Paragraph + Floater(오른쪽), 텍스트가 왼쪽으로 흐름
    /// - WrapMode = InFrontOfText/BehindText → 빈 placeholder Paragraph (실제 렌더링은 MainWindow 의
    ///                                        OverlayImageCanvas/UnderlayImageCanvas 에서 절대 위치로 처리)
    /// 반환된 Block 의 Tag 에 ImageBlock 이 심어져 라운드트립과 우클릭 라우팅에 사용된다.
    /// </summary>
    internal static Wpf.Block BuildImage(ImageBlock image)
    {
        // ── 오버레이 모드 — 본문 흐름에는 위치만 차지하고 실제 그림은 캔버스에서 ──
        if (image.WrapMode is ImageWrapMode.InFrontOfText or ImageWrapMode.BehindText)
        {
            // 투명·최소 높이 플레이스홀더 단락.
            // - FontSize = 0.1 pt → 행 높이를 0에 가깝게 압축해 본문 레이아웃 영향 최소화.
            // - Foreground/Background Transparent → 선택 하이라이트(파란 줄)가 시각적으로 안 보임.
            // - IsEnabled = false 는 Paragraph 에 없으므로 Focusable 을 강제 불가 — 대신 크기로 억제.
            return new Wpf.Paragraph
            {
                Tag        = image,
                Margin     = new Thickness(0),
                FontSize   = 0.1,
                Foreground = WpfMedia.Brushes.Transparent,
                Background = WpfMedia.Brushes.Transparent,
            };
        }

        // ── 미사용 Image 폴백 ────────────────────────────────────────
        if (image.Data.Length == 0)
        {
            var emptyBuc = new Wpf.BlockUIContainer { Tag = image };
            emptyBuc.Child = new System.Windows.Controls.TextBlock
            {
                Text = $"[이미지 누락 — {image.MediaType}]",
                Foreground = WpfMedia.Brushes.Gray,
                FontStyle = FontStyles.Italic,
            };
            return emptyBuc;
        }

        var bitmap = new WpfMedia.Imaging.BitmapImage();
        bitmap.BeginInit();
        bitmap.CacheOption  = WpfMedia.Imaging.BitmapCacheOption.OnLoad;
        bitmap.StreamSource = new MemoryStream(image.Data, writable: false);
        bitmap.EndInit();
        bitmap.Freeze();

        // Image.Tag 에 container 를 저장하지 말 것 — container.Child = image 와 함께 순환 참조가 되어
        // WPF undo 스냅샷의 XamlWriter.Save() 가 StackOverflowException 으로 폭주한다.
        // 우클릭 속성 라우팅은 LogicalTreeHelper 로 image 의 부모를 찾는 방식으로 처리한다.
        var control = new System.Windows.Controls.Image
        {
            Source  = bitmap,
            Stretch = WpfMedia.Stretch.Uniform,
        };
        if (image.WidthMm > 0)  control.Width  = MmToDip(image.WidthMm);
        if (image.HeightMm > 0) control.Height = MmToDip(image.HeightMm);
        if (!string.IsNullOrEmpty(image.Description)) control.ToolTip = image.Description;

        // 테두리 래퍼
        UIElement visual = control;
        if (!string.IsNullOrEmpty(image.BorderColor) && image.BorderThicknessPt > 0)
        {
            WpfMedia.Brush borderBrush;
            try { borderBrush = new WpfMedia.SolidColorBrush(
                    (WpfMedia.Color)WpfMedia.ColorConverter.ConvertFromString(image.BorderColor)!); }
            catch { borderBrush = WpfMedia.Brushes.Black; }

            visual = new System.Windows.Controls.Border
            {
                Child           = control,
                BorderBrush     = borderBrush,
                BorderThickness = new Thickness(PtToDip(image.BorderThicknessPt)),
            };
        }

        var marginTopDip    = MmToDip(image.MarginTopMm);
        var marginBottomDip = MmToDip(image.MarginBottomMm);

        // ── 래핑 없음(Inline) — BlockUIContainer 를 그대로 추가 ────────
        if (image.WrapMode == ImageWrapMode.Inline)
        {
            return new Wpf.BlockUIContainer(visual)
            {
                Tag           = image,
                Margin        = new Thickness(0, marginTopDip, 0, marginBottomDip),
                TextAlignment = image.HAlign switch
                {
                    ImageHAlign.Center => TextAlignment.Center,
                    ImageHAlign.Right  => TextAlignment.Right,
                    _                  => TextAlignment.Left,
                },
            };
        }

        // ── 텍스트 캐릭터처럼(AsText) — Paragraph 안에 InlineUIContainer 로 ────
        // 한 단락 안에서 글자처럼 흐르므로 같은 단락의 텍스트 런과 같은 줄에 들어갈 수 있다.
        // 사용자는 이후 이 단락에 텍스트를 추가해 그림과 같은 줄에 글자를 둘 수 있다.
        if (image.WrapMode == ImageWrapMode.AsText)
        {
            var asTextPara = new Wpf.Paragraph
            {
                Tag           = image,
                Margin        = new Thickness(0, marginTopDip, 0, marginBottomDip),
                TextAlignment = image.HAlign switch
                {
                    ImageHAlign.Center => TextAlignment.Center,
                    ImageHAlign.Right  => TextAlignment.Right,
                    _                  => TextAlignment.Left,
                },
            };
            asTextPara.Inlines.Add(new Wpf.InlineUIContainer(visual)
            {
                BaselineAlignment = BaselineAlignment.Bottom,
            });
            return asTextPara;
        }

        // ── 래핑 있음(WrapLeft/WrapRight) — Floater 가 든 Paragraph ────
        // Floater 는 Inline 이라 Paragraph 안에 들어가야 하며,
        // 인접한 본문 Paragraph 와 같은 흐름 안에 있어야 텍스트가 주변으로 흐른다.
        var floater = new Wpf.Floater
        {
            HorizontalAlignment = image.WrapMode == ImageWrapMode.WrapRight
                ? HorizontalAlignment.Right
                : HorizontalAlignment.Left,
            Padding = new Thickness(0),
            Margin  = new Thickness(
                image.WrapMode == ImageWrapMode.WrapRight ? PtToDip(8) : 0,   // 텍스트와의 좌측 간격
                marginTopDip,
                image.WrapMode == ImageWrapMode.WrapLeft  ? PtToDip(8) : 0,   // 텍스트와의 우측 간격
                marginBottomDip),
        };
        if (image.WidthMm > 0) floater.Width = MmToDip(image.WidthMm);
        floater.Blocks.Add(new Wpf.BlockUIContainer(visual));

        // Paragraph 는 비어 있어도 되지만, Tag 로 ImageBlock 을 보존해 라운드트립 가능.
        // 안에 Floater 만 두고 텍스트가 없으면 다음 Paragraph 의 텍스트가 자동으로 주변에 흐른다.
        // FontSize=0.1 + Transparent 로 선택 하이라이트(파란 가로줄)가 보이지 않도록 억제.
        var paragraph = new Wpf.Paragraph(floater)
        {
            Tag        = image,
            Margin     = new Thickness(0),
            FontSize   = 0.1,
            Foreground = WpfMedia.Brushes.Transparent,
            Background = WpfMedia.Brushes.Transparent,
        };
        return paragraph;
    }

    /// <summary>
    /// ImageBlock 으로부터 캔버스 오버레이용 Image 컨트롤을 생성한다.
    /// MainWindow 가 InFrontOfText/BehindText 모드 그림을 OverlayImageCanvas/UnderlayImageCanvas 에 배치할 때 사용.
    /// 테두리·크기·툴팁(설명)을 적용하지만, 위치(OverlayXMm/OverlayYMm)는 호출측에서 Canvas.Left/Top 으로 설정.
    /// </summary>
    public static System.Windows.FrameworkElement? BuildOverlayImageControl(ImageBlock image)
    {
        if (image.Data.Length == 0) return null;

        var bitmap = new WpfMedia.Imaging.BitmapImage();
        bitmap.BeginInit();
        bitmap.CacheOption  = WpfMedia.Imaging.BitmapCacheOption.OnLoad;
        bitmap.StreamSource = new MemoryStream(image.Data, writable: false);
        bitmap.EndInit();
        bitmap.Freeze();

        var control = new System.Windows.Controls.Image
        {
            Source  = bitmap,
            Stretch = WpfMedia.Stretch.Fill,
        };
        if (image.WidthMm > 0)  control.Width  = MmToDip(image.WidthMm);
        if (image.HeightMm > 0) control.Height = MmToDip(image.HeightMm);
        if (!string.IsNullOrEmpty(image.Description)) control.ToolTip = image.Description;

        if (!string.IsNullOrEmpty(image.BorderColor) && image.BorderThicknessPt > 0)
        {
            WpfMedia.Brush borderBrush;
            try { borderBrush = new WpfMedia.SolidColorBrush(
                    (WpfMedia.Color)WpfMedia.ColorConverter.ConvertFromString(image.BorderColor)!); }
            catch { borderBrush = WpfMedia.Brushes.Black; }

            return new System.Windows.Controls.Border
            {
                Child           = control,
                BorderBrush     = borderBrush,
                BorderThickness = new Thickness(PtToDip(image.BorderThicknessPt)),
            };
        }

        return control;
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
        // 원본 PolyDonky.Paragraph 를 Tag 에 보관 — Parser 가 머지할 때 비-FlowDocument 속성 복원에 사용.
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

    /// <summary>글자폭 != 100% 또는 자간 != 0 이면 Span(per-char InlineUIContainer 들), 그 외에는 Run 반환.
    /// LatexSource 가 있으면 WpfMath FormulaControl 로 렌더링.
    /// EmojiKey 가 있으면 Resources/Emojis/{Section}/{name}.png 를 Image 로 렌더링.</summary>
    public static Wpf.Inline BuildInline(Run run)
    {
        if (run.LatexSource is { Length: > 0 } latex)
            return BuildEquationInline(run, latex);

        if (run.EmojiKey is { Length: > 0 } emojiKey)
            return BuildEmojiInline(run, emojiKey);

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

    private static Wpf.Inline BuildEquationInline(Run run, string latex)
    {
        var img = RenderEquationToImage(latex, run.IsDisplayEquation ? 18.0 : 14.0);
        if (img is null)
            return new Wpf.Run(run.Text) { Tag = run };

        return new Wpf.InlineUIContainer(img)
        {
            Tag               = run,
            BaselineAlignment = BaselineAlignment.Center,
        };
    }

    /// <summary>
    /// 이모지 PNG 를 Image 로 렌더링. 키 형식: "{Section}_{name}" → pack URI 로 해석.
    /// 본문 폰트 크기에 비례한 정사각 크기로 표시 (기본 ~1.4em).
    /// 로드 실패 시 plain-text Run("[Section_name]") 으로 폴백 — 라운드트립용 EmojiKey 는 Tag 로 보존.
    /// </summary>
    private static Wpf.Inline BuildEmojiInline(Run run, string emojiKey)
    {
        double sizePt  = run.Style.FontSizePt > 0 ? run.Style.FontSizePt : 16.0;
        double sizeDip = PtToDip(sizePt);

        var img = LoadEmojiImage(emojiKey, sizeDip);
        if (img is null)
            return new Wpf.Run($"[{emojiKey}]") { Tag = run };

        // img.Tag 에 iuc 를 저장하지 말 것 — iuc.Child = img 와 함께 순환 참조가 되어
        // WPF undo 의 XamlWriter.Save() 가 StackOverflow 로 폭주한다 (수식 IUC 와 동일 이슈).
        return new Wpf.InlineUIContainer(img)
        {
            Tag               = run,
            BaselineAlignment = run.EmojiAlignment switch
            {
                EmojiAlignment.TextTop    => BaselineAlignment.TextTop,
                EmojiAlignment.TextBottom => BaselineAlignment.TextBottom,
                EmojiAlignment.Baseline   => BaselineAlignment.Baseline,
                _                         => BaselineAlignment.Center,
            },
        };
    }

    /// <summary>
    /// EmojiKey ("{Section}_{name}") → pack URI Image. 키가 잘못됐거나 리소스가 없으면 null.
    /// </summary>
    public static System.Windows.Controls.Image? LoadEmojiImage(string emojiKey, double sizeDip)
    {
        var (section, name) = SplitEmojiKey(emojiKey);
        if (section is null || name is null) return null;

        try
        {
            var uri = new Uri($"pack://application:,,,/Resources/Emojis/{section}/{name}.png", UriKind.Absolute);
            var bmp = new WpfMedia.Imaging.BitmapImage();
            bmp.BeginInit();
            bmp.UriSource   = uri;
            bmp.CacheOption = WpfMedia.Imaging.BitmapCacheOption.OnLoad;
            bmp.EndInit();
            bmp.Freeze();
            return new System.Windows.Controls.Image
            {
                Source  = bmp,
                Width   = sizeDip,
                Height  = sizeDip,
                Stretch = WpfMedia.Stretch.Uniform,
            };
        }
        catch
        {
            return null;
        }
    }

    /// <summary>"Status_done" → ("Status", "done"). 잘못된 형식이면 (null, null).
    /// 이름에 밑줄이 포함될 수 있으므로 첫 번째 '_' 만 분리자로 사용한다 (예: "Status_in_progress").</summary>
    private static (string? Section, string? Name) SplitEmojiKey(string key)
    {
        int idx = key.IndexOf('_');
        if (idx <= 0 || idx >= key.Length - 1) return (null, null);
        return (key[..idx], key[(idx + 1)..]);
    }

    /// <summary>
    /// FormulaControl 을 비주얼 트리 없이 오프스크린 렌더링하여 Image 로 반환.
    /// Image(BitmapSource) 는 XamlWriter.Save() 에 안전하므로 RichTextBox undo 와 충돌하지 않는다.
    /// 파싱·렌더링 실패 시 null 반환.
    /// </summary>
    public static System.Windows.Controls.Image? RenderEquationToImage(string latex, double scale)
    {
        try
        {
            var formula = new WpfMath.Controls.FormulaControl
            {
                Formula = latex,
                Scale   = scale,
            };

            // 비주얼 트리 없이 레이아웃 패스를 강제 실행
            formula.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
            formula.Arrange(new Rect(formula.DesiredSize));

            int w = Math.Max(1, (int)Math.Ceiling(formula.ActualWidth));
            int h = Math.Max(1, (int)Math.Ceiling(formula.ActualHeight));

            var rtb = new WpfMedia.Imaging.RenderTargetBitmap(w, h, 96, 96, WpfMedia.PixelFormats.Pbgra32);
            rtb.Render(formula);
            rtb.Freeze();

            return new System.Windows.Controls.Image
            {
                Source  = rtb,
                Width   = w,
                Height  = h,
                Stretch = WpfMedia.Stretch.None,
            };
        }
        catch
        {
            return null;
        }
    }

    private static bool NeedsContainer(RunStyle s)
        => Math.Abs(s.WidthPercent - 100) > 0.5 || Math.Abs(s.LetterSpacingPx) > 0.01;

    /// <summary>
    /// 글자폭·자간을 시각화. WPF 의 Run 은 LayoutTransform/RenderTransform 을 직접 지원하지 않으므로
    /// InlineUIContainer 가 필요하다. 단, 한 Run 전체를 하나의 IUC 로 감싸면 atomic 요소가 되어
    /// 선택이 통째로 묶여 캐럿이 안으로 못 들어가는 UX 문제가 생긴다.
    /// 그래서 문자별로 IUC 를 분리하고 같은 부모 Span 아래에 묶어, WPF 가 IUC 사이에
    /// 캐럿 위치·줄바꿈·문자 단위 선택을 정상적으로 처리하게 한다.
    /// Span.Tag, 각 IUC.Tag 모두 원본 PolyDonky.Run 을 가리켜 라운드트립 머지의 단서가 된다.
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

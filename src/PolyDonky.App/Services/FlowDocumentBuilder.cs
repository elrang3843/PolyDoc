using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using PolyDonky.Core;
using Wpf = System.Windows.Documents;
using WpfMedia = System.Windows.Media;
using WpfShapes = System.Windows.Shapes;

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

    /// <summary>
    /// FlowDocument 레이아웃 폭 = 종이 폭 − 좌여백 − 우여백 (최소 10 DIP).
    /// BodyEditor.Padding 이 좌우 여백을 담당하므로 FlowDocument 는 본문 폭만 책임진다.
    /// ApplyPageSettings 에서도 동일 공식으로 Document.PageWidth 를 갱신해야 한다.
    /// </summary>
    public static double ComputeContentWidthDip(PageSettings page)
    {
        double paperDip = MmToDip(page.EffectiveWidthMm);
        double leftDip  = MmToDip(page.MarginLeftMm);
        double rightDip = MmToDip(page.MarginRightMm);
        return Math.Max(10.0, paperDip - leftDip - rightDip);
    }

    public static Wpf.FlowDocument Build(PolyDonkyument document)
    {
        ArgumentNullException.ThrowIfNull(document);

        var outlineStyles = document.OutlineStyles ?? OutlineStyleSet.CreateDefault();

        // 첫 번째 섹션의 PageSettings 를 FlowDocument 기본으로 사용
        var page = document.Sections.FirstOrDefault()?.Page ?? new PageSettings();

        double wDip       = MmToDip(page.EffectiveWidthMm);
        // BodyEditor.Padding 이 좌우 여백을 차지하므로 FlowDocument 의 레이아웃 폭은
        // 본문 폭(종이 폭 − 좌여백 − 우여백)으로 설정해야 한다.
        // PageWidth = 종이 전체 폭으로 두면 HorizontalAlignment.Right Floater 를 비롯한
        // 모든 우측 정렬 객체가 '우측 여백' 만큼 오른쪽으로 밀려 클리핑된다.
        double contentWDip = ComputeContentWidthDip(page);

        var fd = new Wpf.FlowDocument
        {
            FontFamily  = new WpfMedia.FontFamily("맑은 고딕, Malgun Gothic, Segoe UI"),
            FontSize    = PtToDip(11),
            PageWidth   = contentWDip,
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
        if (page.ColumnCount > 1)
        {
            double gapDip = MmToDip(page.ColumnGapMm);
            fd.ColumnWidth = Math.Max(10, (contentWDip - gapDip * (page.ColumnCount - 1)) / page.ColumnCount);
            fd.ColumnGap   = gapDip;
            fd.IsColumnWidthFlexible = false;
        }


        var fnNums = document.Footnotes.Count > 0
            ? document.Footnotes.Select((f, i) => (f.Id, i + 1)).ToDictionary(x => x.Id, x => x.Item2)
            : null;
        var enNums = document.Endnotes.Count > 0
            ? document.Endnotes.Select((e, i) => (e.Id, i + 1)).ToDictionary(x => x.Id, x => x.Item2)
            : null;

        foreach (var section in document.Sections)
        {
            BuildSection(fd, section, outlineStyles, fnNums, enNums);
        }

        return fd;
    }

    private static void BuildSection(Wpf.FlowDocument fd, Section section, OutlineStyleSet outlineStyles,
        IReadOnlyDictionary<string, int>? fnNums = null, IReadOnlyDictionary<string, int>? enNums = null)
    {
        AppendBlocks(fd.Blocks, section.Blocks, outlineStyles, fnNums, enNums);
    }

    /// <summary>
    /// 지정한 Core.Block 목록만 포함하는 FlowDocument 를 빌드한다.
    /// per-page 편집기·per-page RTB 측정에 사용. STA 스레드 필수.
    /// </summary>
    internal static Wpf.FlowDocument BuildFromBlocks(
        IEnumerable<Block> blocks,
        PageSettings?      page          = null,
        OutlineStyleSet?   outlineStyles = null)
    {
        page          ??= new PageSettings();
        outlineStyles ??= OutlineStyleSet.CreateDefault();

        double contentWDip = ComputeContentWidthDip(page);
        var fd = new Wpf.FlowDocument
        {
            FontFamily  = new WpfMedia.FontFamily("맑은 고딕, Malgun Gothic, Segoe UI"),
            FontSize    = PtToDip(11),
            PageWidth   = contentWDip,
            PagePadding = new Thickness(0),
        };

        if (!string.IsNullOrEmpty(page.PaperColor))
        {
            try
            {
                var c = (WpfMedia.Color)WpfMedia.ColorConverter.ConvertFromString(page.PaperColor)!;
                fd.Background = new WpfMedia.SolidColorBrush(c);
            }
            catch { }
        }

        // 다단 — per-page RTB 도 본문 폭 안에서 단을 분할한다.
        if (page.ColumnCount > 1)
        {
            double gapDip = MmToDip(page.ColumnGapMm);
            fd.ColumnWidth = Math.Max(10, (contentWDip - gapDip * (page.ColumnCount - 1)) / page.ColumnCount);
            fd.ColumnGap   = gapDip;
            fd.IsColumnWidthFlexible = false;
        }

        AppendBlocks(fd.Blocks, blocks.ToList(), outlineStyles);
        return fd;
    }

    /// <summary>FlowDocument 또는 셀(TableCell) 양쪽에서 공유하는 블록 추가 로직.</summary>
    internal static void AppendBlocks(System.Collections.IList target, IList<Block> blocks,
        OutlineStyleSet? outlineStyles = null,
        IReadOnlyDictionary<string, int>? fnNums = null,
        IReadOnlyDictionary<string, int>? enNums = null)
    {
        // 중첩 리스트 지원: (WPF List, Kind) 스택.
        // 인덱스 0 = 최상위 리스트, 인덱스 n = n 단계 중첩 리스트.
        // 비-리스트 블록이 오면 스택을 비워 리스트 컨텍스트를 종료한다.
        var listStack = new Stack<(Wpf.List List, ListKind Kind)>();

        // 지정 level·kind 의 WPF List 를 반환한다.
        // 스택이 부족하면 새 List 를 생성해 부모 ListItem 에 붙이고 push 한다.
        Wpf.List EnsureList(int level, ListKind kind, int startIndex)
        {
            // 초과 레벨 팝
            while (listStack.Count > level + 1)
                listStack.Pop();

            // 현재 레벨에 같은 Kind 리스트가 있으면 재사용
            if (listStack.Count == level + 1 && listStack.Peek().Kind == kind)
                return listStack.Peek().List;

            // 현재 레벨에 다른 Kind 리스트가 있으면 교체 준비
            if (listStack.Count == level + 1)
                listStack.Pop();

            // 레벨이 스택보다 깊으면 중간 레벨을 채운다 (정상 HTML 에선 발생하지 않음)
            while (listStack.Count < level)
            {
                var mid = new Wpf.List { MarkerStyle = TextMarkerStyle.Disc };
                AppendListToParent(mid);
                listStack.Push((mid, ListKind.Bullet));
            }

            var newList = new Wpf.List
            {
                MarkerStyle = kind == ListKind.Bullet ? TextMarkerStyle.Disc : TextMarkerStyle.Decimal,
            };
            if (kind != ListKind.Bullet && startIndex >= 1)
                newList.StartIndex = startIndex;
            AppendListToParent(newList);
            listStack.Push((newList, kind));
            return newList;
        }

        // 새 WPF List 를 최상위(target) 또는 부모 ListItem 의 Blocks 에 붙인다.
        void AppendListToParent(Wpf.List newList)
        {
            if (listStack.Count == 0)
            {
                target.Add(newList);
            }
            else
            {
                var parentList = listStack.Peek().List;
                if (parentList.ListItems.Count > 0)
                {
                    parentList.ListItems.Cast<Wpf.ListItem>().Last().Blocks.Add(newList);
                }
                else
                {
                    var stub = new Wpf.ListItem();
                    parentList.ListItems.Add(stub);
                    stub.Blocks.Add(newList);
                }
            }
        }

        foreach (var block in blocks)
        {
            switch (block)
            {
                case Paragraph p when p.Style.ListMarker is { } marker:
                {
                    int level = Math.Max(0, marker.Level);
                    int start = marker.Kind != ListKind.Bullet && marker.OrderedNumber is { } s && s >= 1 ? s : 1;
                    var list  = EnsureList(level, marker.Kind, start);
                    list.ListItems.Add(new Wpf.ListItem(BuildParagraph(p, outlineStyles, fnNums, enNums)));
                    break;
                }

                case Paragraph p when p.Style.IsThematicBreak:
                    listStack.Clear();
                    target.Add(BuildThematicBreak());
                    break;

                case Paragraph p:
                    listStack.Clear();
                    target.Add(BuildParagraph(p, outlineStyles, fnNums, enNums));
                    break;

                case Table t:
                    listStack.Clear();
                    if (t.WrapMode == TableWrapMode.Block)
                        target.Add(BuildTable(t, outlineStyles));
                    else
                        target.Add(BuildTableAnchor(t));   // 오버레이 모드 — 앵커만 추가
                    break;

                case ImageBlock image:
                    listStack.Clear();
                    target.Add(BuildImage(image));
                    break;

                case ShapeObject shape:
                    listStack.Clear();
                    target.Add(BuildShape(shape));
                    break;

                case OpaqueBlock opaque:
                    listStack.Clear();
                    target.Add(BuildOpaquePlaceholder(opaque));
                    break;

                case TocBlock toc:
                    listStack.Clear();
                    target.Add(BuildTocBlock(toc));
                    break;
            }
        }
    }

    /// <summary>TocBlock 을 시각적 BlockUIContainer 로 빌드한다. Tag = TocBlock 으로 라운드트립 가능.</summary>
    public static Wpf.BlockUIContainer BuildTocBlock(TocBlock toc)
    {
        var stack = new System.Windows.Controls.StackPanel { Margin = new Thickness(2) };

        // 제목
        var titleTb = new System.Windows.Controls.TextBlock
        {
            Text            = "목   차",
            FontWeight      = FontWeights.Bold,
            FontSize        = PtToDip(13),
            TextAlignment   = TextAlignment.Center,
            Padding         = new Thickness(0, 4, 0, 4),
        };
        stack.Children.Add(titleTb);

        // 구분선
        var sep = new System.Windows.Controls.Separator { Margin = new Thickness(0, 2, 0, 6) };
        stack.Children.Add(sep);

        if (toc.Entries.Count == 0)
        {
            stack.Children.Add(new System.Windows.Controls.TextBlock
            {
                Text       = "[목차 항목 없음 — '목차 새로고침'을 실행해 주세요]",
                Foreground = WpfMedia.Brushes.Gray,
                FontStyle  = FontStyles.Italic,
                Margin     = new Thickness(4, 2, 4, 2),
            });
        }
        else
        {
            foreach (var entry in toc.Entries)
            {
                var grid = new System.Windows.Controls.Grid
                {
                    Margin = new Thickness((entry.Level - 1) * 14.0, 1, 0, 1),
                };
                grid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition
                    { Width = new System.Windows.GridLength(1, System.Windows.GridUnitType.Star) });
                grid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition
                    { Width = System.Windows.GridLength.Auto });

                var entryTb = new System.Windows.Controls.TextBlock
                {
                    Text         = entry.Text,
                    FontWeight   = entry.Level == 1 ? FontWeights.SemiBold : FontWeights.Normal,
                    FontSize     = PtToDip(entry.Level == 1 ? 11 : 10),
                    TextTrimming = System.Windows.TextTrimming.CharacterEllipsis,
                };
                System.Windows.Controls.Grid.SetColumn(entryTb, 0);

                var pageTb = new System.Windows.Controls.TextBlock
                {
                    Text      = entry.PageNumber.HasValue
                                    ? entry.PageNumber.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)
                                    : "–",
                    TextAlignment = TextAlignment.Right,
                    MinWidth      = 28,
                    Margin        = new Thickness(8, 0, 0, 0),
                    FontSize      = PtToDip(10),
                    Foreground    = WpfMedia.Brushes.DimGray,
                };
                System.Windows.Controls.Grid.SetColumn(pageTb, 1);

                grid.Children.Add(entryTb);
                grid.Children.Add(pageTb);
                stack.Children.Add(grid);
            }
        }

        var border = new System.Windows.Controls.Border
        {
            BorderBrush     = new WpfMedia.SolidColorBrush(WpfMedia.Color.FromRgb(0xC0, 0xC0, 0xC0)),
            BorderThickness = new Thickness(1),
            CornerRadius    = new System.Windows.CornerRadius(3),
            Padding         = new Thickness(10),
            Margin          = new Thickness(0, 4, 0, 4),
            Background      = new WpfMedia.SolidColorBrush(WpfMedia.Color.FromArgb(15, 0, 0, 0)),
            Child           = stack,
        };

        return new Wpf.BlockUIContainer(border) { Tag = toc };
    }

    internal static Wpf.Table BuildTable(Table table, OutlineStyleSet? outlineStyles = null)
    {
        var wtable = new Wpf.Table { CellSpacing = 0 };

        ApplyTableLevelPropertiesToWpf(wtable, table);

        foreach (var col in table.Columns)
        {
            var width = col.WidthMm > 0
                ? new GridLength(MmToDip(col.WidthMm))
                : GridLength.Auto;
            wtable.Columns.Add(new Wpf.TableColumn { Width = width });
        }

        var rowGroup = new Wpf.TableRowGroup();
        wtable.RowGroups.Add(rowGroup);

        var headerBrush = new WpfMedia.SolidColorBrush(WpfMedia.Color.FromRgb(0xE8, 0xEA, 0xED));
        headerBrush.Freeze();

        foreach (var row in table.Rows)
        {
            var wrow = new Wpf.TableRow();
            if (row.IsHeader)
                wrow.Background = headerBrush;

            foreach (var cell in row.Cells)
            {
                var wcell = new Wpf.TableCell
                {
                    ColumnSpan = Math.Max(cell.ColumnSpan, 1),
                    RowSpan    = Math.Max(cell.RowSpan, 1),
                };

                ApplyCellPropertiesToWpf(wcell, cell, row.IsHeader, table);
                AppendBlocks(wcell.Blocks, cell.Blocks, outlineStyles);
                if (wcell.Blocks.Count == 0)
                    wcell.Blocks.Add(new Wpf.Paragraph(new Wpf.Run(string.Empty)));

                wrow.Cells.Add(wcell);
            }
            rowGroup.Rows.Add(wrow);
        }

        wtable.Tag = table;
        return wtable;
    }

    /// <summary>표 수준 속성(배경·바깥여백·외곽선·정렬)을 WPF Table 에 적용.</summary>
    internal static void ApplyTableLevelPropertiesToWpf(Wpf.Table wtable, Table table)
    {
        // 배경색
        if (!string.IsNullOrEmpty(table.BackgroundColor) &&
            TryParseColor(table.BackgroundColor) is { } bg)
            wtable.Background = new WpfMedia.SolidColorBrush(bg);
        else
            wtable.Background = null;

        // 바깥 여백
        wtable.Margin = new Thickness(
            table.OuterMarginLeftMm   > 0 ? MmToDip(table.OuterMarginLeftMm)   : 0,
            table.OuterMarginTopMm    > 0 ? MmToDip(table.OuterMarginTopMm)    : 0,
            table.OuterMarginRightMm  > 0 ? MmToDip(table.OuterMarginRightMm)  : 0,
            table.OuterMarginBottomMm > 0 ? MmToDip(table.OuterMarginBottomMm) : 0);

        // 표 외곽선
        if (table.BorderThicknessPt > 0)
        {
            var borderColor = TryParseColor(table.BorderColor)
                ?? WpfMedia.Color.FromRgb(0xC8, 0xC8, 0xC8);
            wtable.BorderBrush     = new WpfMedia.SolidColorBrush(borderColor);
            wtable.BorderThickness = new Thickness(PtToDip(table.BorderThicknessPt));
        }
        else
        {
            wtable.BorderBrush     = null;
            wtable.BorderThickness = new Thickness(0);
        }
    }

    internal static void ApplyCellPropertiesToWpf(
        Wpf.TableCell wcell,
        TableCell cell,
        bool isHeader,
        Table? tableDefaults = null)
    {
        var borderColor = TryParseColor(cell.BorderColor)
            ?? WpfMedia.Color.FromRgb(0xC8, 0xC8, 0xC8);
        double borderDip = cell.BorderThicknessPt > 0
            ? PtToDip(cell.BorderThicknessPt) : PtToDip(0.75);
        wcell.BorderBrush     = new WpfMedia.SolidColorBrush(borderColor);
        wcell.BorderThickness = new Thickness(borderDip);

        double defTop    = tableDefaults?.DefaultCellPaddingTopMm    > 0 ? tableDefaults.DefaultCellPaddingTopMm    : 1.0;
        double defBottom = tableDefaults?.DefaultCellPaddingBottomMm > 0 ? tableDefaults.DefaultCellPaddingBottomMm : 1.0;
        double defLeft   = tableDefaults?.DefaultCellPaddingLeftMm   > 0 ? tableDefaults.DefaultCellPaddingLeftMm   : 1.5;
        double defRight  = tableDefaults?.DefaultCellPaddingRightMm  > 0 ? tableDefaults.DefaultCellPaddingRightMm  : 1.5;

        double padTop   = MmToDip(cell.PaddingTopMm    > 0 ? cell.PaddingTopMm    : defTop);
        double padBottom= MmToDip(cell.PaddingBottomMm > 0 ? cell.PaddingBottomMm : defBottom);
        double padLeft  = MmToDip(cell.PaddingLeftMm   > 0 ? cell.PaddingLeftMm   : defLeft);
        double padRight = MmToDip(cell.PaddingRightMm  > 0 ? cell.PaddingRightMm  : defRight);
        wcell.Padding = new Thickness(padLeft, padTop, padRight, padBottom);

        if (!string.IsNullOrEmpty(cell.BackgroundColor) &&
            TryParseColor(cell.BackgroundColor) is { } bg)
            wcell.Background = new WpfMedia.SolidColorBrush(bg);
        else
            wcell.Background = null;

        wcell.FontWeight = isHeader ? FontWeights.SemiBold : FontWeights.Normal;
        ApplyCellTextAlign(wcell, cell.TextAlign);
    }

    private static void ApplyCellTextAlign(Wpf.TableCell wcell, CellTextAlign align)
    {
        var wpfAlign = align switch
        {
            CellTextAlign.Center  => TextAlignment.Center,
            CellTextAlign.Right   => TextAlignment.Right,
            CellTextAlign.Justify => TextAlignment.Justify,
            _                     => TextAlignment.Left,
        };
        foreach (var b in wcell.Blocks)
            if (b is Wpf.Paragraph p) p.TextAlignment = wpfAlign;
    }

    // ── 오버레이 표 지원 ─────────────────────────────────────────────────

    /// <summary><c>IsThematicBreak</c> 단락(hr)을 1px 수평선 BlockUIContainer 로 렌더링한다.</summary>
    private static Wpf.BlockUIContainer BuildThematicBreak()
    {
        var line = new System.Windows.Controls.Border
        {
            Height              = 1,
            Margin              = new Thickness(0, 4, 0, 4),
            Background          = new WpfMedia.SolidColorBrush(WpfMedia.Color.FromRgb(0xCC, 0xCC, 0xCC)),
            HorizontalAlignment = HorizontalAlignment.Stretch,
        };
        return new Wpf.BlockUIContainer(line);
    }

    /// <summary>오버레이(InFrontOfText/BehindText/Fixed) 모드 표를 위한 최소 앵커 단락을 반환한다.</summary>
    internal static Wpf.Paragraph BuildTableAnchor(Table table)
        => new Wpf.Paragraph
        {
            Tag        = table,
            Margin     = new Thickness(0),
            FontSize   = 0.1,
            Foreground = WpfMedia.Brushes.Transparent,
            Background = WpfMedia.Brushes.Transparent,
        };

    /// <summary>
    /// 오버레이 Canvas 에 배치할 표 시각 요소를 생성한다.
    /// System.Windows.Controls.Grid 기반으로 셀 테두리·배경·텍스트를 렌더링한다.
    /// </summary>
    internal static System.Windows.FrameworkElement? BuildOverlayTableControl(Table table)
    {
        if (table.Rows.Count == 0 || table.Columns.Count == 0) return null;

        var grid = new System.Windows.Controls.Grid();
        grid.Tag = table;

        // 컬럼 정의
        foreach (var col in table.Columns)
        {
            var w = col.WidthMm > 0
                ? new System.Windows.GridLength(MmToDip(col.WidthMm))
                : new System.Windows.GridLength(1, System.Windows.GridUnitType.Star);
            grid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = w });
        }

        // 행 정의
        foreach (var row in table.Rows)
        {
            var h = row.HeightMm > 0
                ? new System.Windows.GridLength(MmToDip(row.HeightMm))
                : System.Windows.GridLength.Auto;
            grid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = h });
        }

        // 셀
        var headerBg = new WpfMedia.SolidColorBrush(WpfMedia.Color.FromRgb(0xE8, 0xEA, 0xED));
        for (int r = 0; r < table.Rows.Count; r++)
        {
            var row = table.Rows[r];
            for (int c = 0; c < row.Cells.Count; c++)
            {
                var cell = row.Cells[c];
                var borderColor = TryParseColor(cell.BorderColor) ?? WpfMedia.Color.FromRgb(0xC8, 0xC8, 0xC8);
                double borderDip = cell.BorderThicknessPt > 0 ? PtToDip(cell.BorderThicknessPt) : PtToDip(0.75);

                WpfMedia.Brush? cellBg = null;
                if (row.IsHeader)
                    cellBg = headerBg;
                else if (!string.IsNullOrEmpty(cell.BackgroundColor) &&
                         TryParseColor(cell.BackgroundColor) is { } bg)
                    cellBg = new WpfMedia.SolidColorBrush(bg);

                double padL = MmToDip(cell.PaddingLeftMm   > 0 ? cell.PaddingLeftMm   : 1.5);
                double padT = MmToDip(cell.PaddingTopMm    > 0 ? cell.PaddingTopMm    : 1.0);
                double padR = MmToDip(cell.PaddingRightMm  > 0 ? cell.PaddingRightMm  : 1.5);
                double padB = MmToDip(cell.PaddingBottomMm > 0 ? cell.PaddingBottomMm : 1.0);

                // 셀 내용 텍스트 (첫 Paragraph 의 텍스트만 표시)
                string text = string.Concat(
                    cell.Blocks.OfType<Paragraph>().Take(1)
                               .SelectMany(p => p.Runs.Select(run => run.Text)));

                var textBlock = new System.Windows.Controls.TextBlock
                {
                    Text = text,
                    TextWrapping = System.Windows.TextWrapping.Wrap,
                    Padding = new Thickness(padL, padT, padR, padB),
                    FontWeight = row.IsHeader ? FontWeights.SemiBold : FontWeights.Normal,
                };

                var border = new System.Windows.Controls.Border
                {
                    BorderBrush     = new WpfMedia.SolidColorBrush(borderColor),
                    BorderThickness = new Thickness(borderDip),
                    Background      = cellBg,
                    Child           = textBlock,
                };

                System.Windows.Controls.Grid.SetRow(border, r);
                System.Windows.Controls.Grid.SetColumn(border, c);
                if (cell.ColumnSpan > 1) System.Windows.Controls.Grid.SetColumnSpan(border, cell.ColumnSpan);
                if (cell.RowSpan    > 1) System.Windows.Controls.Grid.SetRowSpan(border, cell.RowSpan);
                grid.Children.Add(border);
            }
        }

        // 표 배경
        if (!string.IsNullOrEmpty(table.BackgroundColor) &&
            TryParseColor(table.BackgroundColor) is { } tableBg)
            grid.Background = new WpfMedia.SolidColorBrush(tableBg);

        return grid;
    }

    private static WpfMedia.Color? TryParseColor(string? hex)
    {
        if (string.IsNullOrEmpty(hex)) return null;
        try { return (WpfMedia.Color)WpfMedia.ColorConverter.ConvertFromString(hex)!; }
        catch { return null; }
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
        // OnLoad + 명시 Dispose: EndInit 단계에서 BitmapImage 가 내부 캐시로 데이터를 복사하므로
        // 그 후엔 원본 MemoryStream 을 즉시 해제해도 안전하다. Freeze 전 시점이 마지막 정리 기회.
        var imgStream = new MemoryStream(image.Data, writable: false);
        bitmap.BeginInit();
        bitmap.CacheOption  = WpfMedia.Imaging.BitmapCacheOption.OnLoad;
        bitmap.StreamSource = imgStream;
        bitmap.EndInit();
        imgStream.Dispose();
        bitmap.Freeze();

        // Image.Tag 에 container 를 저장하지 말 것 — container.Child = image 와 함께 순환 참조가 되어
        // WPF undo 스냅샷의 XamlWriter.Save() 가 StackOverflowException 으로 폭주한다.
        // 우클릭 속성 라우팅은 LogicalTreeHelper 로 image 의 부모를 찾는 방식으로 처리한다.

        // BlockUIContainer 안에서 UIElement 의 가로 위치는 FrameworkElement.HorizontalAlignment 로만
        // 결정된다. BlockUIContainer.TextAlignment 는 텍스트 glyph 정렬이며 UIElement 에는 무효.
        // 명시적 Width 가 있는 UIElement 의 기본 HorizontalAlignment(Stretch) 는 중앙 배치처럼
        // 동작하므로, 의도한 정렬이 Left 여도 가운데에 놓이는 버그가 생긴다.
        var imgHA = image.HAlign switch
        {
            ImageHAlign.Center => HorizontalAlignment.Center,
            ImageHAlign.Right  => HorizontalAlignment.Right,
            _                  => HorizontalAlignment.Left,
        };

        var control = new System.Windows.Controls.Image
        {
            Source              = bitmap,
            Stretch             = WpfMedia.Stretch.Uniform,
            HorizontalAlignment = imgHA,
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

            // Border 가 BlockUIContainer 의 직접 자식일 때 HorizontalAlignment 가 없으면
            // Stretch 로 처리되어 Image 의 정렬이 무시된다 — imgHA 를 함께 지정.
            visual = new System.Windows.Controls.Border
            {
                Child               = control,
                BorderBrush         = borderBrush,
                BorderThickness     = new Thickness(PtToDip(image.BorderThicknessPt)),
                HorizontalAlignment = imgHA,
            };
        }

        // 그림 제목 래퍼
        visual = WrapImageWithTitle(visual, image, imgHA);

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

        // Paragraph 는 Tag 로 ImageBlock 을 보존해 라운드트립 가능. Floater 를 inlines 에 추가하고
        // 그 다음에 anchor Run(비줄바꿈 공백  ) 을 추가한다.
        // - WrapRight(우측 정렬 Floater) 는 라인에 실제 글리프가 없으면 WPF TextFormatter 가
        //   "유효한 라인 없음"으로 처리해 Floater 위치를 결정하지 못하고 그림이 사라진다.
        //   빈 Run("") 은 글리프를 생성하지 않으므로  (보이지 않는 폭 있는 문자) 를 사용한다.
        // - Foreground/Background Transparent 로   과 선택 하이라이트 모두 시각적으로 억제.
        // - LineHeight = 0.1 로 라인 자체를 거의 0 높이로 만들어 본문 흐름 영향 최소화.
        var paragraph = new Wpf.Paragraph
        {
            Tag        = image,
            Margin     = new Thickness(0),
            LineHeight = 0.1,
            Foreground = WpfMedia.Brushes.Transparent,
            Background = WpfMedia.Brushes.Transparent,
        };
        paragraph.Inlines.Add(floater);
        paragraph.Inlines.Add(new Wpf.Run(" ")); // non-breaking space: anchors line for right Floater
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
        // OnLoad + 명시 Dispose: EndInit 단계에서 BitmapImage 가 내부 캐시로 데이터를 복사하므로
        // 그 후엔 원본 MemoryStream 을 즉시 해제해도 안전하다. Freeze 전 시점이 마지막 정리 기회.
        var imgStream = new MemoryStream(image.Data, writable: false);
        bitmap.BeginInit();
        bitmap.CacheOption  = WpfMedia.Imaging.BitmapCacheOption.OnLoad;
        bitmap.StreamSource = imgStream;
        bitmap.EndInit();
        imgStream.Dispose();
        bitmap.Freeze();

        var control = new System.Windows.Controls.Image
        {
            Source  = bitmap,
            Stretch = WpfMedia.Stretch.Fill,
        };
        if (image.WidthMm > 0)  control.Width  = MmToDip(image.WidthMm);
        if (image.HeightMm > 0) control.Height = MmToDip(image.HeightMm);
        if (!string.IsNullOrEmpty(image.Description)) control.ToolTip = image.Description;

        UIElement overlayVisual = control;
        if (!string.IsNullOrEmpty(image.BorderColor) && image.BorderThicknessPt > 0)
        {
            WpfMedia.Brush borderBrush;
            try { borderBrush = new WpfMedia.SolidColorBrush(
                    (WpfMedia.Color)WpfMedia.ColorConverter.ConvertFromString(image.BorderColor)!); }
            catch { borderBrush = WpfMedia.Brushes.Black; }

            overlayVisual = new System.Windows.Controls.Border
            {
                Child           = control,
                BorderBrush     = borderBrush,
                BorderThickness = new Thickness(PtToDip(image.BorderThicknessPt)),
            };
        }

        var withTitle = WrapImageWithTitle(overlayVisual, image, HorizontalAlignment.Left);
        return withTitle as System.Windows.FrameworkElement ?? control;
    }

    /// <summary>
    /// 그림 제목(캡션) 표시가 켜져 있으면 image 시각 요소를 Grid 로 감싸 제목을 함께 배치한다.
    /// 위치(Above/Below/OverlayTop/Middle/Bottom) + 가로 정렬 + X/Y 오프셋(mm)을 적용.
    /// </summary>
    private static UIElement WrapImageWithTitle(UIElement imageVisual, ImageBlock image, HorizontalAlignment imgHA)
    {
        if (!image.ShowTitle || string.IsNullOrWhiteSpace(image.Title)) return imageVisual;

        var s = image.TitleStyle;

        // 제목 텍스트블록 — RunStyle 기반 (CharFormatWindow 와 모델 공유).
        WpfMedia.Brush titleFg = s.Foreground is { } fg
            ? new WpfMedia.SolidColorBrush(WpfMedia.Color.FromArgb(fg.A, fg.R, fg.G, fg.B))
            : WpfMedia.Brushes.Black;
        WpfMedia.Brush? titleBg = s.Background is { } bg
            ? new WpfMedia.SolidColorBrush(WpfMedia.Color.FromArgb(bg.A, bg.R, bg.G, bg.B))
            : null;

        TextAlignment ta = image.TitleHAlign switch
        {
            ImageHAlign.Left   => TextAlignment.Left,
            ImageHAlign.Right  => TextAlignment.Right,
            _                  => TextAlignment.Center,
        };
        var tb = new System.Windows.Controls.TextBlock
        {
            Text          = image.Title,
            Foreground    = titleFg,
            FontSize      = PtToDip(s.FontSizePt > 0 ? s.FontSizePt : 10),
            FontWeight    = s.Bold   ? FontWeights.Bold   : FontWeights.Normal,
            FontStyle     = s.Italic ? FontStyles.Italic  : FontStyles.Normal,
            TextWrapping  = TextWrapping.Wrap,
            TextAlignment = ta,
        };
        if (titleBg is not null) tb.Background = titleBg;
        if (!string.IsNullOrEmpty(s.FontFamily))
            tb.FontFamily = new WpfMedia.FontFamily(s.FontFamily);

        // 텍스트 장식 (밑줄/취소선/위줄)
        if (s.Underline || s.Strikethrough || s.Overline)
        {
            var decos = new System.Windows.TextDecorationCollection();
            if (s.Underline)     foreach (var d in System.Windows.TextDecorations.Underline)    decos.Add(d);
            if (s.Strikethrough) foreach (var d in System.Windows.TextDecorations.Strikethrough) decos.Add(d);
            if (s.Overline)      foreach (var d in System.Windows.TextDecorations.OverLine)     decos.Add(d);
            tb.TextDecorations = decos;
        }

        // 오프셋 — TranslateTransform 으로 적용해 정렬·레이아웃에 영향 없이 미세 이동.
        if (Math.Abs(image.TitleOffsetXMm) > 0.001 || Math.Abs(image.TitleOffsetYMm) > 0.001)
        {
            tb.RenderTransform = new WpfMedia.TranslateTransform(
                MmToDip(image.TitleOffsetXMm), MmToDip(image.TitleOffsetYMm));
        }

        var grid = new System.Windows.Controls.Grid { HorizontalAlignment = imgHA };

        bool isOverlay = image.TitlePosition is ImageTitlePosition.OverlayTop
                                              or ImageTitlePosition.OverlayMiddle
                                              or ImageTitlePosition.OverlayBottom;
        if (isOverlay)
        {
            // 같은 셀에 그림과 제목이 겹침. VerticalAlignment 로 위/가운데/아래 결정.
            grid.Children.Add(imageVisual);
            tb.VerticalAlignment = image.TitlePosition switch
            {
                ImageTitlePosition.OverlayTop    => VerticalAlignment.Top,
                ImageTitlePosition.OverlayBottom => VerticalAlignment.Bottom,
                _                                => VerticalAlignment.Center,
            };
            tb.HorizontalAlignment = HorizontalAlignment.Stretch;
            grid.Children.Add(tb);
        }
        else
        {
            grid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = System.Windows.GridLength.Auto });
            grid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = System.Windows.GridLength.Auto });
            int titleRow = image.TitlePosition == ImageTitlePosition.Above ? 0 : 1;
            int imageRow = image.TitlePosition == ImageTitlePosition.Above ? 1 : 0;
            System.Windows.Controls.Grid.SetRow(tb, titleRow);
            System.Windows.Controls.Grid.SetRow((FrameworkElement)imageVisual, imageRow);
            tb.HorizontalAlignment = HorizontalAlignment.Stretch;
            grid.Children.Add(imageVisual);
            grid.Children.Add(tb);
        }
        return grid;
    }

    // ── 도형 렌더링 ─────────────────────────────────────────────────────────

    /// <summary>
    /// ShapeObject 를 WPF Block 으로 빌드한다.
    /// ImageBlock 과 동일한 5모드 배치 체계를 사용한다.
    /// </summary>
    internal static Wpf.Block BuildShape(ShapeObject shape)
    {
        // ── 오버레이 모드 ─────────────────────────────────────────────────
        if (shape.WrapMode is ImageWrapMode.InFrontOfText or ImageWrapMode.BehindText)
        {
            return new Wpf.Paragraph
            {
                Tag        = shape,
                Margin     = new Thickness(0),
                FontSize   = 0.1,
                Foreground = WpfMedia.Brushes.Transparent,
                Background = WpfMedia.Brushes.Transparent,
            };
        }

        double wDip = MmToDip(shape.WidthMm);
        double hDip = MmToDip(shape.HeightMm);
        var visual  = BuildShapeVisual(shape, wDip, hDip);

        var marginTopDip    = MmToDip(shape.MarginTopMm);
        var marginBottomDip = MmToDip(shape.MarginBottomMm);

        var imgHA = shape.HAlign switch
        {
            ImageHAlign.Center => HorizontalAlignment.Center,
            ImageHAlign.Right  => HorizontalAlignment.Right,
            _                  => HorizontalAlignment.Left,
        };
        visual.HorizontalAlignment = imgHA;

        // ── Inline ────────────────────────────────────────────────────────
        if (shape.WrapMode == ImageWrapMode.Inline)
        {
            return new Wpf.BlockUIContainer(visual)
            {
                Tag           = shape,
                Margin        = new Thickness(0, marginTopDip, 0, marginBottomDip),
                TextAlignment = shape.HAlign switch
                {
                    ImageHAlign.Center => TextAlignment.Center,
                    ImageHAlign.Right  => TextAlignment.Right,
                    _                  => TextAlignment.Left,
                },
            };
        }

        // ── WrapLeft / WrapRight ──────────────────────────────────────────
        var floater = new Wpf.Floater
        {
            HorizontalAlignment = shape.WrapMode == ImageWrapMode.WrapRight
                ? HorizontalAlignment.Right
                : HorizontalAlignment.Left,
            Padding = new Thickness(0),
            Margin  = new Thickness(
                shape.WrapMode == ImageWrapMode.WrapRight ? PtToDip(8) : 0,
                marginTopDip,
                shape.WrapMode == ImageWrapMode.WrapLeft  ? PtToDip(8) : 0,
                marginBottomDip),
        };
        if (shape.WidthMm > 0) floater.Width = wDip;
        floater.Blocks.Add(new Wpf.BlockUIContainer(visual));

        var paragraph = new Wpf.Paragraph
        {
            Tag        = shape,
            Margin     = new Thickness(0),
            LineHeight = 0.1,
            Foreground = WpfMedia.Brushes.Transparent,
            Background = WpfMedia.Brushes.Transparent,
        };
        paragraph.Inlines.Add(floater);
        paragraph.Inlines.Add(new Wpf.Run(" "));
        return paragraph;
    }

    /// <summary>
    /// ShapeObject 로부터 캔버스 오버레이용 프레임워크 요소를 생성한다.
    /// InFrontOfText / BehindText 모드 도형을 오버레이 캔버스에 배치할 때 사용.
    /// 위치(OverlayXMm/OverlayYMm)는 호출측에서 Canvas.Left/Top 으로 설정.
    /// </summary>
    public static FrameworkElement BuildOverlayShapeControl(ShapeObject shape)
    {
        double wDip = MmToDip(shape.WidthMm);
        double hDip = MmToDip(shape.HeightMm);
        return BuildShapeVisual(shape, wDip, hDip);
    }

    private static System.Windows.Controls.Canvas BuildShapeVisual(ShapeObject shape, double wDip, double hDip)
    {
        var canvas = new System.Windows.Controls.Canvas
        {
            Width  = wDip,
            Height = hDip,
        };

        var geometry = BuildShapeGeometry(shape, wDip, hDip);

        // 채우기 브러시
        WpfMedia.Brush fillBrush = WpfMedia.Brushes.Transparent;
        if (!string.IsNullOrEmpty(shape.FillColor))
        {
            try
            {
                var c = (WpfMedia.Color)WpfMedia.ColorConverter.ConvertFromString(shape.FillColor)!;
                var alpha = (byte)Math.Clamp(shape.FillOpacity * 255, 0, 255);
                fillBrush = new WpfMedia.SolidColorBrush(WpfMedia.Color.FromArgb(alpha, c.R, c.G, c.B));
            }
            catch { fillBrush = WpfMedia.Brushes.LightSteelBlue; }
        }

        // 선 브러시
        WpfMedia.Brush strokeBrush = WpfMedia.Brushes.Black;
        if (!string.IsNullOrEmpty(shape.StrokeColor))
        {
            try
            {
                var c = (WpfMedia.Color)WpfMedia.ColorConverter.ConvertFromString(shape.StrokeColor)!;
                strokeBrush = new WpfMedia.SolidColorBrush(c);
            }
            catch { }
        }

        double strokeDip = shape.StrokeThicknessPt > 0 ? PtToDip(shape.StrokeThicknessPt) : 0;

        var path = new WpfShapes.Path
        {
            Data            = geometry,
            Fill            = fillBrush,
            Stroke          = strokeBrush,
            StrokeThickness = strokeDip,
            StrokeDashArray = BuildDashArray(shape.StrokeDash, strokeDip),
            Stretch         = WpfMedia.Stretch.None,
        };

        canvas.Children.Add(path);

        // 끝모양 (선 계열 — 열린 선에만); path 추가 후 그 위에 그림
        if (shape.Kind is ShapeKind.Line or ShapeKind.Polyline or ShapeKind.Spline)
        {
            var ptsDip = GetPointsDip(shape.Points, wDip, hDip);
            if (ptsDip.Count < 2 && shape.Kind == ShapeKind.Line)
            {
                ptsDip = new List<Point> { new(0, hDip / 2), new(wDip, hDip / 2) };
            }
            AddArrowHeads(canvas, shape.StartArrow, shape.EndArrow, shape.EndShapeSizeMm, ptsDip, strokeBrush, strokeDip);
        }

        // 레이블
        if (!string.IsNullOrWhiteSpace(shape.LabelText))
        {
            var label = BuildShapeLabel(shape, wDip, hDip);
            canvas.Children.Add(label);
        }

        // 회전 — canvas 전체에 적용해 path·끝모양·레이블이 함께 회전한다.
        if (Math.Abs(shape.RotationAngleDeg) > 0.01)
        {
            canvas.RenderTransformOrigin = new Point(0.5, 0.5);
            canvas.RenderTransform = new WpfMedia.RotateTransform(shape.RotationAngleDeg);
        }

        return canvas;
    }

    private static WpfMedia.DoubleCollection? BuildDashArray(StrokeDash dash, double strokeDip)
    {
        if (strokeDip <= 0) return null;
        return dash switch
        {
            StrokeDash.Dashed  => new WpfMedia.DoubleCollection { 6, 3 },
            StrokeDash.Dotted  => new WpfMedia.DoubleCollection { 1, 3 },
            StrokeDash.DashDot => new WpfMedia.DoubleCollection { 6, 2, 1, 2 },
            _                  => null,
        };
    }

    private static void AddArrowHeads(
        System.Windows.Controls.Canvas canvas,
        ShapeArrow start, ShapeArrow end,
        double endShapeSizeMm,
        List<Point> ptsDip,
        WpfMedia.Brush strokeBrush,
        double strokeDip)
    {
        if (start == ShapeArrow.None && end == ShapeArrow.None) return;
        if (ptsDip.Count < 2) return;

        // 사용자가 mm 로 명시한 크기가 있으면 그 값 사용, 아니면 선 두께에 비례 (최소 2.5 mm).
        double arrowLen  = endShapeSizeMm > 0
            ? MmToDip(endShapeSizeMm)
            : Math.Max(strokeDip * 5.0, MmToDip(2.5));
        double arrowHalf = arrowLen * 0.38;

        if (start != ShapeArrow.None)
            AddOneArrowHead(canvas, start, ptsDip[0], ptsDip[1], arrowLen, arrowHalf, strokeBrush, strokeDip);

        if (end != ShapeArrow.None)
        {
            int n = ptsDip.Count;
            AddOneArrowHead(canvas, end, ptsDip[n - 1], ptsDip[n - 2], arrowLen, arrowHalf, strokeBrush, strokeDip);
        }
    }

    private static void AddOneArrowHead(
        System.Windows.Controls.Canvas canvas,
        ShapeArrow kind,
        Point tip, Point from,
        double arrowLen, double arrowHalf,
        WpfMedia.Brush brush,
        double strokeDip)
    {
        double dx  = tip.X - from.X;
        double dy  = tip.Y - from.Y;
        double len = Math.Sqrt(dx * dx + dy * dy);
        if (len < 0.001) return;

        double ux = dx / len;   // 화살 방향 단위 벡터
        double uy = dy / len;
        double px = -uy;        // 수직 방향
        double py =  ux;

        // 삼각형 밑변 중점
        double bcx = tip.X - ux * arrowLen;
        double bcy = tip.Y - uy * arrowLen;

        switch (kind)
        {
            case ShapeArrow.Open:
            {
                var left  = new Point(bcx + px * arrowHalf, bcy + py * arrowHalf);
                var right = new Point(bcx - px * arrowHalf, bcy - py * arrowHalf);
                var pg    = new WpfMedia.PathGeometry();
                var fig   = new WpfMedia.PathFigure { StartPoint = left };
                fig.Segments.Add(new WpfMedia.LineSegment(tip,   true));
                fig.Segments.Add(new WpfMedia.LineSegment(right, true));
                pg.Figures.Add(fig);
                canvas.Children.Add(new WpfShapes.Path
                {
                    Data            = pg,
                    Stroke          = brush,
                    StrokeThickness = strokeDip > 0 ? strokeDip : 1.0,
                    Fill            = WpfMedia.Brushes.Transparent,
                    StrokeLineJoin  = WpfMedia.PenLineJoin.Miter,
                });
                break;
            }
            case ShapeArrow.Filled:
            {
                var left  = new Point(bcx + px * arrowHalf, bcy + py * arrowHalf);
                var right = new Point(bcx - px * arrowHalf, bcy - py * arrowHalf);
                var pg    = new WpfMedia.PathGeometry();
                var fig   = new WpfMedia.PathFigure { StartPoint = tip, IsClosed = true, IsFilled = true };
                fig.Segments.Add(new WpfMedia.LineSegment(left,  true));
                fig.Segments.Add(new WpfMedia.LineSegment(right, true));
                pg.Figures.Add(fig);
                canvas.Children.Add(new WpfShapes.Path
                {
                    Data            = pg,
                    Fill            = brush,
                    Stroke          = brush,
                    StrokeThickness = 0,
                });
                break;
            }
            case ShapeArrow.Diamond:
            {
                var midLeft  = new Point(bcx + px * arrowHalf,    bcy + py * arrowHalf);
                var midRight = new Point(bcx - px * arrowHalf,    bcy - py * arrowHalf);
                var back     = new Point(tip.X - ux * arrowLen * 2, tip.Y - uy * arrowLen * 2);
                var pg       = new WpfMedia.PathGeometry();
                var fig      = new WpfMedia.PathFigure { StartPoint = tip, IsClosed = true, IsFilled = true };
                fig.Segments.Add(new WpfMedia.LineSegment(midLeft,  true));
                fig.Segments.Add(new WpfMedia.LineSegment(back,     true));
                fig.Segments.Add(new WpfMedia.LineSegment(midRight, true));
                pg.Figures.Add(fig);
                canvas.Children.Add(new WpfShapes.Path
                {
                    Data            = pg,
                    Fill            = brush,
                    Stroke          = brush,
                    StrokeThickness = 0,
                });
                break;
            }
            case ShapeArrow.Circle:
            {
                double r   = arrowHalf;
                double cx  = tip.X - ux * r;
                double cy  = tip.Y - uy * r;
                var ellipse = new WpfShapes.Ellipse
                {
                    Width           = r * 2,
                    Height          = r * 2,
                    Fill            = brush,
                    Stroke          = brush,
                    StrokeThickness = 0,
                };
                System.Windows.Controls.Canvas.SetLeft(ellipse, cx - r);
                System.Windows.Controls.Canvas.SetTop (ellipse, cy - r);
                canvas.Children.Add(ellipse);
                break;
            }
        }
    }

    private static System.Windows.Controls.TextBlock BuildShapeLabel(ShapeObject shape, double wDip, double hDip)
    {
        WpfMedia.Brush labelBrush = WpfMedia.Brushes.Black;
        if (!string.IsNullOrEmpty(shape.LabelColor))
        {
            try
            {
                var c = (WpfMedia.Color)WpfMedia.ColorConverter.ConvertFromString(shape.LabelColor)!;
                labelBrush = new WpfMedia.SolidColorBrush(c);
            }
            catch { }
        }
        else if (!string.IsNullOrEmpty(shape.FillColor))
        {
            // 채우기가 어두우면 흰색 레이블, 밝으면 검정
            try
            {
                var c = (WpfMedia.Color)WpfMedia.ColorConverter.ConvertFromString(shape.FillColor)!;
                double lum = 0.299 * c.R + 0.587 * c.G + 0.114 * c.B;
                labelBrush = lum < 128 ? WpfMedia.Brushes.White : WpfMedia.Brushes.Black;
            }
            catch { }
        }

        WpfMedia.Brush? labelBgBrush = null;
        if (!string.IsNullOrEmpty(shape.LabelBackgroundColor))
        {
            try
            {
                var c = (WpfMedia.Color)WpfMedia.ColorConverter.ConvertFromString(shape.LabelBackgroundColor)!;
                labelBgBrush = new WpfMedia.SolidColorBrush(c);
            }
            catch { }
        }

        TextAlignment textAlign = shape.LabelHAlign switch
        {
            ShapeLabelHAlign.Left   => TextAlignment.Left,
            ShapeLabelHAlign.Right  => TextAlignment.Right,
            _                       => TextAlignment.Center,
        };

        double fontDip   = PtToDip(shape.LabelFontSizePt > 0 ? shape.LabelFontSizePt : 10);
        double lineDip   = fontDip * 1.4;

        var tb = new System.Windows.Controls.TextBlock
        {
            Text                = shape.LabelText,
            Foreground          = labelBrush,
            FontSize            = fontDip,
            FontWeight          = shape.LabelBold   ? FontWeights.Bold   : FontWeights.Normal,
            FontStyle           = shape.LabelItalic ? FontStyles.Italic  : FontStyles.Normal,
            TextWrapping        = TextWrapping.Wrap,
            TextAlignment       = textAlign,
            Width               = wDip,
        };
        if (labelBgBrush is not null) tb.Background = labelBgBrush;
        if (!string.IsNullOrEmpty(shape.LabelFontFamily))
            tb.FontFamily = new WpfMedia.FontFamily(shape.LabelFontFamily);

        // 세로 정렬 — Top/Middle/Bottom 에 따라 Y 위치 계산.
        double topDip = shape.LabelVAlign switch
        {
            ShapeLabelVAlign.Top    => 0,
            ShapeLabelVAlign.Bottom => Math.Max(0, hDip - lineDip),
            _                       => Math.Max(0, (hDip - lineDip) / 2.0),
        };

        // 사용자 지정 오프셋 (mm) 추가 — 정렬 위치에서 추가로 이동.
        double leftDip = MmToDip(shape.LabelOffsetXMm);
        topDip       += MmToDip(shape.LabelOffsetYMm);

        System.Windows.Controls.Canvas.SetLeft(tb, leftDip);
        System.Windows.Controls.Canvas.SetTop (tb, topDip);
        return tb;
    }

    private static WpfMedia.Geometry BuildShapeGeometry(ShapeObject shape, double wDip, double hDip)
    {
        switch (shape.Kind)
        {
            case ShapeKind.Rectangle:
                return new WpfMedia.RectangleGeometry(new Rect(0, 0, wDip, hDip));

            case ShapeKind.RoundedRect:
            {
                double rx = MmToDip(shape.CornerRadiusMm);
                rx = Math.Clamp(rx, 0, Math.Min(wDip, hDip) / 2.0);
                return new WpfMedia.RectangleGeometry(new Rect(0, 0, wDip, hDip), rx, rx);
            }

            case ShapeKind.Ellipse:
                return new WpfMedia.EllipseGeometry(new Rect(0, 0, wDip, hDip));

            case ShapeKind.Line:
            {
                var pts = GetPointsDip(shape.Points, wDip, hDip);
                var p0  = pts.Count >= 1 ? pts[0] : new Point(0, hDip / 2);
                var p1  = pts.Count >= 2 ? pts[1] : new Point(wDip, hDip / 2);
                return new WpfMedia.LineGeometry(p0, p1);
            }

            case ShapeKind.Polyline:
            {
                var pts = GetPointsDip(shape.Points, wDip, hDip);
                if (pts.Count < 2) goto default;
                var pg = new WpfMedia.PathGeometry();
                var fig = new WpfMedia.PathFigure { StartPoint = pts[0] };
                for (int i = 1; i < pts.Count; i++)
                    fig.Segments.Add(new WpfMedia.LineSegment(pts[i], true));
                pg.Figures.Add(fig);
                return pg;
            }

            case ShapeKind.Spline:
            {
                var corePts = shape.Points;
                var pts     = GetPointsDip(corePts, wDip, hDip);
                if (pts.Count < 2) goto default;
                var pg  = new WpfMedia.PathGeometry();
                var fig = new WpfMedia.PathFigure { StartPoint = pts[0] };

                if (pts.Count == 2)
                {
                    fig.Segments.Add(new WpfMedia.LineSegment(pts[1], true));
                }
                else
                {
                    for (int i = 0; i < pts.Count - 1; i++)
                    {
                        var (c1, c2) = GetBezierControlsDip(corePts, pts, i, closed: false);
                        fig.Segments.Add(new WpfMedia.BezierSegment(c1, c2, pts[i + 1], true));
                    }
                }

                pg.Figures.Add(fig);
                return pg;
            }

            case ShapeKind.Polygon:
            {
                var pts = GetPointsDip(shape.Points, wDip, hDip);
                if (pts.Count < 3) goto default;
                var pg  = new WpfMedia.PathGeometry();
                var fig = new WpfMedia.PathFigure { StartPoint = pts[0], IsClosed = true, IsFilled = true };
                for (int i = 1; i < pts.Count; i++)
                    fig.Segments.Add(new WpfMedia.LineSegment(pts[i], true));
                pg.Figures.Add(fig);
                return pg;
            }

            case ShapeKind.ClosedSpline:
            {
                var corePts = shape.Points;
                var pts     = GetPointsDip(corePts, wDip, hDip);
                if (pts.Count < 3) goto default;
                int n   = pts.Count;
                var pg  = new WpfMedia.PathGeometry();
                // 닫힌 스플라인: 마지막 구간도 wrap-around 해 매끄럽게 처음 점으로 연결.
                var fig = new WpfMedia.PathFigure { StartPoint = pts[0], IsFilled = true };
                for (int i = 0; i < n; i++)
                {
                    var (c1, c2) = GetBezierControlsDip(corePts, pts, i, closed: true);
                    fig.Segments.Add(new WpfMedia.BezierSegment(c1, c2, pts[(i + 1) % n], true));
                }
                pg.Figures.Add(fig);
                return pg;
            }

            case ShapeKind.Triangle:
            {
                var pts = GetPointsDip(shape.Points, wDip, hDip);
                if (pts.Count < 3)
                    pts = new List<Point> { new(wDip / 2, 0), new(0, hDip), new(wDip, hDip) };
                var pg  = new WpfMedia.PathGeometry();
                var fig = new WpfMedia.PathFigure { StartPoint = pts[0], IsClosed = true };
                for (int i = 1; i < pts.Count; i++)
                    fig.Segments.Add(new WpfMedia.LineSegment(pts[i], true));
                pg.Figures.Add(fig);
                return pg;
            }

            case ShapeKind.RegularPolygon:
            {
                var pts = ComputeRegularPolygonPoints(
                    Math.Max(3, shape.SideCount), wDip / 2, hDip / 2, wDip / 2, hDip / 2);
                var pg  = new WpfMedia.PathGeometry();
                var fig = new WpfMedia.PathFigure { StartPoint = pts[0], IsClosed = true };
                for (int i = 1; i < pts.Count; i++)
                    fig.Segments.Add(new WpfMedia.LineSegment(pts[i], true));
                pg.Figures.Add(fig);
                return pg;
            }

            case ShapeKind.Star:
            {
                var pts = ComputeStarPoints(
                    Math.Max(3, shape.SideCount),
                    Math.Clamp(shape.InnerRadiusRatio, 0.1, 0.9),
                    wDip / 2, hDip / 2, wDip / 2, hDip / 2);
                var pg  = new WpfMedia.PathGeometry();
                var fig = new WpfMedia.PathFigure { StartPoint = pts[0], IsClosed = true, IsFilled = true };
                for (int i = 1; i < pts.Count; i++)
                    fig.Segments.Add(new WpfMedia.LineSegment(pts[i], true));
                pg.Figures.Add(fig);
                pg.FillRule = WpfMedia.FillRule.EvenOdd;
                return pg;
            }

            default:
                return new WpfMedia.RectangleGeometry(new Rect(0, 0, wDip, hDip));
        }
    }

    private static List<Point> GetPointsDip(IList<ShapePoint> points, double wDip, double hDip)
    {
        var result = new List<Point>(points.Count);
        foreach (var pt in points)
            result.Add(new Point(MmToDip(pt.X), MmToDip(pt.Y)));
        return result;
    }

    private static List<Point> ComputeRegularPolygonPoints(
        int sides, double rx, double ry, double cx, double cy)
    {
        var pts = new List<Point>(sides);
        double startAngle = -Math.PI / 2; // 정상(12시)에서 시작
        for (int i = 0; i < sides; i++)
        {
            double angle = startAngle + 2 * Math.PI * i / sides;
            pts.Add(new Point(cx + rx * Math.Cos(angle), cy + ry * Math.Sin(angle)));
        }
        return pts;
    }

    /// <summary>
    /// 세그먼트 [i → i+1] 의 cubic Bezier 제어점 (c1, c2) 를 DIP 단위로 반환.
    /// ShapePoint 에 명시적 OutCtrl/InCtrl 이 모두 설정돼 있으면 그것을 사용하고,
    /// 그렇지 않으면 Catmull-Rom 자동 계산으로 폴백한다.
    /// </summary>
    private static (Point c1, Point c2) GetBezierControlsDip(
        IList<PolyDonky.Core.ShapePoint> corePts, IList<Point> dipPts, int i, bool closed)
    {
        int n    = corePts.Count;
        int next = closed ? (i + 1) % n : i + 1;

        var from = corePts[i];
        var to   = corePts[next];

        if (from.OutCtrlX.HasValue && from.OutCtrlY.HasValue
            && to.InCtrlX.HasValue   && to.InCtrlY.HasValue)
        {
            return (new Point(MmToDip(from.OutCtrlX.Value), MmToDip(from.OutCtrlY.Value)),
                    new Point(MmToDip(to.InCtrlX.Value),    MmToDip(to.InCtrlY.Value)));
        }

        // Catmull-Rom: 좌우 이웃점을 이용해 1/6 접선 오프셋 계산.
        var p0 = closed ? dipPts[(i - 1 + n) % n] : (i == 0 ? dipPts[0]     : dipPts[i - 1]);
        var p1 = dipPts[i];
        var p2 = dipPts[next];
        var p3 = closed ? dipPts[(i + 2) % n]     : (i + 2 < n ? dipPts[i + 2] : dipPts[Math.Min(i + 1, n - 1)]);

        var c1 = new Point(p1.X + (p2.X - p0.X) / 6.0, p1.Y + (p2.Y - p0.Y) / 6.0);
        var c2 = new Point(p2.X - (p3.X - p1.X) / 6.0, p2.Y - (p3.Y - p1.Y) / 6.0);
        return (c1, c2);
    }

    private static List<Point> ComputeStarPoints(
        int points, double innerRatio, double rx, double ry, double cx, double cy)
    {
        var pts  = new List<Point>(points * 2);
        double startAngle = -Math.PI / 2;
        for (int i = 0; i < points * 2; i++)
        {
            double angle = startAngle + Math.PI * i / points;
            double r     = i % 2 == 0 ? 1.0 : innerRatio;
            pts.Add(new Point(cx + rx * r * Math.Cos(angle), cy + ry * r * Math.Sin(angle)));
        }
        return pts;
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

    internal static Wpf.Paragraph BuildParagraph(Paragraph p, OutlineStyleSet? outlineStyles = null,
        IReadOnlyDictionary<string, int>? fnNums = null, IReadOnlyDictionary<string, int>? enNums = null)
    {
        var wpfPara = new Wpf.Paragraph();
        ApplyParagraphStyle(wpfPara, p.Style, outlineStyles);
        foreach (var run in p.Runs)
        {
            wpfPara.Inlines.Add(BuildInline(run, fnNums, enNums));
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

        if (style.ForcePageBreakBefore)
            wpfPara.BreakPageBefore = true;

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
            ApplyQuoteLevelStyle(wpfPara, style.QuoteLevel);
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

        ApplyCodeBlockStyle(wpfPara, style.CodeLanguage);
        ApplyQuoteLevelStyle(wpfPara, style.QuoteLevel);
    }

    /// <summary>
    /// 블록쿼트 들여쓰기 + 왼쪽 테두리. 레벨당 5mm 들여쓰고 회색 3px 왼쪽 테두리를 그린다.
    /// heading 분기와 body 분기 양쪽에서 호출되므로 별도 메서드로 분리.
    /// </summary>
    private static void ApplyQuoteLevelStyle(Wpf.Paragraph wpfPara, int quoteLevel)
    {
        if (quoteLevel <= 0) return;
        double indentDip = MmToDip(5.0 * quoteLevel);
        var m = wpfPara.Margin;
        wpfPara.Margin          = new Thickness(indentDip + m.Left, m.Top > 0 ? m.Top : 2, m.Right, m.Bottom > 0 ? m.Bottom : 2);
        wpfPara.Padding         = new Thickness(MmToDip(3.0), 0, 0, 0);
        wpfPara.BorderBrush     = new WpfMedia.SolidColorBrush(WpfMedia.Color.FromRgb(0xC0, 0xC0, 0xC0));
        wpfPara.BorderThickness = new Thickness(3, 0, 0, 0);
    }

    /// <summary>
    /// CodeLanguage != null(= pre/code 블록)이면 회색 배경·테두리·모노스페이스를 적용.
    /// null 이면 일반 단락 — 아무것도 하지 않는다.
    /// </summary>
    private static void ApplyCodeBlockStyle(Wpf.Paragraph wpfPara, string? codeLanguage)
    {
        if (codeLanguage is null) return;
        wpfPara.FontFamily      = new WpfMedia.FontFamily("Consolas, D2Coding, monospace");
        wpfPara.Background      = new WpfMedia.SolidColorBrush(WpfMedia.Color.FromRgb(0xF8, 0xF8, 0xF8));
        wpfPara.BorderBrush     = new WpfMedia.SolidColorBrush(WpfMedia.Color.FromRgb(0xD0, 0xD0, 0xD0));
        wpfPara.BorderThickness = new Thickness(1);
        wpfPara.Padding         = new Thickness(MmToDip(3.0), MmToDip(1.5), MmToDip(3.0), MmToDip(1.5));
        var m = wpfPara.Margin;
        if (m.Top < 2 && m.Bottom < 2)
            wpfPara.Margin = new Thickness(m.Left, 4, m.Right, 4);
    }

    /// <summary>글자폭 != 100% 또는 자간 != 0 이면 Span(per-char InlineUIContainer 들), 그 외에는 Run 반환.
    /// LatexSource 가 있으면 WpfMath FormulaControl 로 렌더링.
    /// EmojiKey 가 있으면 Resources/Emojis/{Section}/{name}.png 를 Image 로 렌더링.</summary>
    public static Wpf.Inline BuildInline(Run run,
        IReadOnlyDictionary<string, int>? fnNums = null,
        IReadOnlyDictionary<string, int>? enNums = null)
    {
        // 각주/미주 참조 런 — 위첨자 숫자로 렌더링, Tag 에 원본 Run 보관.
        if (run.FootnoteId is { Length: > 0 } fnId)
        {
            var fnNum = 0;
            fnNums?.TryGetValue(fnId, out fnNum);
            var label = fnNum > 0 ? fnNum.ToString(System.Globalization.CultureInfo.InvariantCulture) : "†";
            var fnWpfRun = new Wpf.Run(label)
            {
                BaselineAlignment = BaselineAlignment.Superscript,
                FontSize          = PtToDip(8),
                Tag               = run,
            };
            return fnWpfRun;
        }
        if (run.EndnoteId is { Length: > 0 } enId)
        {
            var enNum = 0;
            enNums?.TryGetValue(enId, out enNum);
            var label = enNum > 0 ? enNum.ToString(System.Globalization.CultureInfo.InvariantCulture) : "‡";
            var enWpfRun = new Wpf.Run(label)
            {
                BaselineAlignment = BaselineAlignment.Superscript,
                FontSize          = PtToDip(8),
                Tag               = run,
            };
            return enWpfRun;
        }

        if (run.Field is { } fieldType)
            return BuildFieldInline(run, fieldType);

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

        // URL 이 있으면 WPF Hyperlink 로 감쌈 — Tag 에 원본 Run 보관(파서 라운드트립용).
        if (run.Url is { Length: > 0 } url)
        {
            var hl = new Wpf.Hyperlink(wpfRun);
            try { hl.NavigateUri = new Uri(url, UriKind.RelativeOrAbsolute); }
            catch { /* 잘못된 URI — NavigateUri 생략 */ }
            hl.Tag = run;
            return hl;
        }

        return wpfRun;
    }

    private static Wpf.Inline BuildFieldInline(Run run, FieldType fieldType)
    {
        var text = fieldType switch
        {
            FieldType.Date     => System.DateTime.Now.ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture),
            FieldType.Time     => System.DateTime.Now.ToString("HH:mm",      System.Globalization.CultureInfo.InvariantCulture),
            FieldType.Page     => "‹페이지›",
            FieldType.NumPages => "‹총페이지›",
            FieldType.Author   => string.IsNullOrEmpty(run.Text) ? "‹작성자›" : run.Text,
            FieldType.Title    => string.IsNullOrEmpty(run.Text) ? "‹제목›"   : run.Text,
            _                  => $"‹{fieldType}›",
        };

        return new Wpf.Run(text)
        {
            Tag        = run,
            Background = new WpfMedia.SolidColorBrush(WpfMedia.Color.FromArgb(45, 0, 102, 204)),
        };
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

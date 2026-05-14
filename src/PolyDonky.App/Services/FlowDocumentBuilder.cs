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

        var defaultFontFamily = !string.IsNullOrWhiteSpace(document.Metadata.DefaultFontFamily)
            ? document.Metadata.DefaultFontFamily + ", 맑은 고딕, Malgun Gothic, Segoe UI"
            : "맑은 고딕, Malgun Gothic, Segoe UI";
        var defaultFontSizePt = document.Metadata.DefaultFontSizePt > 0
            ? document.Metadata.DefaultFontSizePt
            : 11.0;

        var fd = new Wpf.FlowDocument
        {
            FontFamily  = new WpfMedia.FontFamily(defaultFontFamily),
            FontSize    = PtToDip(defaultFontSizePt),
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
    /// 리스트 종류·중첩 깊이·대소문자 지정에 따라 브라우저 기본 마커 스타일을 선택한다.
    /// - Bullet: disc → circle → square (≥2단계 동일).
    /// - OrderedDecimal: 모든 깊이 decimal.
    /// - OrderedAlpha:  upperCase 가 명시되면 그 값, null 이면 L0=Upper, L≥1=Lower 휴리스틱.
    /// - OrderedRoman:  같은 규칙.
    /// </summary>
    private static TextMarkerStyle MarkerStyleForLevel(ListKind kind, int level, bool? upperCase = null) => kind switch
    {
        ListKind.Bullet         => level switch { 0 => TextMarkerStyle.Disc, 1 => TextMarkerStyle.Circle, _ => TextMarkerStyle.Square },
        ListKind.OrderedAlpha   => (upperCase ?? level == 0) ? TextMarkerStyle.UpperLatin : TextMarkerStyle.LowerLatin,
        ListKind.OrderedRoman   => (upperCase ?? level == 0) ? TextMarkerStyle.UpperRoman : TextMarkerStyle.LowerRoman,
        _                       => TextMarkerStyle.Decimal,
    };

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
        var listStack = new Stack<(Wpf.List List, ListKind Kind, bool Hidden)>();

        // 지정 level·kind 의 WPF List 를 반환한다.
        // 스택이 부족하면 새 List 를 생성해 부모 ListItem 에 붙이고 push 한다.
        // 작업 목록(checked != null)은 마커가 ☐/☑ 텍스트로 단락 본문에 prepend 되므로
        // WPF MarkerStyle 은 None — 중복 마커 방지.
        // hideBullet=true 도 같은 방식으로 None 마커.
        Wpf.List EnsureList(int level, ListKind kind, int startIndex, bool? upperCase, bool isTaskList, bool hideBullet)
        {
            // 초과 레벨 팝
            while (listStack.Count > level + 1)
                listStack.Pop();

            // 현재 레벨에 같은 Kind/HideBullet 리스트가 있으면 재사용 — 두 속성 중 하나라도
            // 다르면 별개 리스트로 취급해 마커 표시 여부를 정확히 보존.
            if (listStack.Count == level + 1
                && listStack.Peek().Kind == kind
                && listStack.Peek().Hidden == hideBullet)
                return listStack.Peek().List;

            // 현재 레벨에 다른 Kind 리스트가 있으면 교체 준비
            if (listStack.Count == level + 1)
                listStack.Pop();

            // 레벨이 스택보다 깊으면 중간 레벨을 채운다 (정상 HTML 에선 발생하지 않음)
            while (listStack.Count < level)
            {
                var mid = new Wpf.List { MarkerStyle = MarkerStyleForLevel(ListKind.Bullet, listStack.Count) };
                AppendListToParent(mid);
                listStack.Push((mid, ListKind.Bullet, false));
            }

            var newList = new Wpf.List
            {
                MarkerStyle = (isTaskList || hideBullet)
                    ? TextMarkerStyle.None
                    : MarkerStyleForLevel(kind, level, upperCase),
            };
            if (kind != ListKind.Bullet && startIndex >= 1)
                newList.StartIndex = startIndex;
            AppendListToParent(newList);
            listStack.Push((newList, kind, hideBullet));
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

        // 연속된 ShapeObject 구간은 ShapeOrdering 정책(ZOrder + 자동 컨테인먼트 보정)에 따라
        // 그리는 순서를 재배열한다. 다른 블록이 끼어 있으면 그 위치는 유지.
        var orderedBlocks = ReorderShapeRuns(blocks);

        foreach (var block in orderedBlocks)
        {
            switch (block)
            {
                case Paragraph p when p.Style.ListMarker is { } marker:
                {
                    int level = Math.Max(0, marker.Level);
                    int start = marker.Kind != ListKind.Bullet && marker.OrderedNumber is { } s && s >= 1 ? s : 1;
                    bool isTaskList = marker.Checked is not null;
                    var list  = EnsureList(level, marker.Kind, start, marker.UpperCase, isTaskList, marker.HideBullet);
                    var wpfPara = BuildParagraph(p, outlineStyles, fnNums, enNums);
                    if (isTaskList)
                    {
                        // CSS 의 :before 가상 요소처럼 ☐/☑ 를 단락 첫머리에 직접 삽입.
                        // (WPF List MarkerStyle 만으로는 체크 상태를 표현할 수 없으므로 텍스트로 그린다.)
                        var checkRun = new Wpf.Run(marker.Checked == true ? "☑ " : "☐ ");
                        if (wpfPara.Inlines.FirstInline is { } firstInline)
                            wpfPara.Inlines.InsertBefore(firstInline, checkRun);
                        else
                            wpfPara.Inlines.Add(checkRun);
                    }
                    list.ListItems.Add(new Wpf.ListItem(wpfPara));
                    break;
                }

                case ThematicBreakBlock thb:
                    listStack.Clear();
                    target.Add(BuildThematicBreak(thb));
                    break;

                case Paragraph p:
                    listStack.Clear();
                    target.Add(BuildParagraph(p, outlineStyles, fnNums, enNums));
                    MergeAdjacentBlockquoteMargins(target);
                    break;

                case Table t:
                    listStack.Clear();
                    if (t.WrapMode == TableWrapMode.Block)
                    {
                        // <caption> 이 있으면 표 위에 가운데 정렬 단락으로 렌더링.
                        if (!string.IsNullOrEmpty(t.Caption))
                            target.Add(BuildTableCaption(t.Caption));
                        // CSS flex/grid 에서 변환된 레이아웃 표:
                        // - 회전 도형이 포함된 경우: BlockUIContainer(WPF Grid) → ClipToBounds=false 로 오버플로 허용.
                        // - 텍스트/목록만 있는 경우: Wpf.Table → AppendBlocks 재귀로 WPF List 구조를 올바르게 생성.
                        //   (BuildFlexContainer 의 BuildFlexLabel 은 ListMarker 를 TextBlock 으로 처리해 글머리 기호가 사라짐.)
                        if (t.IsFlexLayout)
                        {
                            bool hasRotatedShape = t.Rows.Any(row => row.Cells.Any(cell =>
                                cell.Blocks.Any(b => b is ShapeObject s && s.RotationAngleDeg != 0)));
                            target.Add(hasRotatedShape
                                ? BuildFlexContainer(t, outlineStyles)
                                : BuildTable(t, outlineStyles));
                        }
                        else
                            target.Add(BuildTable(t, outlineStyles));
                    }
                    else
                        target.Add(BuildTableAnchor(t));   // 오버레이 모드 — 앵커만 추가
                    break;

                case ImageBlock image:
                {
                    listStack.Clear();

                    // Inline + Above/Below 캡션은 별도 Paragraph 로 분리한다.
                    // 이유: WPF FlowDocument 의 BlockUIContainer 안에 들어간 UIElement 의
                    // ActualHeight 가 오프스크린 측정 RTB 에서 layout 미완료로 인해 0/과소
                    // 측정되는 케이스가 있다(특히 multi-row Grid 또는 visual tree 미부착 시).
                    // 캡션을 별도 Wpf.Paragraph 로 두면 텍스트 측정이 reliable 하게 동작해
                    // pagination 이 정확한 페이지 슬롯 배정을 한다(워드/한글도 동일 동작).
                    bool separateCaption = image.WrapMode == ImageWrapMode.Inline
                                        && image.ShowTitle
                                        && !string.IsNullOrWhiteSpace(image.Title)
                                        && image.TitlePosition is ImageTitlePosition.Above
                                                                or ImageTitlePosition.Below;

                    if (separateCaption && image.TitlePosition == ImageTitlePosition.Above)
                    {
                        var capPara = BuildImageCaptionParagraph(image);
                        if (capPara is not null) target.Add(capPara);
                    }

                    if (separateCaption)
                    {
                        // BuildImage 가 내부에서 캡션을 그리지 않도록 ShowTitle 을 임시로 끔.
                        // STA 스레드 단일 스레드라 동시성 위험 없음.
                        var savedShow = image.ShowTitle;
                        image.ShowTitle = false;
                        try { target.Add(BuildImage(image)); }
                        finally { image.ShowTitle = savedShow; }
                    }
                    else
                    {
                        target.Add(BuildImage(image));
                    }

                    if (separateCaption && image.TitlePosition == ImageTitlePosition.Below)
                    {
                        var capPara = BuildImageCaptionParagraph(image);
                        if (capPara is not null) target.Add(capPara);
                    }
                    break;
                }

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

                case ContainerBlock box:
                    listStack.Clear();
                    // flex-table 단독 ContainerBlock 특수 처리:
                    // Wpf.Section { BUC(Grid) } 구조는 (1) Section.Background 가 렌더링되지 않고
                    // (2) 인접 Paragraph 의 GetCharacterRect 가 Rect.Empty 를 반환하는 WPF 레이아웃
                    // 퀵을 유발한다. BUC 를 직접 배출하고 박스 스타일을 Grid 래핑 Border 에 적용함으로써
                    // Section 래퍼를 완전히 제거한다.
                    if (box.Children.Count == 1
                        && box.Children[0] is Table { IsFlexLayout: true } singleFlex)
                    {
                        // ContainerBlock 이 박스 스타일(배경·테두리·패딩)을 갖고 있으므로
                        // BuildFlexContainer 에서 boxStyle 파라미터로 직접 처리한다.
                        // BuildContainer(Section 래퍼) 를 쓰면 Section.Background/BorderBrush 가
                        // BUC 단독 자식 시 렌더링되지 않는 WPF FlowDocument 버그가 재발한다.
                        target.Add(BuildFlexContainer(singleFlex, outlineStyles, box));
                    }
                    else
                    {
                        target.Add(BuildContainer(box, outlineStyles, fnNums, enNums));
                    }
                    break;
            }
        }
    }

    /// <summary>MainWindow 에서 그룹 묶기 등 수동 그룹 생성에 사용. 자식 없는 빈 Section 반환.</summary>
    internal static Wpf.Section BuildContainerSection(ContainerBlock box)
        => BuildContainer(box, outlineStyles: null);

    /// <summary>박스 스타일을 가진 <see cref="ContainerBlock"/> 을 WPF FlowDocument 의 <see cref="Wpf.Section"/>
    /// 으로 빌드한다 — Section 은 BorderBrush/BorderThickness/Background/Padding 을 모두 지원하며 Block 트리를 그대로 품을 수 있다.
    /// 자식 블록은 본문과 동일한 dispatch 를 거쳐 Section.Blocks 에 추가된다.</summary>
    private static Wpf.Section BuildContainer(ContainerBlock box, OutlineStyleSet? outlineStyles,
        IReadOnlyDictionary<string,int>? fnNums = null, IReadOnlyDictionary<string,int>? enNums = null)
    {
        var section = new Wpf.Section { Tag = box };

        // 4면 보더 (단일 BorderBrush — 가장 먼저 명시된 색 채택).
        bool anyBorder = box.BorderTopPt > 0 || box.BorderRightPt > 0 ||
                         box.BorderBottomPt > 0 || box.BorderLeftPt > 0;
        if (anyBorder)
        {
            string? colorStr = box.BorderTopColor ?? box.BorderRightColor
                            ?? box.BorderBottomColor ?? box.BorderLeftColor;
            WpfMedia.Brush brush;
            if (!string.IsNullOrEmpty(colorStr))
            {
                try { brush = new WpfMedia.SolidColorBrush((WpfMedia.Color)WpfMedia.ColorConverter.ConvertFromString(colorStr)); }
                catch { brush = WpfMedia.Brushes.DimGray; }
            }
            else brush = WpfMedia.Brushes.DimGray;
            section.BorderBrush     = brush;
            section.BorderThickness = new Thickness(
                box.BorderLeftPt   > 0 ? PtToDip(box.BorderLeftPt)   : 0,
                box.BorderTopPt    > 0 ? PtToDip(box.BorderTopPt)    : 0,
                box.BorderRightPt  > 0 ? PtToDip(box.BorderRightPt)  : 0,
                box.BorderBottomPt > 0 ? PtToDip(box.BorderBottomPt) : 0);
        }

        if (!string.IsNullOrEmpty(box.BackgroundColor))
        {
            try
            {
                section.Background = new WpfMedia.SolidColorBrush(
                    (WpfMedia.Color)WpfMedia.ColorConverter.ConvertFromString(box.BackgroundColor));
            }
            catch { /* 잘못된 색 — 무시 */ }
        }

        section.Padding = new Thickness(
            MmToDip(box.PaddingLeftMm),
            MmToDip(box.PaddingTopMm),
            MmToDip(box.PaddingRightMm),
            MmToDip(box.PaddingBottomMm));
        // RTB 클립 방지: 오른쪽 테두리가 있을 때 최소 1 DIP 의 오른쪽 여백을 주어
        // Section.BorderBrush 및 borderOverlay 오른쪽 선이 RTB 클립 경계 안쪽에 그려지도록 보장.
        // Section.Margin.Right > 0 이면 WPF Section.BorderBrush 오른쪽 선이 안 그려지는 버그가
        // 있으나, 다중-자식(ShapeObject 등) 경로에서 발생하는 RTB 클립 문제가 더 크므로 허용한다.
        double rightSafetyDip = box.BorderRightPt > 0
            ? Math.Max(1.0, PtToDip(box.BorderRightPt)) : 0.0;
        section.Margin = new Thickness(0,
            MmToDip(box.MarginTopMm), rightSafetyDip, MmToDip(box.MarginBottomMm));

        // Group 역할: 별도 테두리 지정 없이 얇은 점선 테두리로 시각 구분.
        if (box.Role == ContainerRole.Group && !anyBorder)
        {
            section.BorderBrush     = new WpfMedia.SolidColorBrush(
                WpfMedia.Color.FromRgb(0x80, 0x80, 0xCC));
            section.BorderThickness = new Thickness(1);
            section.Padding         = new Thickness(Math.Max(section.Padding.Left,  4),
                                                    Math.Max(section.Padding.Top,   4),
                                                    Math.Max(section.Padding.Right,  4),
                                                    Math.Max(section.Padding.Bottom, 4));
        }

        // 자식 dispatch — 본문 AppendBlocks 를 재사용해 일관 처리.
        AppendBlocks(section.Blocks, box.Children, outlineStyles, fnNums, enNums);

        // WPF FlowDocument 에서 Section.Background/BorderBrush 는 자식이 BlockUIContainer 하나뿐일 때
        // 렌더링되지 않는다. 이 경우 자식 UIElement 를 WPF Grid 로 감싸 배경·테두리·패딩을 직접 적용한다.
        if (section.Blocks.Count == 1
            && section.Blocks.FirstBlock is Wpf.BlockUIContainer singleBuc
            && singleBuc.Child is FrameworkElement singleFe
            && (section.Background is not null || section.BorderBrush is not null))
        {

            // singleFe 가 이미 singleBuc 의 논리 자식이므로 먼저 연결을 끊어야 한다.
            singleBuc.Child = null;

            // 배경·패딩은 contentWrap 에 적용.
            // ClipToBounds=true: SVG 오버플로가 contentWrap 경계 밖(BUC 부모 공간)으로
            // 올라가 borderOverlay(z=1) 를 덮는 현상을 방지한다.
            var contentWrap = new System.Windows.Controls.Border
            {
                Child        = singleFe,
                ClipToBounds = true,
            };
            if (section.Background is not null)
            {
                contentWrap.Background = section.Background;
                section.Background     = null;
            }
            var spad = section.Padding;
            if (spad.Left != 0 || spad.Top != 0 || spad.Right != 0 || spad.Bottom != 0)
            {
                contentWrap.Padding = spad;
                section.Padding     = new Thickness(0);
            }

            if (section.BorderBrush is not null)
            {
                // SVG/도형 콘텐츠가 ClipToBounds=false 오버플로로 Border 선을 덮지 않도록
                // Grid 에 ① contentWrap(z=0) + ② borderOverlay(z=1) 구조로 테두리를 위에 그린다.
                var wrapperGrid = new System.Windows.Controls.Grid { ClipToBounds = false };
                wrapperGrid.Children.Add(contentWrap);
                wrapperGrid.Children.Add(new System.Windows.Controls.Border
                {
                    BorderBrush      = section.BorderBrush,
                    BorderThickness  = section.BorderThickness,
                    IsHitTestVisible = false,
                });
                section.BorderBrush     = null;
                section.BorderThickness = new Thickness(0);
                singleBuc.Child = wrapperGrid;
            }
            else
            {
                singleBuc.Child = contentWrap;
            }
        }

        // WPF FlowDocument 에서 Section.Background 는 자식이 Wpf.Table 하나뿐일 때
        // 렌더링되지 않는 경우가 있다. 이 경우 배경을 Table 에 직접 설정하는 방식으로 우회한다.
        if (section.Background is not null
            && section.Blocks.Count == 1
            && section.Blocks.FirstBlock is Wpf.Table innerWpfTable
            && innerWpfTable.Background is null)
        {
            innerWpfTable.Background = section.Background;
            section.Background = null;
        }

        return section;
    }

    /// <summary>
    /// 연속된 <see cref="ShapeObject"/> 묶음을 <see cref="ShapeOrdering"/> 정책에 따라 재배열한다.
    /// 도형이 아닌 블록의 위치는 절대 바꾸지 않으며, 도형이 한 개뿐이거나 0개인 묶음은 그대로 통과시킨다.
    /// </summary>
    private static IList<Block> ReorderShapeRuns(IList<Block> blocks)
    {
        // 도형 묶음이 없으면 비용 0 — 원본 리스트 반환.
        bool anyConsecutive = false;
        for (int i = 1; i < blocks.Count; i++)
        {
            if (blocks[i] is ShapeObject && blocks[i - 1] is ShapeObject)
            {
                anyConsecutive = true; break;
            }
        }
        if (!anyConsecutive) return blocks;

        var result = new List<Block>(blocks.Count);
        int idx = 0;
        while (idx < blocks.Count)
        {
            if (blocks[idx] is ShapeObject)
            {
                int end = idx;
                while (end < blocks.Count && blocks[end] is ShapeObject) end++;
                if (end - idx >= 2)
                {
                    var run = new ShapeObject[end - idx];
                    for (int k = 0; k < run.Length; k++) run[k] = (ShapeObject)blocks[idx + k];
                    foreach (var s in ShapeOrdering.OrderForRendering(run)) result.Add(s);
                }
                else
                {
                    result.Add(blocks[idx]);
                }
                idx = end;
            }
            else
            {
                result.Add(blocks[idx]);
                idx++;
            }
        }
        return result;
    }

    /// <summary>TocBlock 을 시각적 BlockUIContainer 로 빌드한다. Tag = TocBlock 으로 라운드트립 가능.</summary>
    public static Wpf.BlockUIContainer BuildTocBlock(TocBlock toc)
    {
        var stack = new System.Windows.Controls.StackPanel
        {
            Margin              = new Thickness(2),
            HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch,
        };

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
            BorderBrush         = new WpfMedia.SolidColorBrush(WpfMedia.Color.FromRgb(0xC0, 0xC0, 0xC0)),
            BorderThickness     = new Thickness(1),
            CornerRadius        = new System.Windows.CornerRadius(3),
            Padding             = new Thickness(10),
            Margin              = new Thickness(0, 4, 0, 4),
            Background          = new WpfMedia.SolidColorBrush(WpfMedia.Color.FromArgb(15, 0, 0, 0)),
            HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch,
            Child               = stack,
        };

        return new Wpf.BlockUIContainer(border) { Tag = toc };
    }

    // CSS flex/grid 컨테이너 전용 렌더러 — Wpf.Table 대신 BlockUIContainer(WPF Grid) 를 사용해
    // 회전 도형 등의 시각적 오버플로가 TableCell 렌더러의 기하 클리핑에 걸리지 않도록 한다.
    // boxStyle 이 non-null 이면 Grid 를 WPF Border 로 감싸 배경·테두리·패딩을 직접 적용한다
    // (Wpf.Section.Background/BorderBrush 는 BUC 단독 자식 시 렌더링 안 되는 WPF 버그 우회).
    private static Wpf.Block BuildFlexContainer(Table table, OutlineStyleSet? outlineStyles,
        ContainerBlock? boxStyle = null)
    {
        int colCount = Math.Max(table.Columns.Count, 1);

        var grid = new System.Windows.Controls.Grid { ClipToBounds = false };
        for (int c = 0; c < colCount; c++)
            grid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition
                { Width = new GridLength(1, GridUnitType.Star) });
        grid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition
            { Height = GridLength.Auto });

        if (table.Rows.Count > 0)
        {
            int colIdx = 0;
            foreach (var cell in table.Rows[0].Cells)
            {
                double gapRight = cell.PaddingRightMm > 0 ? MmToDip(cell.PaddingRightMm) : 0;
                var cellPanel = new System.Windows.Controls.StackPanel
                {
                    Orientation         = System.Windows.Controls.Orientation.Vertical,
                    ClipToBounds        = false,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Margin              = new Thickness(0, 0, gapRight, 0),
                };

                foreach (var block in cell.Blocks)
                {
                    if (block is ShapeObject shape && shape.WrapMode != ImageWrapMode.InFrontOfText
                                                   && shape.WrapMode != ImageWrapMode.BehindText)
                    {
                        double wDip = MmToDip(shape.WidthMm);
                        double hDip = MmToDip(shape.HeightMm);
                        var visual = BuildShapeVisual(shape, wDip, hDip);
                        var imgHA = shape.HAlign switch
                        {
                            ImageHAlign.Center => HorizontalAlignment.Center,
                            ImageHAlign.Right  => HorizontalAlignment.Right,
                            _                  => HorizontalAlignment.Left,
                        };
                        visual.HorizontalAlignment = imgHA;
                        var host = BuildInlineRotationHost(visual, shape.RotationAngleDeg, wDip, hDip, imgHA, useBboxLayout: true);
                        host.Margin = new Thickness(0, MmToDip(shape.MarginTopMm), 0, MmToDip(shape.MarginBottomMm));
                        cellPanel.Children.Add(host);
                    }
                    else if (block is Paragraph para)
                    {
                        cellPanel.Children.Add(BuildFlexLabel(para));
                    }
                }

                System.Windows.Controls.Grid.SetColumn(cellPanel, colIdx);
                grid.Children.Add(cellPanel);
                colIdx++;
            }
        }

        // boxStyle 이 있으면 배경·테두리·패딩을 적용한다.
        // 핵심 제약: Grid 의 ClipToBounds=false 때문에 도형이 Grid 레이아웃 경계 밖으로 시각적으로
        // 넘쳐나올 수 있다. Border { Grid } 구조에서는 Grid 콘텐츠가 Border 그림 위에 덮여 그려지는
        // WPF z-order 때문에 border 선(특히 오른쪽)이 도형에 가려진다.
        // 해결: 래퍼 Grid 안에 contentGrid(아래)→borderOverlay(위) 순으로 배치해
        // border 선이 항상 콘텐츠 위에 렌더링되도록 보장한다.
        FrameworkElement content = grid;
        if (boxStyle is not null)
        {
            double bL = boxStyle.BorderLeftPt   > 0 ? PtToDip(boxStyle.BorderLeftPt)   : 0;
            double bT = boxStyle.BorderTopPt    > 0 ? PtToDip(boxStyle.BorderTopPt)    : 0;
            double bR = boxStyle.BorderRightPt  > 0 ? PtToDip(boxStyle.BorderRightPt)  : 0;
            double bB = boxStyle.BorderBottomPt > 0 ? PtToDip(boxStyle.BorderBottomPt) : 0;
            double pL = MmToDip(boxStyle.PaddingLeftMm);
            double pT = MmToDip(boxStyle.PaddingTopMm);
            double pR = MmToDip(boxStyle.PaddingRightMm);
            double pB = MmToDip(boxStyle.PaddingBottomMm);

            // 콘텐츠 Grid 를 border+padding 만큼 안으로 들여쓴다.
            grid.Margin = new Thickness(bL + pL, bT + pT, bR + pR, bB + pB);

            var wrapperGrid = new System.Windows.Controls.Grid { ClipToBounds = false };

            // 배경 — 래퍼 Grid 에 직접 설정해 padding 영역까지 채워짐.
            if (!string.IsNullOrEmpty(boxStyle.BackgroundColor))
            {
                try
                {
                    wrapperGrid.Background = new WpfMedia.SolidColorBrush(
                        (WpfMedia.Color)WpfMedia.ColorConverter.ConvertFromString(boxStyle.BackgroundColor));
                }
                catch { /* 잘못된 색 무시 */ }
            }

            wrapperGrid.Children.Add(grid); // ① 콘텐츠 (아래 레이어)

            // 테두리 오버레이 — Grid 콘텐츠 위에 그려져 도형 오버플로에 가려지지 않는다.
            bool anyBorder = bL > 0 || bT > 0 || bR > 0 || bB > 0;
            if (anyBorder)
            {
                string? colorStr = boxStyle.BorderTopColor ?? boxStyle.BorderRightColor
                                ?? boxStyle.BorderBottomColor ?? boxStyle.BorderLeftColor;
                WpfMedia.Brush brush;
                if (!string.IsNullOrEmpty(colorStr))
                {
                    try { brush = new WpfMedia.SolidColorBrush((WpfMedia.Color)WpfMedia.ColorConverter.ConvertFromString(colorStr)); }
                    catch { brush = WpfMedia.Brushes.DimGray; }
                }
                else brush = WpfMedia.Brushes.DimGray;

                var borderOverlay = new System.Windows.Controls.Border
                {
                    BorderBrush      = brush,
                    BorderThickness  = new Thickness(bL, bT, bR, bB),
                    IsHitTestVisible = false,   // 인터랙션은 콘텐츠 Grid 에 위임
                };
                wrapperGrid.Children.Add(borderOverlay); // ② 테두리 오버레이 (위 레이어)
            }

            content = wrapperGrid;
        }

        double marginTop    = boxStyle is not null ? MmToDip(boxStyle.MarginTopMm)    : MmToDip(table.OuterMarginTopMm);
        double marginBottom = boxStyle is not null ? MmToDip(boxStyle.MarginBottomMm) : MmToDip(table.OuterMarginBottomMm);

        return new Wpf.BlockUIContainer(content)
        {
            Tag    = table,
            Margin = new Thickness(
                MmToDip(table.OuterMarginLeftMm),
                marginTop,
                MmToDip(table.OuterMarginRightMm),
                marginBottom),
        };
    }

    // flex 셀 안의 텍스트 단락(레이블 등)을 WPF TextBlock 으로 변환.
    private static System.Windows.Controls.TextBlock BuildFlexLabel(Paragraph para)
    {
        var ta = para.Style.Alignment switch
        {
            Alignment.Center  => TextAlignment.Center,
            Alignment.Right   => TextAlignment.Right,
            Alignment.Justify => TextAlignment.Justify,
            _                 => TextAlignment.Center, // flex 셀은 기본 가운데
        };
        var tb = new System.Windows.Controls.TextBlock
        {
            TextAlignment = ta,
            TextWrapping  = System.Windows.TextWrapping.Wrap,
        };
        if (para.Style.SpaceBeforePt > 0)
            tb.Margin = new Thickness(0, PtToDip(para.Style.SpaceBeforePt), 0, 0);

        foreach (var run in para.Runs)
        {
            if (run.Text is null) continue;
            var wr = new System.Windows.Documents.Run(run.Text);
            if (run.Style.FontSizePt > 0)
                wr.FontSize = PtToDip(run.Style.FontSizePt);
            if (!string.IsNullOrEmpty(run.Style.FontFamily))
                wr.FontFamily = new WpfMedia.FontFamily(run.Style.FontFamily);
            if (run.Style.Bold   == true) wr.FontWeight = System.Windows.FontWeights.Bold;
            if (run.Style.Italic == true) wr.FontStyle  = System.Windows.FontStyles.Italic;
            if (run.Style.Foreground is { } fg)
                wr.Foreground = new WpfMedia.SolidColorBrush(
                    WpfMedia.Color.FromArgb(fg.A, fg.R, fg.G, fg.B));
            tb.Inlines.Add(wr);
        }

        return tb;
    }

    internal static Wpf.Table BuildTable(Table table, OutlineStyleSet? outlineStyles = null)
    {
        var wtable = new Wpf.Table { CellSpacing = 0 };

        ApplyTableLevelPropertiesToWpf(wtable, table);

        foreach (var col in table.Columns)
        {
            // WidthMm > 0 이면 명시 폭, 아니면 Star(1*) — 가용 폭을 균등 분배.
            // Auto 로 두면 셀 콘텐츠 기준으로 좁게 잡혀 텍스트 줄바꿈이 과도해지고
            // 셀 높이가 비정상적으로 커져 표 전체가 페이지를 넘기는 증상이 발생한다.
            var width = col.WidthMm > 0
                ? new GridLength(MmToDip(col.WidthMm))
                : new GridLength(1, GridUnitType.Star);
            wtable.Columns.Add(new Wpf.TableColumn { Width = width });
        }

        var rowGroup = new Wpf.TableRowGroup();
        wtable.RowGroups.Add(rowGroup);

        var headerBrush = new WpfMedia.SolidColorBrush(WpfMedia.Color.FromRgb(0xE8, 0xEA, 0xED));
        headerBrush.Freeze();

        int rowCount = table.Rows.Count;
        int colCount = Math.Max(table.Columns.Count, table.Rows.Count > 0 ? table.Rows[0].Cells.Count : 0);

        int rowIdx = 0;
        foreach (var row in table.Rows)
        {
            var wrow = new Wpf.TableRow();
            if (row.IsHeader)
                wrow.Background = headerBrush;

            int colIdx = 0;
            foreach (var cell in row.Cells)
            {
                int rowSpan = Math.Max(cell.RowSpan, 1);
                int colSpan = Math.Max(cell.ColumnSpan, 1);

                var wcell = new Wpf.TableCell
                {
                    ColumnSpan = colSpan,
                    RowSpan    = rowSpan,
                };

                bool atTop    = rowIdx == 0;
                bool atBottom = rowIdx + rowSpan - 1 >= rowCount - 1;
                bool atLeft   = colIdx == 0;
                bool atRight  = colIdx + colSpan - 1 >= colCount - 1;

                ApplyCellPropertiesToWpf(wcell, cell, row.IsHeader, table,
                    atTop, atBottom, atLeft, atRight, row);
                AppendBlocks(wcell.Blocks, cell.Blocks, outlineStyles);
                if (wcell.Blocks.Count == 0)
                    wcell.Blocks.Add(new Wpf.Paragraph(new Wpf.Run(string.Empty)));

                ApplyVerticalAlignmentToCell(wcell, cell, row);
                wrow.Cells.Add(wcell);
                colIdx += colSpan;
            }
            rowGroup.Rows.Add(wrow);
            rowIdx++;
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

        // 표 외곽선 — Wpf.Table 자체의 BorderBrush/Thickness 는 셀 외곽선과 겹치므로 0 으로 둔다.
        // 외곽선은 가장자리 셀이 자기 면(top/bottom/left/right) 으로 직접 그린다 (ApplyCellPropertiesToWpf).
        wtable.BorderBrush     = null;
        wtable.BorderThickness = new Thickness(0);
    }

    /// <summary>표·셀 면별 테두리 cascade — 셀 면 지정 > 표 외곽/안쪽 면 > 셀/표 공통값.</summary>
    private static CellBorderSide ResolveBorderSide(
        CellBorderSide? cellSide,
        bool isTableEdge,
        CellBorderSide? tableEdge,
        CellBorderSide? tableInner,
        TableCell cell,
        Table? tableDefaults)
    {
        if (cellSide.HasValue) return cellSide.Value;

        var edgeOrInner = isTableEdge ? tableEdge : tableInner;
        if (edgeOrInner.HasValue) return edgeOrInner.Value;

        // 공통값 폴백 — 셀 공통이 없으면 표 공통.
        double thk = cell.BorderThicknessPt > 0
            ? cell.BorderThicknessPt
            : (tableDefaults?.BorderThicknessPt ?? 0);
        string? clr = !string.IsNullOrEmpty(cell.BorderColor)
            ? cell.BorderColor
            : tableDefaults?.BorderColor;
        return new CellBorderSide(thk, clr);
    }

    internal static void ApplyCellPropertiesToWpf(
        Wpf.TableCell wcell,
        TableCell cell,
        bool isHeader,
        Table? tableDefaults = null,
        bool atTopEdge    = false,
        bool atBottomEdge = false,
        bool atLeftEdge   = false,
        bool atRightEdge  = false,
        TableRow? row     = null)
    {
        // 면별 테두리 cascade — 셀 면 지정 > 표 외곽/안쪽 면 > 공통값.
        // 안쪽 면(table 내부) 일 때는 InnerBorderHorizontal/Vertical 을 default 로 쓴다.
        var top    = ResolveBorderSide(cell.BorderTop,    atTopEdge,    tableDefaults?.BorderTop,
                                       tableDefaults?.InnerBorderHorizontal, cell, tableDefaults);
        var bottom = ResolveBorderSide(cell.BorderBottom, atBottomEdge, tableDefaults?.BorderBottom,
                                       tableDefaults?.InnerBorderHorizontal, cell, tableDefaults);
        var left   = ResolveBorderSide(cell.BorderLeft,   atLeftEdge,   tableDefaults?.BorderLeft,
                                       tableDefaults?.InnerBorderVertical,   cell, tableDefaults);
        var right  = ResolveBorderSide(cell.BorderRight,  atRightEdge,  tableDefaults?.BorderRight,
                                       tableDefaults?.InnerBorderVertical,   cell, tableDefaults);

        // Wpf.TableCell 은 단일 BorderBrush 만 지원 — 면 색상 중 가장 먼저 명시된 것을 채택.
        // (모든 면이 같은 색이면 자연스러우며, 다른 색이 섞여 있으면 시각적 한계가 있음.)
        var pickedColor = top.Color ?? bottom.Color ?? left.Color ?? right.Color
                       ?? cell.BorderColor ?? tableDefaults?.BorderColor;
        var borderColor = TryParseColor(pickedColor) ?? WpfMedia.Color.FromRgb(0xC8, 0xC8, 0xC8);

        // 두께가 명시되지 않은 면은 셀/표 공통 두께 → 그것도 0 이면 0 (해당 면 미표시).
        // 단, 공통값이 모두 0 이고 면별도 모두 0 이면 fallback 0.75pt 적용 (기존 동작 호환).
        double commonThk = cell.BorderThicknessPt > 0
            ? cell.BorderThicknessPt
            : (tableDefaults?.BorderThicknessPt ?? 0);
        // 면별·공통 두께가 모두 0 이면 테두리 없음. 0.75pt fallback 은 의도하지 않은
        // 외곽선(레이아웃 표·CSS 도형 그리드 등)을 생성하므로 제거.
        double FallbackThk(double sideThk) => sideThk > 0 ? sideThk
                                            : commonThk > 0 ? commonThk
                                            : 0;

        wcell.BorderBrush     = new WpfMedia.SolidColorBrush(borderColor);
        // border-collapse 시뮬레이션: WPF Table 은 셀별로 보더를 그려 인접 셀의 공유 모서리에서
        // doubled 라인이 생긴다. tableDefaults.BorderCollapse=true 인 경우 셀의 위/왼쪽 보더를
        // 가장자리에서만 그려 한쪽(오른쪽 + 아래) 으로 일관 — 인접 셀 사이엔 단일 라인만 남게 한다.
        // 단, 셀에서 명시적으로 설정한 테두리는 BorderCollapse 를 무시하고 적용한다.
        bool collapse = tableDefaults?.BorderCollapse ?? true;
        double leftDip   = PtToDip(FallbackThk(left.ThicknessPt));
        double topDip    = PtToDip(FallbackThk(top.ThicknessPt));
        double rightDip  = PtToDip(FallbackThk(right.ThicknessPt));
        double bottomDip = PtToDip(FallbackThk(bottom.ThicknessPt));
        if (collapse)
        {
            // 셀이 명시적으로 설정하지 않은 경우에만 BorderCollapse 적용
            if (!atTopEdge && cell.BorderTop is null)  topDip  = 0;
            if (!atLeftEdge && cell.BorderLeft is null) leftDip = 0;
        }
        wcell.BorderThickness = new Thickness(leftDip, topDip, rightDip, bottomDip);

        double defTop    = tableDefaults?.DefaultCellPaddingTopMm    > 0 ? tableDefaults.DefaultCellPaddingTopMm    : Table.FallbackCellPaddingVerticalMm;
        double defBottom = tableDefaults?.DefaultCellPaddingBottomMm > 0 ? tableDefaults.DefaultCellPaddingBottomMm : Table.FallbackCellPaddingVerticalMm;
        double defLeft   = tableDefaults?.DefaultCellPaddingLeftMm   > 0 ? tableDefaults.DefaultCellPaddingLeftMm   : Table.FallbackCellPaddingHorizontalMm;
        double defRight  = tableDefaults?.DefaultCellPaddingRightMm  > 0 ? tableDefaults.DefaultCellPaddingRightMm  : Table.FallbackCellPaddingHorizontalMm;

        double padTop   = MmToDip(cell.PaddingTopMm    > 0 ? cell.PaddingTopMm    : defTop);
        double padBottom= MmToDip(cell.PaddingBottomMm > 0 ? cell.PaddingBottomMm : defBottom);
        double padLeft  = MmToDip(cell.PaddingLeftMm   > 0 ? cell.PaddingLeftMm   : defLeft);
        double padRight = MmToDip(cell.PaddingRightMm  > 0 ? cell.PaddingRightMm  : defRight);
        wcell.Padding = new Thickness(padLeft, padTop, padRight, padBottom);

        // 배경색: 셀 지정 > 행 지정 > null(투명)
        var bgColor = !string.IsNullOrEmpty(cell.BackgroundColor) ? cell.BackgroundColor
                    : (!string.IsNullOrEmpty(row?.BackgroundColor) ? row!.BackgroundColor : null);
        if (bgColor is not null && TryParseColor(bgColor) is { } bg)
            wcell.Background = new WpfMedia.SolidColorBrush(bg);
        else
            wcell.Background = null;

        wcell.FontWeight = isHeader ? FontWeights.SemiBold : FontWeights.Normal;
        ApplyCellTextAlign(wcell, cell.TextAlign);
    }

    /// <summary>
    /// 셀의 수직 정렬을 BlockUIContainer + Grid 로 구현한다.
    /// TableCell 은 TextElement 이므로 VerticalAlignment 를 직접 지원하지 않지만,
    /// BlockUIContainer 내 Grid 는 FrameworkElement 이므로 지원한다.
    /// </summary>
    private static void ApplyVerticalAlignmentToCell(Wpf.TableCell wcell, TableCell cell, TableRow? row)
    {
        // 유효한 수직 정렬 결정: 셀 > 행 > 기본값(Top)
        var align = cell.VerticalAlign;
        if (align == CellVerticalAlign.Top && row?.VerticalAlign.HasValue == true)
            align = row.VerticalAlign.Value;
        if (align == CellVerticalAlign.Top)
            return; // 기본값이면 처리 안 함

        // 셀의 기존 블록들을 모두 수집
        var originalBlocks = wcell.Blocks.Cast<Wpf.Block>().ToList();
        wcell.Blocks.Clear();

        // Grid 생성: 높이는 자동으로 콘텐츠에 맞춤, VerticalAlignment 는 셀 내 정렬 결정
        var grid = new System.Windows.Controls.Grid
        {
            HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch,
            VerticalAlignment = align switch
            {
                CellVerticalAlign.Middle => System.Windows.VerticalAlignment.Center,
                CellVerticalAlign.Bottom => System.Windows.VerticalAlignment.Bottom,
                _ => System.Windows.VerticalAlignment.Top,
            },
        };

        // StackPanel 에 모든 블록을 담는다
        var stack = new System.Windows.Controls.StackPanel
        {
            Orientation = System.Windows.Controls.Orientation.Vertical,
        };

        foreach (var block in originalBlocks)
        {
            // 블록을 UIElement 로 변환하여 StackPanel 에 추가
            if (block is Wpf.Paragraph para)
            {
                // 단락을 TextBlock 으로 변환 (가장 일반적인 경우)
                var tb = new System.Windows.Controls.TextBlock
                {
                    TextWrapping = System.Windows.TextWrapping.Wrap,
                    TextAlignment = para.TextAlignment,
                    Foreground = para.Foreground ?? System.Windows.Media.Brushes.Black,
                    Background = para.Background,
                    Padding = para.Padding,
                    Margin = para.Margin,
                    FontFamily = para.FontFamily,
                    FontSize = para.FontSize,
                    FontWeight = para.FontWeight,
                    FontStyle = para.FontStyle,
                };

                // 단락의 inlines 를 TextBlock 에 복사
                foreach (var inline in para.Inlines)
                    tb.Inlines.Add(CopyInline(inline));

                stack.Children.Add(tb);
            }
            else if (block is Wpf.BlockUIContainer buc)
            {
                // BlockUIContainer 는 그대로 추가
                stack.Children.Add(buc.Child as System.Windows.UIElement ?? new System.Windows.Controls.TextBlock { Text = "(UI element)" });
            }
            else if (block is Wpf.Table table)
            {
                // 표는 복잡하므로 별도 처리 필요 — 현재는 placeholder
                stack.Children.Add(new System.Windows.Controls.TextBlock { Text = "[표 콘텐츠]" });
            }
            else
            {
                // 기타 블록 타입은 텍스트 placeholder
                stack.Children.Add(new System.Windows.Controls.TextBlock { Text = "[블록]" });
            }
        }

        grid.Children.Add(stack);

        // Grid 를 BlockUIContainer 로 감싸서 wcell.Blocks 에 추가
        wcell.Blocks.Add(new Wpf.BlockUIContainer(grid));
    }

    /// <summary>WPF Inline 을 복사한다. (Tag, 스타일 포함)</summary>
    private static Wpf.Inline CopyInline(Wpf.Inline inline)
    {
        if (inline is Wpf.Run run)
        {
            return new Wpf.Run(run.Text)
            {
                Tag = run.Tag,
                Foreground = run.Foreground,
                Background = run.Background,
                FontSize = run.FontSize,
                FontWeight = run.FontWeight,
                FontStyle = run.FontStyle,
                TextDecorations = run.TextDecorations,
                BaselineAlignment = run.BaselineAlignment,
            };
        }
        else if (inline is Wpf.Span span)
        {
            var newSpan = new Wpf.Span
            {
                Tag = span.Tag,
                Foreground = span.Foreground,
                Background = span.Background,
                FontSize = span.FontSize,
                FontWeight = span.FontWeight,
                FontStyle = span.FontStyle,
                TextDecorations = span.TextDecorations,
            };
            foreach (var child in span.Inlines)
                newSpan.Inlines.Add(CopyInline(child));
            return newSpan;
        }
        else if (inline is Wpf.Hyperlink link)
        {
            var newLink = new Wpf.Hyperlink()
            {
                Tag = link.Tag,
                NavigateUri = link.NavigateUri,
                Foreground = link.Foreground,
                Background = link.Background,
                FontSize = link.FontSize,
                FontWeight = link.FontWeight,
                FontStyle = link.FontStyle,
                TextDecorations = link.TextDecorations,
            };
            foreach (var child in link.Inlines)
                newLink.Inlines.Add(CopyInline(child));
            return newLink;
        }
        else if (inline is Wpf.InlineUIContainer iuc)
        {
            return new Wpf.InlineUIContainer(iuc.Child) { Tag = iuc.Tag };
        }
        else
        {
            // 기타 inline: 그냥 반환
            return inline;
        }
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

    /// <summary>
    /// <c>ThematicBreakBlock</c> 을 <c>BlockUIContainer</c> + <c>Border</c> (RichTextBox.ActualWidth 바인딩)
    /// 로 렌더링한다.
    /// <para>
    /// 이전 시도들이 실패한 원인:
    /// (a) <c>Wpf.Paragraph.BorderBottom</c> + <c>FontSize=1/LineHeight=1</c> — 단락 높이가 너무
    ///     작아 보더가 시각적으로 사라지거나, FlowDocument 가 1px 높이를 정밀하게 그리지 못함.
    /// (b) <c>BlockUIContainer</c> + <c>Rectangle</c>/<c>Grid</c> + <c>HorizontalAlignment.Stretch</c>
    ///     — FlowDocument 의 첫 Measure 가 infinite 폭을 패스해서 자식 요소가 폭=0 으로 측정되고,
    ///     그 결과 폭=0 으로 Arrange 됨.
    /// </para>
    /// <para>
    /// 해결: <c>Border</c> 의 <c>Width</c> 를 ancestor <c>FlowDocumentScrollViewer</c> /
    /// <c>RichTextBox</c> 의 <c>ActualWidth</c> 에 바인딩해 명시적으로 컬럼 폭을 받는다.
    /// 바인딩이 visual tree 부착 후 동작하므로, 초기 0 폭 측정 문제를 우회한다.
    /// </para>
    /// </summary>
    internal static Wpf.Block BuildThematicBreak(ThematicBreakBlock thb)
    {
        WpfMedia.Color lineColor = WpfMedia.Color.FromRgb(0xAA, 0xAA, 0xAA);
        if (!string.IsNullOrEmpty(thb.LineColor))
        {
            try { lineColor = (WpfMedia.Color)WpfMedia.ColorConverter.ConvertFromString(thb.LineColor); }
            catch { /* 파싱 실패 시 기본 회색 유지 */ }
        }
        double marginV     = thb.MarginPt    > 0 ? PtToDip(thb.MarginPt)    : 6;
        double thicknessV  = thb.ThicknessPt > 0 ? PtToDip(thb.ThicknessPt) : 1;
        var brush = new WpfMedia.SolidColorBrush(lineColor);

        // 실선(Solid) — Wpf.Paragraph + Background. Paragraph 는 자연스럽게 본문 폭을 채우므로
        // BlockUIContainer + Stretch + Width 바인딩 패턴이 무한대 Measure / 비-RTB ancestor
        // 컨텍스트에서 0 픽셀로 collapse 하던 문제를 회피한다. Background 로 단락 전체를 색칠해
        // BorderTop 이 두께 < 단락 높이 케이스에서 보이지 않던 문제도 같이 해결.
        // FontSize/LineHeight 를 정확히 thicknessV 로 잡아 단락 자체가 바로 그 두께의 가로선이 되도록.
        if (thb.LineStyle == ThematicLineStyle.Solid)
        {
            // FontSize 는 1 이상 권장 — 너무 작으면 WPF 가 텍스트 측정을 거부해 단락 자체가 0 높이로 collapse 한다.
            double lineDip = Math.Max(thicknessV, 1);
            var hrPara = new Wpf.Paragraph(new Wpf.Run("​"))   // ZWSP — 빈 단락 collapse 방지
            {
                // Tag = Core ThematicBreakBlock 인스턴스 — 페이지네이션 (FlowDocumentPaginationAdapter
                // 가 'Tag is Block coreBlock' 으로 필터링) 과 파서 (Tag is ThematicBreakBlock thbCore)
                // 가 동일한 식별을 공유한다. sentinel object 로 두면 페이지네이션이 HR 을 통째로 건너뛰어
                // 어떤 페이지 슬라이스에도 HR 이 들어가지 않아 화면에 안 보이는 결정적 버그가 발생.
                Tag                  = thb,
                Background           = brush,
                Foreground           = brush,                  // ZWSP 가 색깔 차이로 노출되지 않게 동일 색.
                FontSize             = lineDip,
                LineHeight           = lineDip,
                LineStackingStrategy = LineStackingStrategy.BlockLineHeight,
                Padding              = new Thickness(0),
                Margin               = new Thickness(0, marginV, 0, marginV),
            };
            return hrPara;
        }

        FrameworkElement line;
        if (thb.LineStyle == ThematicLineStyle.Double)
        {
            // 이중선 — Rectangle 두 개 + 사이 간격. (단색 BorderTop 으로는 이중선 표현 불가하므로 BlockUIContainer 폴백.)
            var stack = new System.Windows.Controls.StackPanel
            {
                Orientation         = System.Windows.Controls.Orientation.Vertical,
                HorizontalAlignment = HorizontalAlignment.Stretch,
            };
            stack.Children.Add(new System.Windows.Shapes.Rectangle
            {
                Fill                = brush,
                Height              = thicknessV,
                HorizontalAlignment = HorizontalAlignment.Stretch,
                SnapsToDevicePixels = true,
            });
            stack.Children.Add(new System.Windows.Shapes.Rectangle
            {
                Fill                = brush,
                Height              = thicknessV,
                Margin              = new Thickness(0, thicknessV, 0, 0),
                HorizontalAlignment = HorizontalAlignment.Stretch,
                SnapsToDevicePixels = true,
            });
            line = stack;
        }
        else
        {
            // 파선/점선/일점쇄선 — Path + LineGeometry + StrokeDashArray. Stretch=Fill 로 가로 폭 채움.
            var dashArray = thb.LineStyle switch
            {
                ThematicLineStyle.Dashed  => new WpfMedia.DoubleCollection { 4, 2 },
                ThematicLineStyle.Dotted  => new WpfMedia.DoubleCollection { 1, 2 },
                ThematicLineStyle.DashDot => new WpfMedia.DoubleCollection { 4, 2, 1, 2 },
                _                         => null,
            };
            var path = new System.Windows.Shapes.Path
            {
                Stroke              = brush,
                StrokeThickness     = thicknessV,
                Stretch             = WpfMedia.Stretch.Fill,
                Data                = new WpfMedia.LineGeometry(new System.Windows.Point(0, 0),
                                                                new System.Windows.Point(1, 0)),
                Height              = Math.Max(thicknessV, 1),
                HorizontalAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment   = VerticalAlignment.Center,
                SnapsToDevicePixels = true,
            };
            if (dashArray is not null) path.StrokeDashArray = dashArray;
            line = path;
        }

        // 비-Solid 경로(Path/StackPanel) — 본문 폭을 채우기 위해 ancestor RichTextBox.Document.PageWidth 에
        // 폭을 바인딩. Build() 에서 PageWidth = ComputeContentWidthDip(page) 로 설정돼 있다.
        var binding = new System.Windows.Data.Binding("Document.PageWidth")
        {
            RelativeSource = new System.Windows.Data.RelativeSource(System.Windows.Data.RelativeSourceMode.FindAncestor)
            {
                AncestorType = typeof(System.Windows.Controls.RichTextBox),
            },
        };
        line.SetBinding(FrameworkElement.WidthProperty, binding);

        return new Wpf.BlockUIContainer(line)
        {
            Margin  = new Thickness(0, marginV, 0, marginV),
            Padding = new Thickness(0),
            // Tag = Core ThematicBreakBlock 인스턴스 — 페이지네이션 / 파서 식별 공유 (위 Solid 케이스 주석 참조).
            Tag     = thb,
        };
    }

    /// <summary>표 캡션 단락 식별용 Tag 센티넬. Parser 가 이 Tag 를 보면 모델에 추가하지 않고 건너뛴다
    /// (Table.Caption 이 다음 렌더에서 다시 생성). 이 마커가 없으면 매 라이브 페이지네이션마다
    /// 캡션 단락이 모델에 누적되어 여러 개로 보인다.</summary>
    internal static readonly object TableCaptionTag = new();

    /// <summary>줄 번호 InlineUIContainer 를 구분하는 센티넬 — 파서가 복사 대상에서 제외한다.</summary>
    internal static readonly object LineNumberTag = new();

    /// <summary>표 캡션을 가운데 정렬 이탤릭 단락으로 빌드한다 (HTML &lt;caption&gt;, DOCX table title).</summary>
    private static Wpf.Paragraph BuildTableCaption(string caption)
        => new Wpf.Paragraph(new Wpf.Run(caption))
        {
            Tag           = TableCaptionTag,
            TextAlignment = TextAlignment.Center,
            FontStyle     = FontStyles.Italic,
            FontSize      = PtToDip(9.5),
            Foreground    = new WpfMedia.SolidColorBrush(WpfMedia.Color.FromRgb(0x55, 0x55, 0x55)),
            Margin        = new Thickness(0, 0, 0, PtToDip(2)),
        };

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

                double padL = MmToDip(cell.PaddingLeftMm   > 0 ? cell.PaddingLeftMm   : Table.FallbackCellPaddingHorizontalMm);
                double padT = MmToDip(cell.PaddingTopMm    > 0 ? cell.PaddingTopMm    : Table.FallbackCellPaddingVerticalMm);
                double padR = MmToDip(cell.PaddingRightMm  > 0 ? cell.PaddingRightMm  : Table.FallbackCellPaddingHorizontalMm);
                double padB = MmToDip(cell.PaddingBottomMm > 0 ? cell.PaddingBottomMm : Table.FallbackCellPaddingVerticalMm);

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

                border.Tag = (r, c);
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
        // 브라우저(Edge/Chrome) 의 깨진-이미지 모양을 흉내낸다: 작은 X 아이콘 + alt 텍스트(Description).
        if (image.Data.Length == 0)
        {
            var icon = new System.Windows.Controls.Border
            {
                Width               = 16,
                Height              = 16,
                BorderBrush         = WpfMedia.Brushes.Gray,
                BorderThickness     = new Thickness(1),
                Background          = WpfMedia.Brushes.WhiteSmoke,
                VerticalAlignment   = VerticalAlignment.Center,
                Margin              = new Thickness(0, 0, 4, 0),
                Child = new System.Windows.Shapes.Path
                {
                    Stroke           = WpfMedia.Brushes.Gray,
                    StrokeThickness  = 1,
                    Data             = WpfMedia.Geometry.Parse("M2,2 L14,14 M14,2 L2,14"),
                },
            };

            var emptyHA = image.HAlign switch
            {
                ImageHAlign.Center => HorizontalAlignment.Center,
                ImageHAlign.Right  => HorizontalAlignment.Right,
                _                  => HorizontalAlignment.Left,
            };
            var stack = new System.Windows.Controls.StackPanel
            {
                Orientation         = System.Windows.Controls.Orientation.Horizontal,
                HorizontalAlignment = emptyHA,
            };
            stack.Children.Add(icon);

            // alt 텍스트가 있을 때만 표시 — 없으면 아이콘만(브라우저 동작과 동일).
            if (!string.IsNullOrWhiteSpace(image.Description))
            {
                stack.Children.Add(new System.Windows.Controls.TextBlock
                {
                    Text              = image.Description,
                    VerticalAlignment = VerticalAlignment.Center,
                });
            }

            // 캡션(`figcaption`) 도 함께 그리도록 WrapImageWithTitle 거쳐 BlockUIContainer 로 포장.
            UIElement emptyVisual = WrapImageWithTitle(stack, image, emptyHA);
            var emptyMargin = new Thickness(0, MmToDip(image.MarginTopMm), 0, MmToDip(image.MarginBottomMm));
            var emptyBuc = new Wpf.BlockUIContainer(emptyVisual) { Tag = image, Margin = emptyMargin };
            if (!string.IsNullOrEmpty(image.Description)) emptyBuc.ToolTip = image.Description;
            return emptyBuc;
        }

        // SVG: WPF BitmapImage 는 image/svg+xml 을 지원하지 않으므로 자체 SvgRenderer 가
        // <rect>/<circle>/<ellipse>/<line>/<polygon>/<polyline>/<path>/<text> + <g> transform 을
        // WPF Canvas + Shape 로 변환해 실제 도형을 그린다. 미지원 요소(marker, gradient 등) 는
        // 무시되며, 라운드트립 원본 SVG 바이트는 image.Data 에 그대로 보존된다.
        if (image.MediaType == "image/svg+xml")
        {
            var svgHa = image.HAlign switch
            {
                ImageHAlign.Center => HorizontalAlignment.Center,
                ImageHAlign.Right  => HorizontalAlignment.Right,
                _                  => HorizontalAlignment.Left,
            };
            var widthDip  = image.WidthMm  > 0 ? MmToDip(image.WidthMm)  : 0;
            var heightDip = image.HeightMm > 0 ? MmToDip(image.HeightMm) : 0;

            UIElement svgVisual;
            var rendered = SvgRenderer.TryRender(image.Data, widthDip, heightDip);
            if (rendered is not null)
            {
                if (rendered is FrameworkElement fe) fe.HorizontalAlignment = svgHa;
                if (!string.IsNullOrEmpty(image.Description) && rendered is FrameworkElement fe2) fe2.ToolTip = image.Description;
                svgVisual = rendered;
            }
            else
            {
                // 파싱 실패 시 기존 placeholder 폴백.
                var svgBox = new System.Windows.Controls.Border
                {
                    Background          = new WpfMedia.SolidColorBrush(WpfMedia.Color.FromRgb(248, 249, 250)),
                    BorderBrush         = new WpfMedia.SolidColorBrush(WpfMedia.Color.FromRgb(221, 221, 221)),
                    BorderThickness     = new Thickness(1),
                    HorizontalAlignment = svgHa,
                };
                if (widthDip  > 0) svgBox.Width  = widthDip;
                if (heightDip > 0) svgBox.Height = heightDip;
                svgBox.Child = new System.Windows.Controls.TextBlock
                {
                    Text                = "[SVG 다이어그램]",
                    Foreground          = WpfMedia.Brushes.Gray,
                    FontStyle           = FontStyles.Italic,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment   = VerticalAlignment.Center,
                    Margin              = new Thickness(8),
                };
                if (!string.IsNullOrEmpty(image.Description)) svgBox.ToolTip = image.Description;
                svgVisual = svgBox;
            }

            var wrappedSvg = WrapImageWithTitle(svgVisual, image, svgHa);
            var svgMargin  = new Thickness(0, MmToDip(image.MarginTopMm), 0, MmToDip(image.MarginBottomMm));
            return new Wpf.BlockUIContainer(wrappedSvg) { Tag = image, Margin = svgMargin };
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

    /// <summary>분리된 이미지 캡션 Paragraph 의 Tag — Parser 가 이 sentinel 을 보면
    /// 캡션 정보는 이미 ImageBlock.ShowTitle/Title 에 보존돼 있으므로 모델에 별도 추가하지 않고 skip 한다.</summary>
    internal sealed class ImageCaptionTag
    {
        public ImageBlock Image { get; }
        public ImageCaptionTag(ImageBlock image) { Image = image; }
    }

    /// <summary>
    /// Inline + Above/Below 위치의 그림 캡션을 별도 Wpf.Paragraph 로 빌드한다.
    /// AppendBlocks 가 ImageBlock 의 BlockUIContainer 와 별개로 이 Paragraph 를 추가해
    /// pagination 이 캡션 텍스트 높이를 정확히 측정·배정할 수 있게 한다(WPF Paragraph 측정은
    /// 오프스크린 RTB 에서도 reliable). overlay/floater/AsText 모드는 같은 시각 단위 안에
    /// 캡션이 묶여야 하므로 이 함수를 사용하지 않고 WrapImageWithTitle 로 처리한다.
    /// Tag = ImageCaptionTag(image) — Parser 가 이 sentinel 을 보고 round-trip 에서 skip(중복 방지).
    /// </summary>
    internal static Wpf.Paragraph? BuildImageCaptionParagraph(ImageBlock image)
    {
        if (!image.ShowTitle || string.IsNullOrWhiteSpace(image.Title)) return null;

        var s = image.TitleStyle;

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

        var run = new Wpf.Run(image.Title)
        {
            FontSize   = PtToDip(s.FontSizePt > 0 ? s.FontSizePt : 10),
            Foreground = titleFg,
            FontWeight = s.Bold   ? FontWeights.Bold   : FontWeights.Normal,
            FontStyle  = s.Italic ? FontStyles.Italic  : FontStyles.Normal,
        };
        if (!string.IsNullOrEmpty(s.FontFamily))
            run.FontFamily = new WpfMedia.FontFamily(s.FontFamily);
        if (titleBg is not null)
            run.Background = titleBg;

        if (s.Underline || s.Strikethrough || s.Overline)
        {
            var decos = new System.Windows.TextDecorationCollection();
            if (s.Underline)     foreach (var d in System.Windows.TextDecorations.Underline)    decos.Add(d);
            if (s.Strikethrough) foreach (var d in System.Windows.TextDecorations.Strikethrough) decos.Add(d);
            if (s.Overline)      foreach (var d in System.Windows.TextDecorations.OverLine)     decos.Add(d);
            run.TextDecorations = decos;
        }

        return new Wpf.Paragraph(run)
        {
            Tag           = new ImageCaptionTag(image),
            TextAlignment = ta,
            Margin        = new Thickness(0, 2, 0, 4),
        };
    }

    /// <summary>
    /// 그림 제목(캡션) 표시가 켜져 있으면 image 시각 요소를 Grid 로 감싸 제목을 함께 배치한다.
    /// 위치(Above/Below/OverlayTop/Middle/Bottom) + 가로 정렬 + X/Y 오프셋(mm)을 적용.
    /// Inline + Above/Below 케이스는 AppendBlocks 가 BuildImageCaptionParagraph 로 분리 처리하므로
    /// 이 함수가 캡션을 그리는 경우는 overlay/floater/AsText 모드에 한정.
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

        bool isOverlay = image.TitlePosition is ImageTitlePosition.OverlayTop
                                              or ImageTitlePosition.OverlayMiddle
                                              or ImageTitlePosition.OverlayBottom;
        if (isOverlay)
        {
            // 같은 셀에 그림과 제목이 겹침 — Grid 단일 셀에 overlap.
            var grid = new System.Windows.Controls.Grid { HorizontalAlignment = imgHA };
            grid.Children.Add(imageVisual);
            tb.VerticalAlignment = image.TitlePosition switch
            {
                ImageTitlePosition.OverlayTop    => VerticalAlignment.Top,
                ImageTitlePosition.OverlayBottom => VerticalAlignment.Bottom,
                _                                => VerticalAlignment.Center,
            };
            tb.HorizontalAlignment = HorizontalAlignment.Stretch;
            grid.Children.Add(tb);
            return grid;
        }
        else
        {
            // 그림 위/아래에 제목을 별도 행으로 배치 — StackPanel 이 FlowDocument 안에서 안정적.
            tb.HorizontalAlignment = HorizontalAlignment.Stretch;
            var sp = new System.Windows.Controls.StackPanel { HorizontalAlignment = imgHA };
            if (image.TitlePosition == ImageTitlePosition.Above)
            {
                sp.Children.Add(tb);
                sp.Children.Add(imageVisual);
            }
            else
            {
                sp.Children.Add(imageVisual);
                sp.Children.Add(tb);
            }
            return sp;
        }
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

        // 인라인 도형의 회전: 회전된 경계 박스 크기의 Grid 로 감싸 레이아웃이 그 공간을 확보하도록 한다.
        // (Canvas 자체의 LayoutTransform 은 explicit Width/Height 와 BlockUIContainer 의
        // measure 와 상호작용이 불안정해 신뢰할 수 없음.)
        FrameworkElement inlineHost = BuildInlineRotationHost(visual, shape.RotationAngleDeg, wDip, hDip, imgHA);

        // ── Inline ────────────────────────────────────────────────────────
        if (shape.WrapMode == ImageWrapMode.Inline)
        {
            return new Wpf.BlockUIContainer(inlineHost)
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
        floater.Blocks.Add(new Wpf.BlockUIContainer(inlineHost));

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
        var canvas  = BuildShapeVisual(shape, wDip, hDip);
        ApplyOverlayShapeRotation(canvas, shape.RotationAngleDeg);
        return canvas;
    }

    private static System.Windows.Controls.Canvas BuildShapeVisual(ShapeObject shape, double wDip, double hDip)
    {
        var canvas = new System.Windows.Controls.Canvas
        {
            Width  = wDip,
            Height = hDip,
        };

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
        WpfMedia.Brush strokeBrushVal = WpfMedia.Brushes.Black;
        if (!string.IsNullOrEmpty(shape.StrokeColor))
        {
            try
            {
                var c = (WpfMedia.Color)WpfMedia.ColorConverter.ConvertFromString(shape.StrokeColor)!;
                strokeBrushVal = new WpfMedia.SolidColorBrush(c);
            }
            catch { }
        }

        double strokeDip = shape.StrokeThicknessPt > 0 ? PtToDip(shape.StrokeThicknessPt) : 0;
        // StrokeThickness=0 일 때 Stroke 를 Transparent 로 설정해 렌더링 경계 계산 오류를 방지한다.
        WpfMedia.Brush effectiveStroke = strokeDip > 0 ? strokeBrushVal : WpfMedia.Brushes.Transparent;

        // 단순 박스 도형(Rectangle·RoundedRect·Ellipse)은 WPF 전용 Shape 컨트롤 사용.
        // Path + RectangleGeometry + Stretch.None 조합은 일부 WPF 버전에서 채우기가 보이지 않는 경우가 있어
        // WPF Shape(System.Windows.Shapes.Rectangle / Ellipse) 으로 대체한다.
        if (shape.Kind is ShapeKind.Rectangle or ShapeKind.RoundedRect)
        {
            double rx = shape.Kind == ShapeKind.RoundedRect
                ? Math.Clamp(MmToDip(shape.CornerRadiusMm), 0, Math.Min(wDip, hDip) / 2.0)
                : 0;
            var rect = new WpfShapes.Rectangle
            {
                Width           = wDip,
                Height          = hDip,
                Fill            = fillBrush,
                Stroke          = effectiveStroke,
                StrokeThickness = strokeDip,
                RadiusX         = rx,
                RadiusY         = rx,
            };
            if (strokeDip > 0)
                rect.StrokeDashArray = BuildDashArray(shape.StrokeDash, strokeDip);
            canvas.Children.Add(rect);
        }
        else if (shape.Kind == ShapeKind.Ellipse)
        {
            var ellipse = new WpfShapes.Ellipse
            {
                Width           = wDip,
                Height          = hDip,
                Fill            = fillBrush,
                Stroke          = effectiveStroke,
                StrokeThickness = strokeDip,
            };
            if (strokeDip > 0)
                ellipse.StrokeDashArray = BuildDashArray(shape.StrokeDash, strokeDip);
            canvas.Children.Add(ellipse);
        }
        else
        {
        var geometry = BuildShapeGeometry(shape, wDip, hDip);
        var path = new WpfShapes.Path
        {
            Data            = geometry,
            Fill            = fillBrush,
            Stroke          = effectiveStroke,
            StrokeThickness = strokeDip,
            StrokeDashArray = BuildDashArray(shape.StrokeDash, strokeDip),
            Stretch         = WpfMedia.Stretch.None,
        };

        canvas.Children.Add(path);
        }

        // 끝모양 (선 계열 — 열린 선에만); path 추가 후 그 위에 그림
        if (shape.Kind is ShapeKind.Line or ShapeKind.Polyline or ShapeKind.Spline)
        {
            var ptsDip = GetPointsDip(shape.Points, wDip, hDip);
            if (ptsDip.Count < 2 && shape.Kind == ShapeKind.Line)
            {
                ptsDip = new List<Point> { new(0, hDip / 2), new(wDip, hDip / 2) };
            }
            AddArrowHeads(canvas, shape.StartArrow, shape.EndArrow, shape.EndShapeSizeMm, ptsDip, strokeBrushVal, strokeDip);
        }

        // 레이블
        if (!string.IsNullOrWhiteSpace(shape.LabelText))
        {
            var label = BuildShapeLabel(shape, wDip, hDip);
            canvas.Children.Add(label);
        }

        return canvas;
    }

    // 오버레이 도형 회전 적용 — 절대 위치 배치이므로 RenderTransform 으로 충분.
    private static void ApplyOverlayShapeRotation(System.Windows.Controls.Canvas canvas, double angleDeg)
    {
        if (Math.Abs(angleDeg) < 0.01) return;
        canvas.RenderTransformOrigin = new Point(0.5, 0.5);
        canvas.RenderTransform = new WpfMedia.RotateTransform(angleDeg);
    }

    // 인라인 도형 회전 호스트 — CSS transform:rotate() 와 동일하게 레이아웃 풋프린트는 원본 크기(wDip×hDip) 유지.
    // 외부 Canvas 의 Width/Height 를 원본 크기로 고정하고 ClipToBounds=false 로 시각적 오버플로를 허용한다.
    // 좁은 표 셀에서 회전 bounding-box 크기로 잘리는 문제를 방지하는 CSS-correct 접근.
    // (BuildShape 에서 회전 오버플로만큼 상하 마진을 보충해 위아래 꼭짓점도 레이아웃 흐름 안에 들어온다.)
    private static FrameworkElement BuildInlineRotationHost(
        System.Windows.Controls.Canvas canvas,
        double angleDeg,
        double wDip,
        double hDip,
        HorizontalAlignment hAlign,
        bool useBboxLayout = false)
    {
        if (Math.Abs(angleDeg) < 0.01) return canvas;

        double outerW, outerH;
        if (useBboxLayout)
        {
            // flex/grid 셀 안에서는 회전 bounding-box 를 레이아웃 크기로 사용해
            // 도형이 시각적으로 셀 영역 안에 담히도록 한다.
            // 내부 캔버스(원본 크기)를 외부 bbox 캔버스 중앙에 배치하고 거기서 회전.
            double rad  = angleDeg * Math.PI / 180.0;
            double cosA = Math.Abs(Math.Cos(rad));
            double sinA = Math.Abs(Math.Sin(rad));
            outerW = cosA * wDip + sinA * hDip;
            outerH = sinA * wDip + cosA * hDip;
            System.Windows.Controls.Canvas.SetLeft(canvas, (outerW - wDip) / 2);
            System.Windows.Controls.Canvas.SetTop(canvas,  (outerH - hDip) / 2);
        }
        else
        {
            // CSS transform:rotate() 는 레이아웃 풋프린트를 변경하지 않는다.
            // 외부 Canvas 는 원본 크기(wDip×hDip)로 유지하고 ClipToBounds=false 로 시각적 오버플로를 허용.
            outerW = wDip;
            outerH = hDip;
            System.Windows.Controls.Canvas.SetLeft(canvas, 0);
            System.Windows.Controls.Canvas.SetTop(canvas, 0);
        }

        canvas.RenderTransformOrigin = new Point(0.5, 0.5);
        canvas.RenderTransform       = new WpfMedia.RotateTransform(angleDeg);

        var outer = new System.Windows.Controls.Canvas
        {
            Width               = outerW,
            Height              = outerH,
            ClipToBounds        = false,
            HorizontalAlignment = hAlign,
        };
        outer.Children.Add(canvas);
        return outer;
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

            case ShapeKind.HalfCircle:
            {
                // 위쪽이 둥근 반원 (CSS border-radius: r r 0 0 패턴 대응).
                // (0, hDip) → 타원호 통해 (wDip, hDip) 까지, 위로 볼록.
                // 회전이 필요하면 ShapeObject.RotationAngleDeg 로 처리한다.
                var fig = new WpfMedia.PathFigure
                {
                    StartPoint = new Point(0, hDip),
                    IsClosed   = true,
                };
                fig.Segments.Add(new WpfMedia.ArcSegment(
                    point:          new Point(wDip, hDip),
                    size:           new Size(wDip / 2.0, hDip),
                    rotationAngle:  0,
                    isLargeArc:     false,
                    sweepDirection: WpfMedia.SweepDirection.Clockwise,
                    isStroked:      true));
                var pg = new WpfMedia.PathGeometry();
                pg.Figures.Add(fig);
                return pg;
            }

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

        if (p.Style.ShowLineNumbers && p.Style.CodeLanguage is not null)
            BuildCodeBlockWithLineNumbers(wpfPara, p.Runs);
        else
            foreach (var run in p.Runs)
                AppendRunInlines(wpfPara, run, fnNums, enNums);

        // 원본 PolyDonky.Paragraph 를 Tag 에 보관 — Parser 가 머지할 때 비-FlowDocument 속성 복원에 사용.
        wpfPara.Tag = p;
        return wpfPara;
    }

    /// HTML &lt;br&gt; 에서 변환된 '\n' 런을 WPF LineBreak 인라인으로 분리·삽입한다.
    /// WPF FlowDocument 의 Run 은 '\n' 문자를 시각적 줄바꿈으로 렌더링하지 않으므로
    /// '\n' 이 포함된 Run 텍스트를 분할해 LineBreak 요소를 사이에 끼워야 한다.
    /// LatexSource·EmojiKey 등 특수 Run 은 BuildInline 에 그대로 위임한다.
    private static void AppendRunInlines(Wpf.Paragraph wpfPara, Run run,
        IReadOnlyDictionary<string, int>? fnNums, IReadOnlyDictionary<string, int>? enNums)
    {
        // 특수 런(수식·이모지·각주 등)이나 '\n' 없는 일반 런은 직접 위임.
        if (run.LatexSource is not null || run.EmojiKey is not null
            || run.FootnoteId is not null || run.EndnoteId is not null
            || run.Field is not null
            || !run.Text.Contains('\n'))
        {
            wpfPara.Inlines.Add(BuildInline(run, fnNums, enNums));
            return;
        }

        // '\n' 분리: 각 조각을 원본 Run 의 스타일로 빌드, 조각 사이에 WPF LineBreak 삽입.
        var parts = run.Text.Split('\n');
        for (int i = 0; i < parts.Length; i++)
        {
            if (i > 0)
                wpfPara.Inlines.Add(new Wpf.LineBreak());
            if (parts[i].Length > 0)
            {
                var sub = new Run { Text = parts[i], Style = run.Style, Url = run.Url };
                wpfPara.Inlines.Add(BuildInline(sub, fnNums, enNums));
            }
        }
    }

    /// <summary>
    /// 코드 블록의 각 줄 앞에 줄 번호를 <see cref="Wpf.InlineUIContainer"/> 로 삽입한다.
    /// InlineUIContainer 는 WPF 텍스트 선택·복사 대상에서 제외되므로 Ctrl+C 시 줄 번호가 빠진다.
    /// </summary>
    private static void BuildCodeBlockWithLineNumbers(Wpf.Paragraph wpfPara, IList<Run> sourceRuns)
    {
        var fullText = string.Concat(sourceRuns.Select(r => r.Text));
        if (fullText.EndsWith('\n')) fullText = fullText[..^1];

        var lines = fullText.Split('\n');
        int maxDigits = lines.Length.ToString(System.Globalization.CultureInfo.InvariantCulture).Length;
        double fontSize = wpfPara.FontSize > 0 ? wpfPara.FontSize : PtToDip(11);
        double numWidth = fontSize * 0.6 * maxDigits + PtToDip(8);

        var numFg     = new WpfMedia.SolidColorBrush(WpfMedia.Color.FromRgb(0x88, 0x88, 0x88));
        var numBorder = new WpfMedia.SolidColorBrush(WpfMedia.Color.FromRgb(0xC0, 0xC0, 0xC0));
        var monoFamily = new WpfMedia.FontFamily("Consolas, D2Coding, monospace");

        // CSS color 속성이 Run.Style.Foreground 에 저장된 경우 그 색을 사용, 없으면 기본값.
        var codeFg = sourceRuns.Count > 0 && sourceRuns[0].Style.Foreground is { } cfgC
            ? new WpfMedia.SolidColorBrush(WpfMedia.Color.FromArgb(cfgC.A, cfgC.R, cfgC.G, cfgC.B))
            : new WpfMedia.SolidColorBrush(WpfMedia.Color.FromRgb(0x1A, 0x1A, 0x1A));

        for (int i = 0; i < lines.Length; i++)
        {
            var numTb = new System.Windows.Controls.TextBlock
            {
                Text          = (i + 1).ToString(System.Globalization.CultureInfo.InvariantCulture)
                                       .PadLeft(maxDigits),
                Width         = numWidth,
                TextAlignment = TextAlignment.Right,
                Foreground    = numFg,
                FontFamily    = monoFamily,
                FontSize      = fontSize,
            };
            var numBorderEl = new System.Windows.Controls.Border
            {
                Child           = numTb,
                BorderBrush     = numBorder,
                BorderThickness = new Thickness(0, 0, 1, 0),
                Padding         = new Thickness(0, 0, 6, 0),
            };
            wpfPara.Inlines.Add(new Wpf.InlineUIContainer(numBorderEl)
            {
                BaselineAlignment = BaselineAlignment.TextTop,
                Tag               = LineNumberTag,
            });

            var lineRun = new Wpf.Run(" " + lines[i])
            {
                FontFamily = monoFamily,
                FontSize   = fontSize,
                Foreground = codeFg,
            };
            if (sourceRuns.Count > 0)
            {
                var rs = sourceRuns[0].Style;
                if (rs.Bold)   lineRun.FontWeight = FontWeights.Bold;
                if (rs.Italic) lineRun.FontStyle  = FontStyles.Italic;
            }
            wpfPara.Inlines.Add(lineRun);

            if (i < lines.Length - 1)
                wpfPara.Inlines.Add(new Wpf.LineBreak());
        }
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
            // OutlineStyle 경계선 + CSS border-bottom 통합.
            bool showTop    = ls.Border.ShowTop;
            bool showBottom = ls.Border.ShowBottom || style.BorderBottomPt > 0;
            if (showTop || showBottom)
            {
                WpfMedia.SolidColorBrush borderBrush;
                // CSS border-bottom 색이 있으면 우선 사용.
                string? borderColorStr = style.BorderBottomPt > 0 ? style.BorderBottomColor : null;
                borderColorStr ??= ls.Border.Color;
                if (!string.IsNullOrEmpty(borderColorStr))
                {
                    try { borderBrush = new WpfMedia.SolidColorBrush((WpfMedia.Color)WpfMedia.ColorConverter.ConvertFromString(borderColorStr)); }
                    catch { borderBrush = WpfMedia.Brushes.DimGray; }
                }
                else
                {
                    borderBrush = WpfMedia.Brushes.DimGray;
                }
                double bottomThick = style.BorderBottomPt > 0 ? PtToDip(style.BorderBottomPt)
                                   : showBottom ? 1.0 : 0.0;
                wpfPara.BorderBrush     = borderBrush;
                wpfPara.BorderThickness = new Thickness(0, showTop ? 1 : 0, 0, bottomThick);
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
            // 사용자 CSS 4면 보더/배경/padding 으로 OutlineStyle 기본값을 덮어쓴다.
            ApplyParagraphBoxStyle(wpfPara, style);
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

        ApplyCodeBlockStyle(wpfPara, style);
        ApplyQuoteLevelStyle(wpfPara, style.QuoteLevel);

        // CSS 4면 보더 + 배경 + 위/아래 padding (ApplyCodeBlockStyle 이 비어 있는 속성에만 기본값을 채우므로
        // 여기서 덮어 쓰지 않아도 되지만, QuoteLevel 기본값 위에서 CSS 가 이길 수 있도록 그대로 유지).
        ApplyParagraphBoxStyle(wpfPara, style);
    }

    private static void ApplyParagraphBoxStyle(Wpf.Paragraph wpfPara, ParagraphStyle s)
    {
        bool anyBorder = s.BorderTopPt    > 0 || s.BorderBottomPt > 0 ||
                         s.BorderLeftPt   > 0 || s.BorderRightPt  > 0;
        if (anyBorder)
        {
            // BorderBrush 는 단일이라 면 색상이 다르면 가장 먼저 명시된 것 채택.
            string? colorStr = s.BorderTopColor ?? s.BorderRightColor ??
                               s.BorderBottomColor ?? s.BorderLeftColor;
            WpfMedia.Brush brush;
            if (!string.IsNullOrEmpty(colorStr))
            {
                try { brush = new WpfMedia.SolidColorBrush((WpfMedia.Color)WpfMedia.ColorConverter.ConvertFromString(colorStr)); }
                catch { brush = WpfMedia.Brushes.DimGray; }
            }
            else brush = WpfMedia.Brushes.DimGray;

            wpfPara.BorderBrush = brush;
            wpfPara.BorderThickness = new Thickness(
                s.BorderLeftPt   > 0 ? PtToDip(s.BorderLeftPt)   : 0,
                s.BorderTopPt    > 0 ? PtToDip(s.BorderTopPt)    : 0,
                s.BorderRightPt  > 0 ? PtToDip(s.BorderRightPt)  : 0,
                s.BorderBottomPt > 0 ? PtToDip(s.BorderBottomPt) : 0);
        }

        // 배경색.
        if (!string.IsNullOrEmpty(s.BackgroundColor))
        {
            try
            {
                var bgBrush = new WpfMedia.SolidColorBrush(
                    (WpfMedia.Color)WpfMedia.ColorConverter.ConvertFromString(s.BackgroundColor));
                wpfPara.Background = bgBrush;
            }
            catch { /* 잘못된 색상 — 무시 */ }
        }

        // 위/아래 padding (좌우 padding 은 IndentLeft/RightMm 가 wpfPara.Padding 로 별도 설정됨).
        if (s.PaddingTopMm > 0 || s.PaddingBottomMm > 0)
        {
            var pad = wpfPara.Padding;
            wpfPara.Padding = new Thickness(
                pad.Left,
                s.PaddingTopMm    > 0 ? MmToDip(s.PaddingTopMm)    : pad.Top,
                pad.Right,
                s.PaddingBottomMm > 0 ? MmToDip(s.PaddingBottomMm) : pad.Bottom);
        }
    }

    /// <summary>
    /// 직전 블록과 현재 블록이 같은 QuoteLevel(>0) 의 단락이면 두 단락 사이의
    /// 위/아래 마진을 0 으로 만들어 좌측 인용 바가 끊기지 않게 이어 붙인다.
    /// </summary>
    private static void MergeAdjacentBlockquoteMargins(System.Collections.IList target)
    {
        if (target.Count < 2) return;
        if (target[target.Count - 1] is not Wpf.Paragraph cur) return;
        if (target[target.Count - 2] is not Wpf.Paragraph prev) return;
        if (cur.Tag is not Paragraph curP || prev.Tag is not Paragraph prevP) return;
        if (curP.Style.QuoteLevel <= 0) return;
        if (curP.Style.QuoteLevel != prevP.Style.QuoteLevel) return;

        var pm = prev.Margin;
        prev.Margin = new Thickness(pm.Left, pm.Top, pm.Right, 0);
        var cm = cur.Margin;
        cur.Margin  = new Thickness(cm.Left, 0, cm.Right, cm.Bottom);
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
    /// CodeLanguage != null(= pre/code 블록)이면 모노스페이스 + 기본 박스 스타일을 적용.
    /// CSS 에 이미 값이 있는 속성(배경·보더·Foreground)은 건드리지 않는다 — CSS 가 항상 우선.
    /// </summary>
    private static void ApplyCodeBlockStyle(Wpf.Paragraph wpfPara, ParagraphStyle style)
    {
        if (style.CodeLanguage is null) return;
        wpfPara.FontFamily = new WpfMedia.FontFamily("Consolas, D2Coding, monospace");

        // CSS background-color 가 없을 때만 기본 밝은 회색 배경 적용.
        if (string.IsNullOrEmpty(style.BackgroundColor))
            wpfPara.Background = new WpfMedia.SolidColorBrush(WpfMedia.Color.FromRgb(0xF8, 0xF8, 0xF8));

        // Foreground 는 절대 하드코딩하지 않는다 — CSS color 는 Run 레벨에 이미 반영되어 있고,
        // 단락 레벨 기본값을 강제하면 테마·CSS 색상 모두 덮어써 버린다.

        // CSS border 가 없을 때만 기본 회색 테두리 적용.
        bool hasCssBorder = style.BorderTopPt > 0 || style.BorderBottomPt > 0 ||
                            style.BorderLeftPt > 0 || style.BorderRightPt  > 0;
        if (!hasCssBorder)
        {
            wpfPara.BorderBrush     = new WpfMedia.SolidColorBrush(WpfMedia.Color.FromRgb(0xD0, 0xD0, 0xD0));
            wpfPara.BorderThickness = new Thickness(1);
        }

        wpfPara.Padding = new Thickness(MmToDip(3.0), MmToDip(1.5), MmToDip(3.0), MmToDip(1.5));
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

        if (run.RubyText is { Length: > 0 })
            return BuildRubyInline(run);

        var s = run.Style;
        if (NeedsContainer(s))
            return BuildScaledContainer(run);

        var displayText = ApplyTextTransform(run.Text, s.TextTransform);
        var wpfRun = new Wpf.Run(displayText);

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

        if (s.FontVariantSmallCaps)
            Wpf.Typography.SetCapitals(wpfRun, System.Windows.FontCapitals.SmallCaps);

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

    private static Wpf.Inline BuildRubyInline(Run run)
    {
        var s        = run.Style;
        double baseFontSize = PtToDip(s.FontSizePt > 0 ? s.FontSizePt : 11);
        double rubyFontSize = baseFontSize * 0.55;

        var baseText = ApplyTextTransform(run.Text, s.TextTransform);
        var rubyText = run.RubyText!;

        var baseTb = new System.Windows.Controls.TextBlock
        {
            Text          = baseText,
            FontSize      = baseFontSize,
            TextAlignment = TextAlignment.Center,
        };
        ApplyStyleToTextBlock(baseTb, s);

        var rubyTb = new System.Windows.Controls.TextBlock
        {
            Text          = rubyText,
            FontSize      = rubyFontSize,
            TextAlignment = TextAlignment.Center,
            Foreground    = s.Foreground is { } rfg
                ? new WpfMedia.SolidColorBrush(WpfMedia.Color.FromArgb(rfg.A, rfg.R, rfg.G, rfg.B))
                : System.Windows.SystemColors.ControlTextBrush,
        };
        if (!string.IsNullOrEmpty(s.FontFamily))
            rubyTb.FontFamily = new WpfMedia.FontFamily(s.FontFamily);

        var panel = new System.Windows.Controls.StackPanel
        {
            Orientation = System.Windows.Controls.Orientation.Vertical,
        };
        panel.Children.Add(rubyTb);
        panel.Children.Add(baseTb);

        return new Wpf.InlineUIContainer(panel)
        {
            Tag               = run,
            BaselineAlignment = BaselineAlignment.Baseline,
        };
    }

    private static string ApplyTextTransform(string text, TextTransform transform) => transform switch
    {
        TextTransform.Uppercase  => text.ToUpperInvariant(),
        TextTransform.Lowercase  => text.ToLowerInvariant(),
        TextTransform.Capitalize => System.Text.RegularExpressions.Regex.Replace(
            text, @"(?:^|\s)\S", m => m.Value.ToUpperInvariant()),
        _ => text,
    };

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

        var text = ApplyTextTransform(run.Text.Length > 0 ? run.Text : " ", s.TextTransform);
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
        if (s.FontVariantSmallCaps)
            Wpf.Typography.SetCapitals(tb, System.Windows.FontCapitals.SmallCaps);
    }
}

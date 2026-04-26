using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using PolyDoc.App.Services;
using PolyDoc.Core;
using Wpf = System.Windows.Documents;

namespace PolyDoc.App.Views;

/// <summary>
/// 문단 서식 다이얼로그.
/// <list type="bullet">
/// <item>RichTextBox 모드: 선택 영역 단락의 서식을 읽어 초기화, OK 시 일괄 적용.</item>
/// <item>단독 모드(ParagraphStyle 생성자): 에디터 없이 스타일 객체만 편집. ResultStyle 로 결과 반환.</item>
/// </list>
/// </summary>
public partial class ParaFormatWindow : Window
{
    private readonly RichTextBox? _editor;
    private readonly ParagraphStyle? _standaloneStyle;
    private bool _suppressPreview;

    /// <summary>RichTextBox 선택 영역 편집 모드.</summary>
    public ParaFormatWindow(RichTextBox editor)
    {
        _editor = editor;
        InitializeComponent();
        PopulateOutlineLevels();
        Loaded += (_, _) => LoadCurrentFormatting();
    }

    /// <summary>단독 모드 — 에디터 없이 ParagraphStyle 직접 편집. OK 후 ResultStyle 로 결과 읽기.</summary>
    public ParaFormatWindow(ParagraphStyle initial)
    {
        _standaloneStyle = CloneParagraphStyle(initial);
        InitializeComponent();
        PopulateOutlineLevels();
        Loaded += (_, _) => LoadCurrentFormatting();
    }

    /// <summary>단독 모드에서 OK 후 편집 결과를 담는다.</summary>
    public ParagraphStyle? ResultStyle { get; private set; }

    private void PopulateOutlineLevels()
    {
        var loc = LocalizedStrings.Instance;
        // 본문 + H1~H6
        CboOutline.Items.Add(new ComboBoxItem { Content = loc.FormatParaOutlineBody, Tag = OutlineLevel.Body });
        for (int i = 1; i <= 6; i++)
        {
            CboOutline.Items.Add(new ComboBoxItem { Content = $"H{i}", Tag = (OutlineLevel)i });
        }
    }

    // ── 현재 서식 읽기 ─────────────────────────────────────────

    private void LoadCurrentFormatting()
    {
        _suppressPreview = true;
        try
        {
            ParagraphStyle style;
            if (_standaloneStyle is { } s)
            {
                // 단독 모드: ParagraphStyle 에서 직접 읽음
                style = s;
                switch (s.Alignment)
                {
                    case Alignment.Center:      RbAlignCenter.IsChecked = true; break;
                    case Alignment.Right:       RbAlignRight.IsChecked = true; break;
                    case Alignment.Justify:     RbAlignJustify.IsChecked = true; break;
                    case Alignment.Distributed: RbAlignDistributed.IsChecked = true; break;
                    default:                    RbAlignLeft.IsChecked = true; break;
                }
            }
            else
            {
                // RichTextBox 모드
                var paragraph = GetFirstParagraph();
                var alignVal = _editor!.Selection.GetPropertyValue(Wpf.Paragraph.TextAlignmentProperty);
                var align = alignVal is TextAlignment ta ? ta : TextAlignment.Left;
                switch (align)
                {
                    case TextAlignment.Center:  RbAlignCenter.IsChecked = true; break;
                    case TextAlignment.Right:   RbAlignRight.IsChecked = true; break;
                    case TextAlignment.Justify: RbAlignJustify.IsChecked = true; break;
                    default:                    RbAlignLeft.IsChecked = true; break;
                }
                if (paragraph?.Tag is PolyDoc.Core.Paragraph polyP && polyP.Style.Alignment == Alignment.Distributed)
                    RbAlignDistributed.IsChecked = true;
                style = (paragraph?.Tag as PolyDoc.Core.Paragraph)?.Style ?? new ParagraphStyle();
            }

            TxtLineHeight.Text   = (style.LineHeightFactor * 100).ToString("0");
            TxtSpaceBefore.Text  = style.SpaceBeforePt.ToString("0.#");
            TxtSpaceAfter.Text   = style.SpaceAfterPt.ToString("0.#");
            TxtIndentFirst.Text  = style.IndentFirstLineMm.ToString("0.#");
            TxtIndentLeft.Text   = style.IndentLeftMm.ToString("0.#");
            TxtIndentRight.Text  = style.IndentRightMm.ToString("0.#");

            var outline = style.Outline;
            for (int i = 0; i < CboOutline.Items.Count; i++)
            {
                if (CboOutline.Items[i] is ComboBoxItem ci && ci.Tag is OutlineLevel ol && ol == outline)
                {
                    CboOutline.SelectedIndex = i;
                    break;
                }
            }
            if (CboOutline.SelectedIndex < 0) CboOutline.SelectedIndex = 0;
        }
        finally
        {
            _suppressPreview = false;
        }

        UpdatePreview();
    }

    private Wpf.Paragraph? GetFirstParagraph()
    {
        var sel = _editor!.Selection;
        return sel.Start.Paragraph ?? sel.End.Paragraph;
    }

    // ── 미리보기 ───────────────────────────────────────────────

    private void OnPreviewUpdate(object sender, EventArgs e) => UpdatePreview();

    private void UpdatePreview()
    {
        if (_suppressPreview) return;

        var fd = new FlowDocument
        {
            FontFamily = new System.Windows.Media.FontFamily("맑은 고딕, Malgun Gothic, Segoe UI"),
            FontSize   = FlowDocumentBuilder.PtToDip(11),
            PagePadding = new Thickness(0),
        };

        var p = new Wpf.Paragraph();
        p.Inlines.Add(new Wpf.Run("가나다라 ABCD 1234. 미리보기 텍스트입니다. 정렬과 들여쓰기, 줄 간격이 표시됩니다."));
        p.Inlines.Add(new LineBreak());
        p.Inlines.Add(new Wpf.Run("두 번째 줄. 줄 간격이 적용된 모습을 확인하세요."));

        ApplyToWpfParagraph(p, includeMargin: true);
        fd.Blocks.Add(p);

        PreviewViewer.Document = fd;
    }

    private void ApplyToWpfParagraph(Wpf.Paragraph p, bool includeMargin)
    {
        // 정렬
        if (RbAlignCenter.IsChecked == true)        p.TextAlignment = TextAlignment.Center;
        else if (RbAlignRight.IsChecked == true)    p.TextAlignment = TextAlignment.Right;
        else if (RbAlignJustify.IsChecked == true || RbAlignDistributed.IsChecked == true)
            p.TextAlignment = TextAlignment.Justify;
        else                                        p.TextAlignment = TextAlignment.Left;

        // 줄 간격 (% → factor → DIP)
        if (double.TryParse(TxtLineHeight.Text, out var lh) && lh >= 50)
        {
            var fontSize = p.FontSize > 0 ? p.FontSize : FlowDocumentBuilder.PtToDip(11);
            p.LineHeight = fontSize * (lh / 100.0);
        }

        // 단락 위/아래 + 좌/우 — Margin 한꺼번에
        if (includeMargin)
        {
            double top    = double.TryParse(TxtSpaceBefore.Text, out var sb) && sb >= 0 ? FlowDocumentBuilder.PtToDip(sb) : 0;
            double bottom = double.TryParse(TxtSpaceAfter.Text,  out var sa) && sa >= 0 ? FlowDocumentBuilder.PtToDip(sa) : 0;
            double left   = double.TryParse(TxtIndentLeft.Text,  out var il) && il >= 0 ? FlowDocumentBuilder.MmToDip(il) : 0;
            double right  = double.TryParse(TxtIndentRight.Text, out var ir) && ir >= 0 ? FlowDocumentBuilder.MmToDip(ir) : 0;
            p.Margin = new Thickness(left, top, right, bottom);
        }

        // 첫 줄 들여쓰기
        if (double.TryParse(TxtIndentFirst.Text, out var fl))
            p.TextIndent = FlowDocumentBuilder.MmToDip(fl);
    }

    // ── OK / Cancel ────────────────────────────────────────────

    private void OnOk(object sender, RoutedEventArgs e)
    {
        if (_standaloneStyle != null)
        {
            WriteToParagraphStyle(_standaloneStyle);
            ResultStyle = CloneParagraphStyle(_standaloneStyle);
        }
        else
        {
            ApplyToSelection();
        }
        DialogResult = true;
    }

    private void OnCancel(object sender, RoutedEventArgs e) => DialogResult = false;

    // ── UI → ParagraphStyle 쓰기 (단독 모드용) ──────────────────

    private void WriteToParagraphStyle(ParagraphStyle s)
    {
        if (RbAlignCenter.IsChecked == true)        s.Alignment = Alignment.Center;
        else if (RbAlignRight.IsChecked == true)    s.Alignment = Alignment.Right;
        else if (RbAlignJustify.IsChecked == true)  s.Alignment = Alignment.Justify;
        else if (RbAlignDistributed.IsChecked == true) s.Alignment = Alignment.Distributed;
        else                                        s.Alignment = Alignment.Left;

        if (double.TryParse(TxtLineHeight.Text, out var lh) && lh >= 50) s.LineHeightFactor = lh / 100.0;
        if (double.TryParse(TxtSpaceBefore.Text, out var sb2) && sb2 >= 0) s.SpaceBeforePt = sb2;
        if (double.TryParse(TxtSpaceAfter.Text,  out var sa)  && sa  >= 0) s.SpaceAfterPt  = sa;
        if (double.TryParse(TxtIndentFirst.Text, out var fl)) s.IndentFirstLineMm = fl;
        if (double.TryParse(TxtIndentLeft.Text,  out var il) && il >= 0) s.IndentLeftMm  = il;
        if (double.TryParse(TxtIndentRight.Text, out var ir) && ir >= 0) s.IndentRightMm = ir;
        if (CboOutline.SelectedItem is ComboBoxItem ci && ci.Tag is OutlineLevel ol) s.Outline = ol;
    }

    private static ParagraphStyle CloneParagraphStyle(ParagraphStyle s) => new()
    {
        Alignment        = s.Alignment,
        LineHeightFactor = s.LineHeightFactor,
        SpaceBeforePt    = s.SpaceBeforePt,
        SpaceAfterPt     = s.SpaceAfterPt,
        IndentFirstLineMm = s.IndentFirstLineMm,
        IndentLeftMm     = s.IndentLeftMm,
        IndentRightMm    = s.IndentRightMm,
        Outline          = s.Outline,
        ListMarker       = s.ListMarker,
    };

    // ── 선택 영역 단락에 적용 (RichTextBox 모드) ────────────────

    private void ApplyToSelection()
    {
        var sel = _editor!.Selection;
        var paragraphs = CollectParagraphs(sel);

        // 사용자 입력 파싱
        Alignment polyAlign;
        if (RbAlignCenter.IsChecked == true)        polyAlign = Alignment.Center;
        else if (RbAlignRight.IsChecked == true)    polyAlign = Alignment.Right;
        else if (RbAlignJustify.IsChecked == true)  polyAlign = Alignment.Justify;
        else if (RbAlignDistributed.IsChecked == true) polyAlign = Alignment.Distributed;
        else                                        polyAlign = Alignment.Left;

        double lineHeightFactor = double.TryParse(TxtLineHeight.Text, out var lh) && lh >= 50 ? lh / 100.0 : 1.2;
        double spaceBefore      = double.TryParse(TxtSpaceBefore.Text, out var sb) && sb >= 0 ? sb : 0;
        double spaceAfter       = double.TryParse(TxtSpaceAfter.Text,  out var sa) && sa >= 0 ? sa : 0;
        double indentFirst      = double.TryParse(TxtIndentFirst.Text, out var fl) ? fl : 0;
        double indentLeft       = double.TryParse(TxtIndentLeft.Text,  out var il) && il >= 0 ? il : 0;
        double indentRight      = double.TryParse(TxtIndentRight.Text, out var ir) && ir >= 0 ? ir : 0;

        OutlineLevel outline = OutlineLevel.Body;
        if (CboOutline.SelectedItem is ComboBoxItem ci && ci.Tag is OutlineLevel ol)
            outline = ol;

        foreach (var wpfPara in paragraphs)
        {
            // WPF 시각 속성
            ApplyToWpfParagraph(wpfPara, includeMargin: true);

            // Tag(PolyDoc.Paragraph) 도 갱신 — 라운드트립 보존
            if (wpfPara.Tag is PolyDoc.Core.Paragraph polyPara)
            {
                polyPara.Style.Alignment         = polyAlign;
                polyPara.Style.LineHeightFactor  = lineHeightFactor;
                polyPara.Style.SpaceBeforePt     = spaceBefore;
                polyPara.Style.SpaceAfterPt      = spaceAfter;
                polyPara.Style.IndentFirstLineMm = indentFirst;
                polyPara.Style.IndentLeftMm      = indentLeft;
                polyPara.Style.IndentRightMm    = indentRight;
                polyPara.Style.Outline           = outline;
            }
        }
    }

    // ── 선택 영역 단락 수집 ─────────────────────────────────────

    private static List<Wpf.Paragraph> CollectParagraphs(TextSelection sel)
    {
        var result = new List<Wpf.Paragraph>();
        var seen = new HashSet<Wpf.Paragraph>();

        var startPara = sel.Start.Paragraph;
        if (startPara != null && seen.Add(startPara)) result.Add(startPara);

        var ptr = sel.Start;
        while (ptr != null && ptr.CompareTo(sel.End) < 0)
        {
            var para = ptr.Paragraph;
            if (para != null && seen.Add(para)) result.Add(para);
            ptr = ptr.GetNextContextPosition(LogicalDirection.Forward);
        }

        var endPara = sel.End.Paragraph;
        if (endPara != null && seen.Add(endPara)) result.Add(endPara);

        return result;
    }
}

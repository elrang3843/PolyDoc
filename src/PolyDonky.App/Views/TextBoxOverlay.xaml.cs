using System;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Media3D;
using System.Windows.Shapes;
using PolyDonky.Core;
using WpfMedia = System.Windows.Media;

namespace PolyDonky.App.Views;

/// <summary>
/// 글상자(<see cref="TextBoxObject"/>) 의 시각 + 인터랙션 컨테이너.
///
/// - 부모는 <c>FloatingCanvas</c> (MainWindow). Canvas.Left/Top + Width/Height 로 위치/크기 결정.
/// - 모델은 mm 단위, 캔버스는 DIP 단위 — <see cref="DipsPerMm"/> 로 변환.
/// - 외곽(Border/Path) 클릭 = 드래그 이동, 내부 RichTextBox 클릭 = 서식 편집.
/// - 선택 시 4코너 핸들로 리사이즈.
/// - 우클릭 컨텍스트 메뉴: 속성, 앞/뒤 순서, 삭제.
/// </summary>
public partial class TextBoxOverlay : UserControl
{
    public const double DipsPerMm = 96.0 / 25.4;

    // ── PathGeometry 생성기 (100×100 정규화 공간, Stretch=Fill 로 자동 스케일) ──

    private static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

    /// <summary>
    /// 말풍선 PathGeometry 생성. 본체는 둥근 사각형, 꼬리는 <paramref name="dir"/>
    /// 방향으로 100×100 박스 변/모서리에 닿도록 그린다.
    /// 꼬리 쪽 변에는 15 단위 여백을 두어 본체 영역을 줄인다.
    /// </summary>
    private static string BuildSpeechPath(SpeechPointerDirection dir)
    {
        const double m = 15;   // 꼬리 여백
        const double r = 6;    // 본체 모서리 반경
        const double tw = 9;   // 직선변 꼬리 밑변 절반 너비
        const double offs = 3; // 꼬리 끝점 살짝 비대칭 (자연스럽게)

        double bL = 0, bT = 0, bR = 100, bB = 100;
        bool tL = dir is SpeechPointerDirection.Left  or SpeechPointerDirection.TopLeft  or SpeechPointerDirection.BottomLeft;
        bool tR = dir is SpeechPointerDirection.Right or SpeechPointerDirection.TopRight or SpeechPointerDirection.BottomRight;
        bool tT = dir is SpeechPointerDirection.Top   or SpeechPointerDirection.TopLeft  or SpeechPointerDirection.TopRight;
        bool tB = dir is SpeechPointerDirection.Bottom or SpeechPointerDirection.BottomLeft or SpeechPointerDirection.BottomRight;
        if (tL) bL = m;
        if (tR) bR = 100 - m;
        if (tT) bT = m;
        if (tB) bB = 100 - m;

        var sb = new StringBuilder();
        // 시계방향: 좌상 모서리(반경 시작점) → 상변 → 우상 → 우변 → 우하 → 하변 → 좌하 → 좌변 → 좌상.
        sb.AppendFormat(Inv, "M {0:0.##},{1:0.##} ", bL + r, bT);

        // ── 상변 (Top 꼬리 분기) ──
        if (dir == SpeechPointerDirection.Top)
        {
            double cx = (bL + bR) / 2;
            sb.AppendFormat(Inv, "L {0:0.##},{1:0.##} L {2:0.##},0 L {3:0.##},{1:0.##} ",
                cx - tw, bT, cx + offs, cx + tw);
        }
        sb.AppendFormat(Inv, "L {0:0.##},{1:0.##} ", bR - r, bT);

        // ── 우상 모서리 (TopRight 꼬리 분기) ──
        if (dir == SpeechPointerDirection.TopRight)
            sb.AppendFormat(Inv, "L 100,0 L {0:0.##},{1:0.##} ", bR, bT + r);
        else
            sb.AppendFormat(Inv, "Q {0:0.##},{1:0.##} {0:0.##},{2:0.##} ", bR, bT, bT + r);

        // ── 우변 (Right 꼬리 분기) ──
        if (dir == SpeechPointerDirection.Right)
        {
            double cy = (bT + bB) / 2;
            sb.AppendFormat(Inv, "L {0:0.##},{1:0.##} L 100,{2:0.##} L {0:0.##},{3:0.##} ",
                bR, cy - tw, cy + offs, cy + tw);
        }
        sb.AppendFormat(Inv, "L {0:0.##},{1:0.##} ", bR, bB - r);

        // ── 우하 모서리 ──
        if (dir == SpeechPointerDirection.BottomRight)
            sb.AppendFormat(Inv, "L 100,100 L {0:0.##},{1:0.##} ", bR - r, bB);
        else
            sb.AppendFormat(Inv, "Q {0:0.##},{1:0.##} {2:0.##},{1:0.##} ", bR, bB, bR - r);

        // ── 하변 (Bottom 꼬리 분기) ──
        if (dir == SpeechPointerDirection.Bottom)
        {
            double cx = (bL + bR) / 2;
            sb.AppendFormat(Inv, "L {0:0.##},{1:0.##} L {2:0.##},100 L {3:0.##},{1:0.##} ",
                cx + tw, bB, cx - offs, cx - tw);
        }
        sb.AppendFormat(Inv, "L {0:0.##},{1:0.##} ", bL + r, bB);

        // ── 좌하 모서리 ──
        if (dir == SpeechPointerDirection.BottomLeft)
            sb.AppendFormat(Inv, "L 0,100 L {0:0.##},{1:0.##} ", bL, bB - r);
        else
            sb.AppendFormat(Inv, "Q {0:0.##},{1:0.##} {0:0.##},{2:0.##} ", bL, bB, bB - r);

        // ── 좌변 (Left 꼬리 분기) ──
        if (dir == SpeechPointerDirection.Left)
        {
            double cy = (bT + bB) / 2;
            sb.AppendFormat(Inv, "L {0:0.##},{1:0.##} L 0,{2:0.##} L {0:0.##},{3:0.##} ",
                bL, cy + tw, cy - offs, cy - tw);
        }
        sb.AppendFormat(Inv, "L {0:0.##},{1:0.##} ", bL, bT + r);

        // ── 좌상 모서리 ──
        if (dir == SpeechPointerDirection.TopLeft)
            sb.AppendFormat(Inv, "L 0,0 L {0:0.##},{1:0.##} ", bL + r, bT);
        else
            sb.AppendFormat(Inv, "Q {0:0.##},{1:0.##} {2:0.##},{1:0.##} ", bL, bT, bL + r);

        sb.Append('Z');
        return sb.ToString();
    }

    /// <summary>
    /// 구름풍선 PathGeometry. 둘레를 따라 N개의 볼록한 호(quadratic Bezier)를 이어 그린다.
    /// 안쪽 타원 위에 N개 기준점을 배치하고, 인접 기준점 사이를 외곽 타원의 점을 control 로 한 호로 잇는다.
    /// </summary>
    private static string BuildCloudPath(int puffCount)
    {
        int n = Math.Clamp(puffCount, 6, 32);
        const double cx = 50, cy = 50;
        const double rxIn = 38, ryIn = 32;
        const double rxOut = 56, ryOut = 56;
        const double startA = -Math.PI / 2;

        var basePts = new (double X, double Y)[n];
        for (int i = 0; i < n; i++)
        {
            double a = startA + 2 * Math.PI * i / n;
            basePts[i] = (cx + rxIn * Math.Cos(a), cy + ryIn * Math.Sin(a));
        }

        var sb = new StringBuilder();
        sb.AppendFormat(Inv, "M {0:0.##},{1:0.##} ", basePts[0].X, basePts[0].Y);
        for (int i = 0; i < n; i++)
        {
            int j = (i + 1) % n;
            double a = startA + 2 * Math.PI * (i + 0.5) / n;
            double mx = cx + rxOut * Math.Cos(a);
            double my = cy + ryOut * Math.Sin(a);
            sb.AppendFormat(Inv, "Q {0:0.##},{1:0.##} {2:0.##},{3:0.##} ",
                mx, my, basePts[j].X, basePts[j].Y);
        }
        sb.Append('Z');
        return sb.ToString();
    }

    /// <summary>
    /// 가시풍선 PathGeometry. N각 별 — 외곽 정점(50→0) 과 안쪽 정점(rIn) 을 교대로 잇는다.
    /// </summary>
    private static string BuildSpikyPath(int spikeCount)
    {
        int n = Math.Clamp(spikeCount, 5, 24);
        const double cx = 50, cy = 50;
        const double rOut = 50;
        const double rIn  = 32;
        const double startA = -Math.PI / 2;

        var sb = new StringBuilder();
        int total = n * 2;
        for (int i = 0; i < total; i++)
        {
            double a = startA + Math.PI * i / n;
            double r = (i % 2 == 0) ? rOut : rIn;
            double x = cx + r * Math.Cos(a);
            double y = cy + r * Math.Sin(a);
            sb.AppendFormat(Inv, "{0} {1:0.##},{2:0.##} ", i == 0 ? "M" : "L", x, y);
        }
        sb.Append('Z');
        return sb.ToString();
    }

    /// <summary>
    /// 타원 PathGeometry. 4개의 cubic Bézier 호로 원을 근사하며,
    /// Stretch="Fill" 로 인해 실제 박스 비율에 맞춰 타원으로 자동 스케일된다.
    /// </summary>
    private static string BuildEllipsePath()
        => "M 50,0 C 77.6,0 100,22.4 100,50 C 100,77.6 77.6,100 50,100 C 22.4,100 0,77.6 0,50 C 0,22.4 22.4,0 50,0 Z";

    /// <summary>
    /// 파이(부채꼴) PathGeometry. 중심(50,50)에서 반경 50의 호를 그린 후 다시 중심으로 닫는다.
    /// <paramref name="startAngleDeg"/>: 시계방향 시작 각도 (0 = 오른쪽).
    /// <paramref name="sweepAngleDeg"/>: 시계방향 호 범위 (5~355).
    /// </summary>
    private static string BuildPiePath(double startAngleDeg, double sweepAngleDeg)
    {
        const double cx = 50, cy = 50, r = 50;
        double sweep = Math.Clamp(sweepAngleDeg, 5, 355);
        double startRad = startAngleDeg * Math.PI / 180.0;
        double endRad   = (startAngleDeg + sweep) * Math.PI / 180.0;
        double x1 = cx + r * Math.Cos(startRad);
        double y1 = cy + r * Math.Sin(startRad);
        double x2 = cx + r * Math.Cos(endRad);
        double y2 = cy + r * Math.Sin(endRad);
        int largeArc = sweep > 180 ? 1 : 0;
        return string.Format(Inv,
            "M {0:0.##},{1:0.##} L {2:0.##},{3:0.##} A {4},{4} 0 {5} 1 {6:0.##},{7:0.##} Z",
            cx, cy, x1, y1, r, largeArc, x2, y2);
    }

    /// <summary>
    /// 번개상자 PathGeometry. 꺽임 개수에 따라 미리 정의된 템플릿을 선택.
    /// 1=단순(작은 V형), 2=기본 볼트(원래 모양), 3~5=촘촘한 지그재그.
    /// </summary>
    private static string BuildLightningPath(int bendCount)
    {
        int b = Math.Clamp(bendCount, 1, 5);
        return b switch
        {
            1 => "M 60,0 L 32,60 L 52,60 L 40,100 L 70,42 L 50,42 Z",
            2 => "M 65,0 L 22,52 L 46,52 L 35,100 L 78,48 L 54,48 Z",
            3 => "M 70,0 L 30,32 L 50,32 L 22,64 L 42,64 L 32,100 L 76,58 L 56,58 L 78,28 L 60,28 Z",
            4 => "M 72,0 L 32,24 L 50,24 L 22,48 L 42,48 L 26,72 L 46,72 L 35,100 L 76,68 L 56,68 L 78,44 L 58,44 L 78,18 L 60,18 Z",
            5 => "M 72,0 L 32,18 L 48,18 L 22,38 L 40,38 L 18,56 L 38,56 L 22,76 L 42,76 L 32,100 L 76,72 L 56,72 L 78,54 L 58,54 L 78,34 L 58,34 L 78,14 L 60,14 Z",
            _ => "M 65,0 L 22,52 L 46,52 L 35,100 L 78,48 L 54,48 Z",
        };
    }

    public TextBoxObject Model { get; }

    /// <summary>이 오버레이가 선택됨 — 호출자가 다른 오버레이 선택 해제할 때 사용.</summary>
    public event EventHandler? Selected;

    /// <summary>드래그/리사이즈 종료 — 호출자가 모델 mm 좌표 갱신 + Dirty 표시.</summary>
    public event EventHandler? GeometryChangedCommitted;

    /// <summary>본문 편집 종료 — 호출자가 모델 동기화 + Dirty 표시.</summary>
    public event EventHandler? ContentChangedCommitted;

    /// <summary>오버레이 삭제 요청 (Delete 키 / 컨텍스트 메뉴). 호출자가 모델/캔버스에서 제거.</summary>
    public event EventHandler? DeleteRequested;

    /// <summary>앞으로 가져오기 요청. 호출자가 Canvas ZOrder 조정.</summary>
    public event EventHandler? BringForwardRequested;

    /// <summary>뒤로 보내기 요청. 호출자가 Canvas ZOrder 조정.</summary>
    public event EventHandler? SendBackRequested;

    /// <summary>속성 변경 확정 (속성 대화상자 OK). 호출자가 Dirty 표시.</summary>
    public event EventHandler? AppearanceChangedCommitted;

    private bool _suppressTextChanged;

    public TextBoxOverlay(TextBoxObject model)
    {
        Model = model ?? throw new ArgumentNullException(nameof(model));
        InitializeComponent();

        // 초기 텍스트 로드 (plain text → FlowDocument)
        _suppressTextChanged = true;
        LoadModelTextToEditor();
        _suppressTextChanged = false;

        ApplyShapeFromModel();
        Loaded += (_, _) =>
        {
            UpdateHandlePositions();
            ApplySpellCheckLanguage();
        };
        Services.LanguageService.LanguageChanged += (_, _) => ApplySpellCheckLanguage();
        SizeChanged += (_, _) =>
        {
            UpdateHandlePositions();
            // 비사각형 모양은 인셋이 박스 크기에 비례하므로 크기 변경 시 재계산.
            if (Model.Shape != TextBoxShape.Rectangle)
                UpdateInnerEditorMargin();
        };
    }

    // ── 선택 상태 ────────────────────────────────────────────────────

    private bool _isSelected;
    public bool IsSelected
    {
        get => _isSelected;
        set
        {
            if (_isSelected == value) return;
            _isSelected = value;
            SelectionFrame.Visibility = value ? Visibility.Visible : Visibility.Collapsed;
            HandlesCanvas.Visibility  = value ? Visibility.Visible : Visibility.Collapsed;
            if (!value && InnerEditor.IsKeyboardFocusWithin)
                Keyboard.ClearFocus();
        }
    }

    /// <summary>드래그 생성 직후 호출 — 선택 + 즉시 본문 편집 모드 진입.</summary>
    public void BeginEditing()
    {
        IsSelected = true;
        InnerEditor.Focus();
        Keyboard.Focus(InnerEditor);
    }

    private void ApplySpellCheckLanguage()
    {
        var tag = Services.LanguageService.Current == Services.LanguageService.Language.English
            ? "en-US" : "ko-KR";
        InnerEditor.Language = System.Windows.Markup.XmlLanguage.GetLanguage(tag);
    }

    // ── 모양/배경/테두리 적용 ────────────────────────────────────────

    private void ApplyShapeFromModel()
    {
        if (Model.Shape == TextBoxShape.Rectangle)
        {
            ShapeBorder.Visibility = Visibility.Visible;
            ShapePath.Visibility   = Visibility.Collapsed;

            ShapeBorder.BorderThickness = new Thickness(Math.Max(0.5, Model.BorderThicknessPt));
            if (TryParseColor(Model.BorderColor, out var bc))
                ShapeBorder.BorderBrush = new SolidColorBrush(bc);
            else
                ShapeBorder.BorderBrush = Brushes.Black;

            if (TryParseColor(Model.BackgroundColor, out var fillc))
                ShapeBorder.Background = new SolidColorBrush(fillc);
            else
                ShapeBorder.Background = Brushes.White;
        }
        else
        {
            ShapeBorder.Visibility = Visibility.Collapsed;
            ShapePath.Visibility   = Visibility.Visible;

            var pathData = Model.Shape switch
            {
                TextBoxShape.Speech    => BuildSpeechPath(Model.SpeechDirection),
                TextBoxShape.Cloud     => BuildCloudPath(Model.CloudPuffCount),
                TextBoxShape.Spiky     => BuildSpikyPath(Model.SpikeCount),
                TextBoxShape.Lightning => BuildLightningPath(Model.LightningBendCount),
                TextBoxShape.Ellipse   => BuildEllipsePath(),
                TextBoxShape.Pie       => BuildPiePath(Model.PieStartAngleDeg, Model.PieSweepAngleDeg),
                _                      => BuildSpeechPath(Model.SpeechDirection),
            };
            ShapePath.Data = Geometry.Parse(pathData);
            ShapePath.StrokeThickness = Math.Max(0.5, Model.BorderThicknessPt);

            if (TryParseColor(Model.BorderColor, out var bc))
                ShapePath.Stroke = new SolidColorBrush(bc);
            else
                ShapePath.Stroke = Brushes.Black;

            if (TryParseColor(Model.BackgroundColor, out var fillc))
                ShapePath.Fill = new SolidColorBrush(fillc);
            else
                ShapePath.Fill = Brushes.White;
        }

        // ── 안쪽 여백 (Margin = padding + shape inset) ────────────────
        UpdateInnerEditorMargin();

        // ── 가로 정렬 ─────────────────────────────────────────────────
        var ta = Model.HAlign switch
        {
            TextBoxHAlign.Center  => System.Windows.TextAlignment.Center,
            TextBoxHAlign.Right   => System.Windows.TextAlignment.Right,
            TextBoxHAlign.Justify => System.Windows.TextAlignment.Justify,
            _                     => System.Windows.TextAlignment.Left,
        };
        InnerEditor.Document.TextAlignment = ta;
        foreach (var b in InnerEditor.Document.Blocks)
            if (b is System.Windows.Documents.Paragraph wp) wp.TextAlignment = ta;

        // ── 세로 정렬 ─────────────────────────────────────────────────
        // RichTextBox 는 VerticalContentAlignment 를 지원하지 않으므로
        // VerticalAlignment(Stretch/Center/Bottom) 로 컨테이너 내 위치를 제어.
        // 내용이 박스보다 클 때는 세 경우 모두 스크롤 동작.
        InnerEditor.VerticalAlignment = Model.VAlign switch
        {
            TextBoxVAlign.Middle => VerticalAlignment.Center,
            TextBoxVAlign.Bottom => VerticalAlignment.Bottom,
            _                    => VerticalAlignment.Stretch,
        };

        // 글자 방향은 추후 지원 예정 — 현재 항상 LTR.
        InnerEditor.FlowDirection = FlowDirection.LeftToRight;

        // ── 회전 ───────────────────────────────────────────────────────
        // 박스 중심을 피벗으로 모양·본문 모두 함께 회전. 0이면 transform 제거.
        // 드래그/리사이즈 좌표는 부모 캔버스 좌표계 기반이라 회전과 무관하게 동작 —
        // 다만 회전 상태에서 핸들 드래그 시 마우스 이동 방향과 박스 변 방향이 어긋나
        // 약간 어색해 보일 수 있다 (회전 0 으로 되돌려 조절 권장).
        if (Math.Abs(Model.RotationAngleDeg) < 0.01)
        {
            RenderTransform = Transform.Identity;
        }
        else
        {
            RenderTransformOrigin = new Point(0.5, 0.5);
            RenderTransform = new RotateTransform(Model.RotationAngleDeg);
        }
    }

    // PolyDonky.Core.Color 와 충돌하므로 WpfMedia alias 로 명시.
    private static bool TryParseColor(string? hex, out WpfMedia.Color color)
    {
        color = default;
        if (string.IsNullOrWhiteSpace(hex)) return false;
        var s = hex.Trim();
        if (!s.StartsWith('#')) s = '#' + s;
        try { color = (WpfMedia.Color)WpfMedia.ColorConverter.ConvertFromString(s)!; return true; }
        catch { return false; }
    }

    // ── 모델 ↔ RichTextBox 동기화 ────────────────────────────────────

    private void LoadModelTextToEditor()
    {
        // 글자 방향은 추후 지원 예정 — 현재 항상 LTR.
        // 본문(MainWindow) 와 동일한 Run 빌더를 써서 글자 속성(폰트·크기·볼드·이탤릭·
        // 색·밑줄 등) 을 그대로 FlowDocument 에 반영한다.
        // 추가로 각 Wpf.Run.Tag 에 원본 PolyDonky.Run 을 심어둔다 — 라운드트립 시
        // FlowDocumentParser 가 Tag 를 우선 시드로 사용해 비-Wpf 속성(LetterSpacingPx,
        // WidthPercent 등) 까지 정확히 복원하고, WPF 의 inheritance 로 인한 미세한
        // 속성 drift(폰트 패밀리 자동 stamping 등) 도 막을 수 있다.
        var doc = new System.Windows.Documents.FlowDocument();
        foreach (var block in Model.Content)
        {
            if (block is PolyDonky.Core.Paragraph cp)
            {
                var wpfPara = new System.Windows.Documents.Paragraph
                {
                    Margin = new Thickness(0),
                };
                foreach (var run in cp.Runs)
                {
                    var inline = PolyDonky.App.Services.FlowDocumentBuilder.BuildInline(run);
                    if (inline is System.Windows.Documents.Run wpfRun && wpfRun.Tag is null)
                        wpfRun.Tag = run;
                    wpfPara.Inlines.Add(inline);
                }
                doc.Blocks.Add(wpfPara);
            }
        }
        if (!doc.Blocks.Any())
            doc.Blocks.Add(new System.Windows.Documents.Paragraph());

        InnerEditor.Document = doc;
    }

    private void SyncEditorToModel()
    {
        // 본문 파서를 재사용해 InnerEditor 의 FlowDocument 를 PolyDonky.Core.Block
        // 트리로 변환 — 글자 속성(폰트·볼드·이탤릭·색·밑줄 등) 과 단락 구조를 그대로 보존.
        var parsed = PolyDonky.App.Services.FlowDocumentParser.Parse(InnerEditor.Document);
        Model.Content.Clear();
        foreach (var b in parsed.Sections.FirstOrDefault()?.Blocks ?? new List<PolyDonky.Core.Block>())
            Model.Content.Add(b);
        if (Model.Content.Count == 0)
            Model.Content.Add(new PolyDonky.Core.Paragraph());
    }

    // ── 내부 편집 ────────────────────────────────────────────────────

    private void OnInnerTextChanged(object sender, TextChangedEventArgs e)
    {
        if (_suppressTextChanged) return;
        SyncEditorToModel();
        Model.Status = NodeStatus.Modified;
        ContentChangedCommitted?.Invoke(this, EventArgs.Empty);
    }


    // ── 컨텍스트 메뉴 ────────────────────────────────────────────────

    private void OnContextMenuOpened(object sender, RoutedEventArgs e)
    {
        // 선택 상태를 컨텍스트 메뉴 열릴 때 확보
        Selected?.Invoke(this, EventArgs.Empty);
        IsSelected = true;
    }

    private void OnContextMenuProperties(object sender, RoutedEventArgs e)
    {
        var dlg = new TextBoxPropertiesWindow(Model) { Owner = Window.GetWindow(this) };
        if (dlg.ShowDialog() == true)
        {
            Model.Shape              = dlg.ResultShape;
            Model.BorderColor        = dlg.ResultBorderColor;
            Model.BorderThicknessPt  = dlg.ResultBorderThicknessPt;
            Model.BackgroundColor    = dlg.ResultBackgroundColor;
            Model.PaddingTopMm       = dlg.ResultPaddingTopMm;
            Model.PaddingBottomMm    = dlg.ResultPaddingBottomMm;
            Model.PaddingLeftMm      = dlg.ResultPaddingLeftMm;
            Model.PaddingRightMm     = dlg.ResultPaddingRightMm;
            Model.HAlign             = dlg.ResultHAlign;
            Model.VAlign             = dlg.ResultVAlign;
            Model.SpeechDirection    = dlg.ResultSpeechDirection;
            Model.CloudPuffCount     = dlg.ResultCloudPuffCount;
            Model.SpikeCount         = dlg.ResultSpikeCount;
            Model.LightningBendCount = dlg.ResultLightningBendCount;
            Model.PieStartAngleDeg   = dlg.ResultPieStartAngleDeg;
            Model.PieSweepAngleDeg   = dlg.ResultPieSweepAngleDeg;
            Model.RotationAngleDeg   = dlg.ResultRotationAngleDeg;
            Model.TextOrientation    = dlg.ResultTextOrientation;
            Model.TextProgression    = dlg.ResultTextProgression;
            Model.Status             = NodeStatus.Modified;
            ApplyShapeFromModel();
            AppearanceChangedCommitted?.Invoke(this, EventArgs.Empty);
        }
    }

    private void OnContextMenuCharFormat(object sender, RoutedEventArgs e)
    {
        // Focus 를 미리 복귀시키면 inactive selection 이 collapse 되므로
        // 다이얼로그에 InnerEditor 를 그대로 전달 — Selection 포인터는 포커스
        // 없이도 유효하게 유지된다.
        // selection 이 비어 있으면 안쪽 텍스트 전체에 적용 — chrome 우클릭 흐름에서
        // 사용자가 아무것도 선택 안 했어도 시각적으로 결과가 나오도록.
        if (InnerEditor.Selection.IsEmpty) InnerEditor.SelectAll();
        var dlg = new CharFormatWindow(InnerEditor) { Owner = Window.GetWindow(this) };
        if (dlg.ShowDialog() == true)
        {
            SyncEditorToModel();
            Model.Status = NodeStatus.Modified;
            ContentChangedCommitted?.Invoke(this, EventArgs.Empty);
        }
        InnerEditor.Focus();
    }

    private void OnContextMenuParaFormat(object sender, RoutedEventArgs e)
    {
        if (InnerEditor.Selection.IsEmpty) InnerEditor.SelectAll();
        var dlg = new ParaFormatWindow(InnerEditor) { Owner = Window.GetWindow(this) };
        if (dlg.ShowDialog() == true)
        {
            SyncEditorToModel();
            Model.Status = NodeStatus.Modified;
            ContentChangedCommitted?.Invoke(this, EventArgs.Empty);
        }
        InnerEditor.Focus();
    }

    private void OnContextMenuBringForward(object sender, RoutedEventArgs e)
        => BringForwardRequested?.Invoke(this, EventArgs.Empty);

    private void OnContextMenuSendBack(object sender, RoutedEventArgs e)
        => SendBackRequested?.Invoke(this, EventArgs.Empty);

    private void OnContextMenuDelete(object sender, RoutedEventArgs e)
        => DeleteRequested?.Invoke(this, EventArgs.Empty);

    // ── 드래그 이동 ──────────────────────────────────────────────────

    private bool _dragging;
    private Point _dragStart;
    private double _dragOrigLeft, _dragOrigTop;

    private void OnRootMouseDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ChangedButton != MouseButton.Left) return;

        // 핸들 Rectangle 클릭은 OnHandleMouseDown 이 처리 — 여기서 e.Handled 를 세팅하면
        // tunneling 이 멈춰 OnHandleMouseDown 이 호출되지 않으므로 반드시 빠져나온다.
        if (e.OriginalSource is Rectangle { Tag: string rTag } &&
            rTag is "TL" or "TR" or "BL" or "BR")
            return;

        Selected?.Invoke(this, EventArgs.Empty);

        if (IsInsideEditor(e.OriginalSource as DependencyObject))
            return;

        Focus();
        if (Parent is not IInputElement parent) return;
        _dragStart    = e.GetPosition(parent);
        _dragOrigLeft = SafeGetCanvasLeft(this);
        _dragOrigTop  = SafeGetCanvasTop(this);
        _dragging     = true;
        CaptureMouse();
        e.Handled = true;
    }

    private void OnRootMouseMove(object sender, MouseEventArgs e)
    {
        if (_dragging && Parent is IInputElement parent)
        {
            var pos = e.GetPosition(parent);
            Canvas.SetLeft(this, _dragOrigLeft + (pos.X - _dragStart.X));
            Canvas.SetTop(this,  _dragOrigTop  + (pos.Y - _dragStart.Y));
            return;
        }
        if (_resizing) DoResize(e);
    }

    private void OnRootMouseUp(object sender, MouseButtonEventArgs e)
    {
        if (_dragging)
        {
            _dragging = false;
            ReleaseMouseCapture();
            GeometryChangedCommitted?.Invoke(this, EventArgs.Empty);
            e.Handled = true;
        }
        if (_resizing)
        {
            _resizing = false;
            ReleaseMouseCapture();
            GeometryChangedCommitted?.Invoke(this, EventArgs.Empty);
            e.Handled = true;
        }
    }

    private static bool IsInsideEditor(DependencyObject? source)
    {
        while (source is not null)
        {
            if (source is RichTextBox) return true;
            // Paragraph/Run 등 FrameworkContentElement 는 Visual 이 아니므로
            // VisualTreeHelper 대신 LogicalTreeHelper 를 사용해야 한다.
            source = source is Visual or Visual3D
                ? VisualTreeHelper.GetParent(source)
                : LogicalTreeHelper.GetParent(source);
        }
        return false;
    }

    // ── 4코너 리사이즈 ───────────────────────────────────────────────

    private bool _resizing;
    private string _resizeCorner = "";
    private Point _resizeStart;
    private double _resizeOrigLeft, _resizeOrigTop, _resizeOrigW, _resizeOrigH;

    private void OnHandleMouseDown(object sender, MouseButtonEventArgs e)
    {
        if (sender is not Rectangle r || r.Tag is not string corner) return;
        if (Parent is not IInputElement parent) return;

        IsSelected = true;
        Selected?.Invoke(this, EventArgs.Empty);

        _resizeCorner   = corner;
        _resizing       = true;
        _resizeStart    = e.GetPosition(parent);
        _resizeOrigLeft = SafeGetCanvasLeft(this);
        _resizeOrigTop  = SafeGetCanvasTop(this);
        _resizeOrigW    = ActualWidth;
        _resizeOrigH    = ActualHeight;
        CaptureMouse();
        e.Handled = true;
    }

    private void DoResize(MouseEventArgs e)
    {
        if (!_resizing || Parent is not IInputElement parent) return;
        var pos = e.GetPosition(parent);
        double dx = pos.X - _resizeStart.X;
        double dy = pos.Y - _resizeStart.Y;
        const double minSize = 20;

        double newL = _resizeOrigLeft, newT = _resizeOrigTop;
        double newW = _resizeOrigW,    newH = _resizeOrigH;

        switch (_resizeCorner)
        {
            case "BR":
                newW = Math.Max(minSize, _resizeOrigW + dx);
                newH = Math.Max(minSize, _resizeOrigH + dy);
                break;
            case "TR":
                newW = Math.Max(minSize, _resizeOrigW + dx);
                newH = Math.Max(minSize, _resizeOrigH - dy);
                newT = _resizeOrigTop + (_resizeOrigH - newH);
                break;
            case "BL":
                newW = Math.Max(minSize, _resizeOrigW - dx);
                newH = Math.Max(minSize, _resizeOrigH + dy);
                newL = _resizeOrigLeft + (_resizeOrigW - newW);
                break;
            case "TL":
                newW = Math.Max(minSize, _resizeOrigW - dx);
                newH = Math.Max(minSize, _resizeOrigH - dy);
                newL = _resizeOrigLeft + (_resizeOrigW - newW);
                newT = _resizeOrigTop  + (_resizeOrigH - newH);
                break;
        }

        Canvas.SetLeft(this, newL);
        Canvas.SetTop(this,  newT);
        Width  = newW;
        Height = newH;
    }

    // ── 키 입력 ─────────────────────────────────────────────────────

    private void OnRootKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Delete && !InnerEditor.IsKeyboardFocusWithin)
        {
            DeleteRequested?.Invoke(this, EventArgs.Empty);
            e.Handled = true;
        }
    }

    // ── 안쪽 편집기 여백 (사용자 padding + 비사각형 모양의 인셋) ────────
    private void UpdateInnerEditorMargin()
    {
        double padL = Model.PaddingLeftMm   * DipsPerMm;
        double padT = Model.PaddingTopMm    * DipsPerMm;
        double padR = Model.PaddingRightMm  * DipsPerMm;
        double padB = Model.PaddingBottomMm * DipsPerMm;

        var (sl, st, sr, sb) = ComputeShapeInset();
        InnerEditor.Margin = new Thickness(padL + sl, padT + st, padR + sr, padB + sb);
    }

    /// <summary>
    /// 비사각형 모양(Speech/Cloud/Spiky/Lightning) 의 외곽 돌출부(꼬리·뭉게·가시)에
    /// 텍스트가 침범하지 않도록 박스 크기에 비례한 추가 인셋을 반환한다. 사각형은 0.
    /// 박스 크기 변경 시 SizeChanged 핸들러가 다시 호출해 재계산한다.
    /// </summary>
    private (double Left, double Top, double Right, double Bottom) ComputeShapeInset()
    {
        if (Model.Shape == TextBoxShape.Rectangle) return (0, 0, 0, 0);

        double w = ActualWidth;
        double h = ActualHeight;
        if (double.IsNaN(w) || double.IsNaN(h) || w <= 0 || h <= 0) return (0, 0, 0, 0);

        switch (Model.Shape)
        {
            case TextBoxShape.Speech:
            {
                // BuildSpeechPath 의 꼬리 여백(m=15 / 100) 과 일치.
                const double pct = 0.15;
                double l = 0, t = 0, r = 0, b = 0;
                bool tL = Model.SpeechDirection is SpeechPointerDirection.Left  or SpeechPointerDirection.TopLeft  or SpeechPointerDirection.BottomLeft;
                bool tR = Model.SpeechDirection is SpeechPointerDirection.Right or SpeechPointerDirection.TopRight or SpeechPointerDirection.BottomRight;
                bool tT = Model.SpeechDirection is SpeechPointerDirection.Top   or SpeechPointerDirection.TopLeft  or SpeechPointerDirection.TopRight;
                bool tB = Model.SpeechDirection is SpeechPointerDirection.Bottom or SpeechPointerDirection.BottomLeft or SpeechPointerDirection.BottomRight;
                if (tL) l = w * pct;
                if (tR) r = w * pct;
                if (tT) t = h * pct;
                if (tB) b = h * pct;
                return (l, t, r, b);
            }
            case TextBoxShape.Cloud:
                // 베지어 뭉게가 사방으로 약간 돌출 — 10% 인셋.
                return (w * 0.10, h * 0.10, w * 0.10, h * 0.10);
            case TextBoxShape.Spiky:
                // 가시별의 외곽 정점 → 내접원 영역만 안전. 약 22% 인셋.
                return (w * 0.22, h * 0.22, w * 0.22, h * 0.22);
            case TextBoxShape.Lightning:
                // 번개는 중앙 폭이 좁아 텍스트가 잘 안 어울리지만, 최소 영역만 확보.
                return (w * 0.15, h * 0.20, w * 0.15, h * 0.20);
            case TextBoxShape.Ellipse:
                // 타원의 내접 사각형은 각 변에서 약 15% 안쪽에 위치 (r / √2 ≈ 0.707r).
                return (w * 0.15, h * 0.15, w * 0.15, h * 0.15);
            case TextBoxShape.Pie:
                // 부채꼴 — 각도에 따라 달라지므로 보수적으로 10% 인셋 적용.
                return (w * 0.10, h * 0.10, w * 0.10, h * 0.10);
            default:
                return (0, 0, 0, 0);
        }
    }

    // ── 핸들 위치 계산 ───────────────────────────────────────────────

    private void UpdateHandlePositions()
    {
        const double half = 4;
        double w = ActualWidth;
        double h = ActualHeight;
        Canvas.SetLeft(HandleTL, -half);     Canvas.SetTop(HandleTL, -half);
        Canvas.SetLeft(HandleTR, w - half);  Canvas.SetTop(HandleTR, -half);
        Canvas.SetLeft(HandleBL, -half);     Canvas.SetTop(HandleBL, h - half);
        Canvas.SetLeft(HandleBR, w - half);  Canvas.SetTop(HandleBR, h - half);
    }

    // ── 유틸 ────────────────────────────────────────────────────────

    private static double SafeGetCanvasLeft(UIElement el)
    {
        var v = Canvas.GetLeft(el);
        return double.IsNaN(v) ? 0 : v;
    }

    private static double SafeGetCanvasTop(UIElement el)
    {
        var v = Canvas.GetTop(el);
        return double.IsNaN(v) ? 0 : v;
    }

}

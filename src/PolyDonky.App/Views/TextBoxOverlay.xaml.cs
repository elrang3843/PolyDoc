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
    private bool _liveReflowQueued;
    private bool _isImeComposing;
    private bool _pendingReflow;

    /// <summary>다단 모드 여부 — Model.ColumnCount &gt; 1.</summary>
    private bool IsMultiCol => Model.ColumnCount > 1;

    /// <summary>
    /// 단일 InnerEditor 또는 다단 column RTB 중 하나라도 키보드 포커스를 가지면 true.
    /// MainWindow 의 글상자 삭제·전체 선택 단축키가 편집 중인 상태를 검사할 때 사용.
    /// </summary>
    public bool IsEditorFocusWithin
        => InnerEditor.IsKeyboardFocusWithin || MultiColHost.IsKeyboardFocusWithin;

    /// <summary>활성 편집기(단일 InnerEditor 또는 다단 활성 column RTB)에 비어있지 않은 텍스트 선택이 있으면 true.</summary>
    public bool HasEditorTextSelection
    {
        get
        {
            if (IsMultiCol)
                return MultiColHost.ActiveEditor is { } a && !a.Selection.IsEmpty;
            return !InnerEditor.Selection.IsEmpty;
        }
    }

    /// <summary>현재 키보드 포커스를 가진 편집기 (다단/단일 자동 분기).</summary>
    private RichTextBox ActiveEditor =>
        IsMultiCol ? (MultiColHost.ActiveEditor ?? MultiColHost.FirstEditor ?? InnerEditor)
                   : InnerEditor;

    /// <summary>모든 편집기 (다단=N개, 단일=1개).</summary>
    private IReadOnlyList<RichTextBox> AllEditors =>
        IsMultiCol ? MultiColHost.Editors : new[] { InnerEditor };

    public TextBoxOverlay(TextBoxObject model)
    {
        Model = model ?? throw new ArgumentNullException(nameof(model));
        InitializeComponent();

        MultiColHost.ColumnTextChanged += OnColumnTextChanged;

        // IME 조합 중에는 reflow 보류 — 단 RTB 재구성이 IME 상태를 깨뜨려
        // 한글 한 글자(예: ㅈ + ㅏ) 가 두 글자로 분리되는 문제를 방지.
        // PreviewTextInputStartEvent = IME 조합 시작, PreviewTextInputEvent = 조합 종료(확정).
        AddHandler(System.Windows.Input.TextCompositionManager.PreviewTextInputStartEvent,
            new System.Windows.Input.TextCompositionEventHandler(OnImeStart),
            handledEventsToo: true);
        AddHandler(System.Windows.Input.TextCompositionManager.PreviewTextInputEvent,
            new System.Windows.Input.TextCompositionEventHandler(OnImeEnd),
            handledEventsToo: true);

        // 초기 텍스트 로드 (plain text → FlowDocument)
        _suppressTextChanged = true;
        LoadModelTextToEditor();
        _suppressTextChanged = false;

        ApplyShapeFromModel();
        Loaded += (_, _) => UpdateHandlePositions();
        SizeChanged += (_, _) =>
        {
            UpdateHandlePositions();
            // 비사각형 모양은 인셋이 박스 크기에 비례하므로 크기 변경 시 재계산.
            if (Model.Shape != TextBoxShape.Rectangle)
                UpdateInnerEditorMargin();
            // 다단 모드는 박스 크기 변경에 따라 단 너비/높이가 달라지므로 재배치.
            if (IsMultiCol) ScheduleColumnReflow();
        };
    }

    private void OnImeStart(object sender, System.Windows.Input.TextCompositionEventArgs e)
        => _isImeComposing = true;

    private void OnImeEnd(object sender, System.Windows.Input.TextCompositionEventArgs e)
    {
        _isImeComposing = false;
        if (_pendingReflow)
        {
            _pendingReflow = false;
            ScheduleColumnReflow();
        }
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
            if (!value && (InnerEditor.IsKeyboardFocusWithin || MultiColHost.IsKeyboardFocusWithin))
                Keyboard.ClearFocus();
        }
    }

    /// <summary>드래그 생성 직후 호출 — 선택 + 즉시 본문 편집 모드 진입.</summary>
    public void BeginEditing()
    {
        IsSelected = true;
        var ed = ActiveEditor;
        ed.Focus();
        Keyboard.Focus(ed);
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

        // ── 다단 / 단일 단 분기 ─────────────────────────────────────────
        // ColumnCount > 1 이면 InnerEditor 숨기고 MultiColHost 활성화.
        if (IsMultiCol)
        {
            InnerEditor.Visibility = Visibility.Collapsed;
            MultiColHost.Visibility = Visibility.Visible;
            // 단일→다단 전환 직후 또는 박스 크기 변경 직후에는 콘텐츠를 다시 채워야 함.
            // (콘텐츠가 InnerEditor 에만 있고 MultiColHost 가 비어 있을 수 있으므로)
            if (MultiColHost.Editors.Count == 0)
            {
                _suppressTextChanged = true;
                try { LoadMultiColContent(); }
                finally { _suppressTextChanged = false; }
            }
        }
        else
        {
            InnerEditor.Visibility = Visibility.Visible;
            MultiColHost.Visibility = Visibility.Collapsed;
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
        foreach (var ed in AllEditors)
        {
            ed.Document.TextAlignment = ta;
            foreach (var b in ed.Document.Blocks)
                if (b is System.Windows.Documents.Paragraph wp) wp.TextAlignment = ta;
        }

        // ── 세로 정렬 ─────────────────────────────────────────────────
        // RichTextBox 는 VerticalContentAlignment 를 지원하지 않으므로
        // VerticalAlignment(Stretch/Center/Bottom) 로 컨테이너 내 위치를 제어.
        InnerEditor.VerticalAlignment = Model.VAlign switch
        {
            TextBoxVAlign.Middle => VerticalAlignment.Center,
            TextBoxVAlign.Bottom => VerticalAlignment.Bottom,
            _                    => VerticalAlignment.Stretch,
        };
        // 다단 모드는 호스트가 단들을 가로로 배치하므로 세로 정렬은 호스트 자체에 적용.
        MultiColHost.VerticalAlignment = InnerEditor.VerticalAlignment;

        // 글자 방향은 추후 지원 예정 — 현재 항상 LTR.
        InnerEditor.FlowDirection = FlowDirection.LeftToRight;
        MultiColHost.FlowDirection = FlowDirection.LeftToRight;

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
        if (IsMultiCol)
        {
            LoadMultiColContent();
            return;
        }

        // 단일 단 — 기존 경로.
        // 글자 방향은 추후 지원 예정 — 현재 항상 LTR.
        // 본문(MainWindow) 와 동일한 Run 빌더를 써서 글자 속성(폰트·크기·볼드·이탤릭·
        // 색·밑줄 등) 을 그대로 FlowDocument 에 반영한다.
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

    /// <summary>다단 모드 — 콘텐츠를 단별로 분배해 MultiColHost 채우기.</summary>
    private void LoadMultiColContent()
    {
        var (innerW, innerH) = ComputeInnerArea();
        if (innerW <= 0 || innerH <= 0) return;

        double colGap   = Model.ColumnGapMm * DipsPerMm;
        var widthsDip   = Model.ColumnWidthsMm?.Select(mm => mm * DipsPerMm).ToList();

        var slices = TextBoxColumnLayout.Distribute(
            Model.Content, Model.ColumnCount, innerW, innerH, colGap, widthsDip);

        MultiColHost.SetupColumns(slices, ConfigureColumnRtb);
        RefreshDividerLines(slices);
    }

    private void RefreshDividerLines(IReadOnlyList<TextBoxColumnLayout.ColumnSlice>? slices = null)
    {
        if (!IsMultiCol) return;
        if (slices is null)
        {
            // slices 없이 호출됐을 때 — 재분배 없이 현재 레이아웃 기준으로 갱신.
            var (innerW, innerH) = ComputeInnerArea();
            if (innerW <= 0 || innerH <= 0) return;
            double colGap     = Model.ColumnGapMm * DipsPerMm;
            var    widthsDip  = Model.ColumnWidthsMm?.Select(mm => mm * DipsPerMm).ToList();
            slices = TextBoxColumnLayout.Distribute(
                Model.Content, Model.ColumnCount, innerW, innerH, colGap, widthsDip);
        }
        MultiColHost.UpdateDividerLines(
            slices,
            Model.ColumnDividerVisible,
            Model.ColumnDividerColor,
            Model.ColumnDividerThicknessPt,
            Model.ColumnDividerStyle);
    }

    /// <summary>각 단 RTB 생성 직후 호출되는 콜백 — 이벤트·속성 설정.</summary>
    private void ConfigureColumnRtb(RichTextBox rtb)
    {
        rtb.SpellCheck.IsEnabled = false;
        rtb.FontFamily = InnerEditor.FontFamily;
        rtb.FlowDirection = FlowDirection.LeftToRight;
        // 본문 컨텍스트 메뉴는 InnerEditor 와 동일한 항목으로 (글자/문단 속성).
        rtb.ContextMenu = InnerEditor.ContextMenu;
        // 단락 정렬은 ApplyShapeFromModel 에서 일괄 적용.

        // 단 경계 좌/우/상/하 화살표 — RTB 가 KeyDown(bubble) 에서 화살표를 처리해 e.Handled=true
        // 로 만들기 때문에 PreviewKeyDown(tunnel) 에서 가로채야 단 이동이 가능.
        rtb.PreviewKeyDown += OnColumnPreviewKeyDown;
    }

    private void OnColumnPreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (sender is not RichTextBox cur) return;
        int idx = MultiColHost.IndexOf(cur);
        if (idx < 0) return;

        int targetIdx = -1;
        bool moveToEnd = false;

        if (e.Key == Key.Right || e.Key == Key.Down)
        {
            if (cur.CaretPosition.GetNextInsertionPosition(LogicalDirection.Forward) is null
                && idx + 1 < MultiColHost.Editors.Count)
            {
                targetIdx = idx + 1;
                moveToEnd = false;
            }
        }
        else if (e.Key == Key.Left || e.Key == Key.Up)
        {
            if (cur.CaretPosition.GetNextInsertionPosition(LogicalDirection.Backward) is null
                && idx - 1 >= 0)
            {
                targetIdx = idx - 1;
                moveToEnd = true;
            }
        }

        if (targetIdx >= 0)
        {
            var next = MultiColHost.Editors[targetIdx];
            next.Focus();
            Keyboard.Focus(next);
            next.CaretPosition = moveToEnd ? next.Document.ContentEnd : next.Document.ContentStart;
            e.Handled = true;
        }
    }

    private void SyncEditorToModel()
    {
        if (IsMultiCol)
        {
            SyncMultiColContentToModel();
            return;
        }

        // 단일 단 — 기존 경로.
        var parsed = PolyDonky.App.Services.FlowDocumentParser.Parse(InnerEditor.Document);
        Model.Content.Clear();
        foreach (var b in parsed.Sections.FirstOrDefault()?.Blocks ?? new List<PolyDonky.Core.Block>())
            Model.Content.Add(b);
        if (Model.Content.Count == 0)
            Model.Content.Add(new PolyDonky.Core.Paragraph());
    }

    /// <summary>다단 모드 — 모든 단의 RTB 를 파싱하고 frag0/frag1 단락 머지 후 모델 갱신.</summary>
    private void SyncMultiColContentToModel()
    {
        var rawBlocks = new List<PolyDonky.Core.Block>();
        foreach (var rtb in MultiColHost.Editors)
        {
            var parsed = PolyDonky.App.Services.FlowDocumentParser.Parse(rtb.Document);
            if (parsed.Sections.FirstOrDefault() is { } sec)
                foreach (var b in sec.Blocks)
                    rawBlocks.Add(b);
        }

        Model.Content.Clear();
        foreach (var b in MergeColumnFragments(rawBlocks))
        {
            // 빈 단 placeholder(§tbph) 처리: 비어있으면 모델에서 제거, 사용자 입력이 있으면 마커만 클리어.
            if (b is PolyDonky.Core.Paragraph p && p.Id == TextBoxColumnLayout.PlaceholderId)
            {
                if (p.Runs.All(r => string.IsNullOrEmpty(r.Text)))
                    continue;
                p.Id = null;
            }
            Model.Content.Add(b);
        }
        if (Model.Content.Count == 0)
            Model.Content.Add(new PolyDonky.Core.Paragraph());
    }

    /// <summary>
    /// 본문 다단의 <c>MainWindow.MergeColumnFragments</c> 와 동일 로직 — frag1+frag2 재결합.
    /// <para>
    /// WPF 의 엔터 단락 분할(<c>EditingCommands.EnterParagraphBreak</c>) 은 첫째 반쪽이
    /// 원본 element 를 유지하므로 그 element 의 <c>Tag</c>(=원본 Core.Paragraph) 가 그대로 남는다.
    /// 둘째 반쪽도 일부 시나리오에서 같은 Tag 가 따라가는 경우가 있어, 두 반쪽 모두 같은
    /// <c>§f</c> fragIdx 를 갖고 머지에서 §f0 으로 되돌아가는 버그가 발생한다(엔터가
    /// 시각적으로 사라짐). 같은 group 안에서 같은 fragIdx 가 두 번째로 등장하면 그 단락은
    /// 사용자의 엔터로 새로 생긴 단락으로 간주해 머지하지 않고 결과에 그대로 추가한다.
    /// </para>
    /// </summary>
    private static IEnumerable<PolyDonky.Core.Block> MergeColumnFragments(
        IList<PolyDonky.Core.Block> blocks)
    {
        const string FragSep   = "§f";
        const string GenPrefix = "§g";

        var result      = new List<PolyDonky.Core.Block>();
        var openFrags   = new Dictionary<string, PolyDonky.Core.Paragraph>();
        var seenFragIdx = new Dictionary<string, HashSet<string>>();

        foreach (var block in blocks)
        {
            if (block is PolyDonky.Core.Paragraph p && p.Id is { } id)
            {
                int sepIdx = id.LastIndexOf(FragSep, StringComparison.Ordinal);
                if (sepIdx >= 0)
                {
                    string groupId    = id[..sepIdx];
                    string fragIdxStr = id[(sepIdx + FragSep.Length)..];

                    // 같은 group 의 같은 fragIdx 가 두 번째 등장이면 — WPF 엔터로 분할된
                    // 두 반쪽이 같은 §f Id 를 공유하는 케이스. 둘째는 사용자가 새로 만든 단락이므로
                    // §f 머지 대상에서 빼고 그대로 추가한다.
                    if (!seenFragIdx.TryGetValue(groupId, out var seen))
                    {
                        seen = new HashSet<string>(StringComparer.Ordinal);
                        seenFragIdx[groupId] = seen;
                    }
                    if (!seen.Add(fragIdxStr))
                    {
                        p.Id = null;
                        result.Add(p);
                        continue;
                    }

                    bool isFirst = fragIdxStr == "0";
                    if (isFirst)
                    {
                        p.Id = groupId.StartsWith(GenPrefix, StringComparison.Ordinal) ? null : groupId;
                        result.Add(p);
                        openFrags[groupId] = p;
                    }
                    else
                    {
                        if (openFrags.TryGetValue(groupId, out var target))
                            foreach (var run in p.Runs) target.Runs.Add(run);
                        else
                        {
                            p.Id = groupId.StartsWith(GenPrefix, StringComparison.Ordinal) ? null : groupId;
                            result.Add(p);
                        }
                    }
                    continue;
                }
            }
            result.Add(block);
        }
        return result;
    }

    /// <summary>박스 안쪽 영역(여백·모양 인셋 제외) DIP 단위 (W, H).</summary>
    private (double W, double H) ComputeInnerArea()
    {
        double w = ActualWidth, h = ActualHeight;
        if (double.IsNaN(w) || w <= 0)
            w = Model.WidthMm * DipsPerMm;
        if (double.IsNaN(h) || h <= 0)
            h = Model.HeightMm * DipsPerMm;

        double padL = Model.PaddingLeftMm   * DipsPerMm;
        double padT = Model.PaddingTopMm    * DipsPerMm;
        double padR = Model.PaddingRightMm  * DipsPerMm;
        double padB = Model.PaddingBottomMm * DipsPerMm;
        var (sl, st, sr, sb) = ComputeShapeInset();

        double innerW = Math.Max(1.0, w - padL - padR - sl - sr);
        double innerH = Math.Max(1.0, h - padT - padB - st - sb);
        return (innerW, innerH);
    }

    /// <summary>
    /// 캐럿 위치 마커 — 리플로우 동안 콘텐츠와 함께 떠다니다가 재분배 후 캐럿 위치를 알려준다.
    /// PUA(사용자 영역) 코드포인트 3 개로 구성 — 사용자 텍스트와 충돌 가능성이 거의 없고,
    /// 폭이 0(또는 폴백 글리프) 라 시각 영향 미미. offset 기반 추적으로는
    /// 단락 경계(엔터 직후) / 다단 fragment 분할에서 위치를 정확히 복원할 수 없어
    /// "콘텐츠를 따라다니는 마커" 방식으로 대체.
    /// </summary>
    private const string CaretMarker = "\u2060\u2060\u2060";

    /// <summary>다단 라이브 리플로우 — TextChanged 시 콘텐츠 재분배 + RTB 재구성.</summary>
    private void ScheduleColumnReflow()
    {
        if (_liveReflowQueued || _suppressTextChanged) return;
        if (!IsMultiCol) return;
        _liveReflowQueued = true;
        Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Background, () =>
        {
            _liveReflowQueued = false;

            // IME 조합 중에는 reflow 보류 — 조합이 끝나면 OnImeEnd 가 다시 호출한다.
            if (_isImeComposing)
            {
                _pendingReflow = true;
                return;
            }

            try
            {
                // 1) 캐럿 위치에 마커 삽입 — 리플로우/재분배 동안 콘텐츠와 함께 이동.
                bool markerInserted = InsertCaretMarker();

                // 2) 모든 단 파싱 + frag 머지 → 모델 갱신
                SyncMultiColContentToModel();
                // 3) 새 분배로 다시 채움
                _suppressTextChanged = true;
                try { LoadMultiColContent(); }
                finally { _suppressTextChanged = false; }

                // 4) 마커 위치 찾아 캐럿/포커스 복원 + 마커 제거.
                if (markerInserted)
                    RestoreCaretAtMarker();
            }
            catch { /* 실패 시 캐시 유지 */ }
        });
    }

    /// <summary>활성 단 RTB 의 현재 캐럿 위치에 <see cref="CaretMarker"/> 텍스트 삽입.</summary>
    private bool InsertCaretMarker()
    {
        var actEd = MultiColHost.ActiveEditor;
        if (actEd is null) return false;
        try
        {
            _suppressTextChanged = true;
            try
            {
                new TextRange(actEd.CaretPosition, actEd.CaretPosition).Text = CaretMarker;
            }
            finally { _suppressTextChanged = false; }
            return true;
        }
        catch { return false; }
    }

    /// <summary>모든 단 RTB 에서 마커를 찾아 첫 발견 위치에 캐럿/포커스 설정 + 마커 제거.</summary>
    private void RestoreCaretAtMarker()
    {
        for (int i = 0; i < MultiColHost.Editors.Count; i++)
        {
            var ed = MultiColHost.Editors[i];
            var (start, end) = FindMarker(ed.Document);
            if (start is not null && end is not null)
            {
                _suppressTextChanged = true;
                try
                {
                    // 마커 텍스트 제거 — TextPointer 는 live 라 삭제 후에도 같은 위치 유지.
                    new TextRange(start, end).Text = "";
                }
                finally { _suppressTextChanged = false; }
                FocusEditorWithCaret(ed, start);
                return;
            }
        }
        // 마커를 못 찾으면 폴백 — 마지막 단 끝(예전 동작 보존).
        if (MultiColHost.Editors.Count > 0)
        {
            var last = MultiColHost.Editors[^1];
            FocusEditorWithCaret(last, last.Document.ContentEnd);
        }
    }

    /// <summary>FlowDocument 안에서 <see cref="CaretMarker"/> 의 시작/끝 TextPointer 를 반환.</summary>
    private static (TextPointer? start, TextPointer? end) FindMarker(FlowDocument doc)
    {
        var pointer = doc.ContentStart;
        while (pointer != null && pointer.CompareTo(doc.ContentEnd) < 0)
        {
            if (pointer.GetPointerContext(LogicalDirection.Forward) == TextPointerContext.Text)
            {
                int run = pointer.GetTextRunLength(LogicalDirection.Forward);
                string text = pointer.GetTextInRun(LogicalDirection.Forward);
                int idx = text.IndexOf(CaretMarker, StringComparison.Ordinal);
                if (idx >= 0)
                {
                    var start = pointer.GetPositionAtOffset(idx, LogicalDirection.Forward);
                    var end   = pointer.GetPositionAtOffset(idx + CaretMarker.Length, LogicalDirection.Backward);
                    return (start, end);
                }
                pointer = pointer.GetPositionAtOffset(run, LogicalDirection.Forward) ?? doc.ContentEnd;
            }
            else
            {
                pointer = pointer.GetNextContextPosition(LogicalDirection.Forward) ?? doc.ContentEnd;
            }
        }
        return (null, null);
    }

    /// <summary>편집기에 포커스 설정 + 캐럿 위치 적용. SetupColumns 직후 발생할 수 있는 focus race 회피.</summary>
    private static void FocusEditorWithCaret(RichTextBox ed, TextPointer caret)
    {
        // Focus → Keyboard.Focus → CaretPosition 순서 — Focus 가 안정된 뒤 Caret 설정해야
        // WPF 가 caret 위치를 visual tree 정착 후 정확히 표시.
        ed.Focus();
        Keyboard.Focus(ed);
        ed.CaretPosition = caret;
    }

    private void OnColumnTextChanged(object? sender, TextChangedEventArgs e)
    {
        if (_suppressTextChanged) return;
        // Model.Content 는 reflow 가 끝난 뒤 SyncMultiColContentToModel 에서 갱신.
        Model.Status = NodeStatus.Modified;
        ContentChangedCommitted?.Invoke(this, EventArgs.Empty);
        ScheduleColumnReflow();
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

            // 단 수 / 단 간격 변경 — 단일↔다단 전환 시 콘텐츠 재로드.
            int oldColCount = Model.ColumnCount;
            Model.ColumnCount = Math.Max(1, dlg.ResultColumnCount);
            Model.ColumnGapMm = Math.Max(0, dlg.ResultColumnGapMm);
            // 단 수가 바뀌면 ColumnWidthsMm(사용자 지정)는 무효화 — 균등 배분으로 리셋.
            if (Model.ColumnCount != oldColCount)
                Model.ColumnWidthsMm = null;

            // 단 구분선 속성
            Model.ColumnDividerVisible     = dlg.ResultColumnDividerVisible;
            Model.ColumnDividerColor       = dlg.ResultColumnDividerColor;
            Model.ColumnDividerThicknessPt = dlg.ResultColumnDividerThicknessPt;
            Model.ColumnDividerStyle       = dlg.ResultColumnDividerStyle;

            Model.Status = NodeStatus.Modified;

            // 단일→다단 또는 다단→단일 전환 시 콘텐츠를 새 모드로 재로드.
            if (oldColCount != Model.ColumnCount)
            {
                // 기존 편집기에서 모델로 sync (전환 전 마지막 상태 보존)
                _suppressTextChanged = true;
                try
                {
                    if (oldColCount > 1)
                    {
                        // 다단 → 단일: 모든 단 RTB 머지해서 모델에 저장 후 InnerEditor 로 로드
                        SyncMultiColContentToModel();
                    }
                    else
                    {
                        // 단일 → 다단: InnerEditor 의 콘텐츠를 모델에 저장
                        var parsed = PolyDonky.App.Services.FlowDocumentParser.Parse(InnerEditor.Document);
                        Model.Content.Clear();
                        foreach (var b in parsed.Sections.FirstOrDefault()?.Blocks
                                 ?? new List<PolyDonky.Core.Block>())
                            Model.Content.Add(b);
                        if (Model.Content.Count == 0)
                            Model.Content.Add(new PolyDonky.Core.Paragraph());
                    }
                    LoadModelTextToEditor();
                }
                finally { _suppressTextChanged = false; }
            }
            else if (IsMultiCol)
            {
                // 단 간격/구분선만 바뀐 경우 — 분배 다시.
                ScheduleColumnReflow();
                RefreshDividerLines();
            }

            ApplyShapeFromModel();
            AppearanceChangedCommitted?.Invoke(this, EventArgs.Empty);
        }
    }

    private void OnContextMenuCharFormat(object sender, RoutedEventArgs e)
    {
        // 글자 속성: 선택이 있으면 그 영역, 없으면 caret 위치(이후 입력에 적용) — SelectAll 강제 안 함.
        // 사용자가 "글상자 전체가 바뀌는" 동작을 원하지 않을 때 selection 없이 호출 가능.
        var ed = ActiveEditor;
        var dlg = new CharFormatWindow(ed) { Owner = Window.GetWindow(this) };
        if (dlg.ShowDialog() == true)
        {
            SyncEditorToModel();
            Model.Status = NodeStatus.Modified;
            ContentChangedCommitted?.Invoke(this, EventArgs.Empty);
            if (IsMultiCol) ScheduleColumnReflow();
        }
        ed.Focus();
    }

    private void OnContextMenuParaFormat(object sender, RoutedEventArgs e)
    {
        // 문단 속성은 selection 이 비어있으면 caret 위치 단락에만 적용 — SelectAll 하지 않음.
        // (CharFormat 과 다름: 글자 속성은 chars 가 선택돼야 의미 있어 SelectAll 강제하지만,
        //  문단 속성은 caret 만 있어도 ParaFormatWindow.CollectParagraphs 가 caret 단락 1개 반환.)
        var ed = ActiveEditor;
        var dlg = new ParaFormatWindow(ed) { Owner = Window.GetWindow(this) };
        if (dlg.ShowDialog() == true)
        {
            SyncEditorToModel();
            Model.Status = NodeStatus.Modified;
            ContentChangedCommitted?.Invoke(this, EventArgs.Empty);
            if (IsMultiCol) ScheduleColumnReflow();
        }
        ed.Focus();
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
        // Del 은 chrome(편집기 외부 영역) 에서만 박스 삭제. 편집기(InnerEditor 또는 다단 RTB) 안에서
        // 발생한 Del 은 RTB 가 텍스트 삭제로 처리하도록 그대로 흘려보낸다.
        if (e.Key == Key.Delete && !IsInsideEditor(e.OriginalSource as DependencyObject))
        {
            DeleteRequested?.Invoke(this, EventArgs.Empty);
            e.Handled = true;
            return;
        }
        // 다단 단 경계 화살표 이동은 OnColumnPreviewKeyDown 에서 처리 — RTB 가 화살표를
        // 가로채기 전에 PreviewKeyDown 에서 잡아야 함.
    }

    // ── 안쪽 편집기 여백 (사용자 padding + 비사각형 모양의 인셋) ────────
    private void UpdateInnerEditorMargin()
    {
        double padL = Model.PaddingLeftMm   * DipsPerMm;
        double padT = Model.PaddingTopMm    * DipsPerMm;
        double padR = Model.PaddingRightMm  * DipsPerMm;
        double padB = Model.PaddingBottomMm * DipsPerMm;

        var (sl, st, sr, sb) = ComputeShapeInset();
        var margin = new Thickness(padL + sl, padT + st, padR + sr, padB + sb);
        InnerEditor.Margin  = margin;
        MultiColHost.Margin = margin;
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

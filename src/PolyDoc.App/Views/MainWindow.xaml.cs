using System;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
using PolyDoc.App.ViewModels;
using PolyDoc.Core;
using SR = PolyDoc.App.Properties.Resources;
using WpfMedia = System.Windows.Media;

namespace PolyDoc.App.Views;

public partial class MainWindow : Window
{
    private MainViewModel? _viewModel;
    private bool _suppressTextChanged;
    private DispatcherTimer? _statusTimer;

    // ── 글상자 드래그 생성 / 선택 상태 ────────────────────────────
    private bool _drawingTextBox;
    private bool _drawingInProgress;
    private Point _drawStart;
    private TextBoxShape _drawingShape = TextBoxShape.Rectangle;
    private TextBoxOverlay? _selectedOverlay;

    public MainWindow()
    {
        InitializeComponent();

        Loaded += OnLoaded;
        Unloaded += OnUnloaded;
        BodyEditor.TextChanged += OnEditorTextChanged;
        // 상태 표시줄 Insert/CapsLock/NumLock 갱신을 위해 윈도우 레벨 키 입력 가로채기.
        PreviewKeyDown += OnPreviewKeyDown;
    }

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        if (DataContext is MainViewModel vm)
        {
            _viewModel = vm;
            ApplyFlowDocument(vm.FlowDocument);
            vm.PropertyChanged += OnViewModelPropertyChanged;
            vm.FindReplaceRequested  += OnFindReplaceRequested;
            vm.SettingsRequested     += OnSettingsRequested;
            vm.OutlineStyleRequested += OnOutlineStyleRequested;
            vm.InsertTextBoxRequested += OnInsertTextBoxRequested;
            vm.RefreshSystemKeys();
            vm.RefreshMemoryUsage();
        }

        // RichTextBox 클릭 = 본문 편집 의도. 드래그 생성 모드가 아니면 글상자 선택 해제.
        BodyEditor.PreviewMouseLeftButtonDown += (_, _) =>
        {
            if (!_drawingTextBox) DeselectAllOverlays();
        };

        _statusTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(1),
        };
        _statusTimer.Tick += OnStatusTimerTick;
        _statusTimer.Start();
    }

    private void OnUnloaded(object sender, RoutedEventArgs e)
    {
        if (_statusTimer is not null)
        {
            _statusTimer.Stop();
            _statusTimer.Tick -= OnStatusTimerTick;
            _statusTimer = null;
        }
    }

    private void OnStatusTimerTick(object? sender, EventArgs e)
    {
        _viewModel?.RefreshSystemKeys();
        _viewModel?.RefreshMemoryUsage();
    }

    private void OnPreviewKeyDown(object sender, KeyEventArgs e)
    {
        switch (e.Key)
        {
            case Key.Insert:
                _viewModel?.ToggleInsertMode();
                break;
            case Key.CapsLock:
            case Key.NumLock:
                // 토글 상태는 KeyDown 직후엔 Keyboard.IsKeyToggled 가 갱신 전이라
                // Dispatcher.BeginInvoke 로 다음 사이클에 읽는다.
                Dispatcher.BeginInvoke(new Action(() => _viewModel?.RefreshSystemKeys()),
                    DispatcherPriority.Input);
                break;
            case Key.Escape:
                if (_drawingTextBox)
                {
                    EndDrawingMode();
                    e.Handled = true;
                }
                else if (_selectedOverlay is not null)
                {
                    DeselectAllOverlays();
                    e.Handled = true;
                }
                break;
        }
    }

    private void OnViewModelPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (_viewModel is null) return;
        if (e.PropertyName == nameof(MainViewModel.FlowDocument))
        {
            ApplyFlowDocument(_viewModel.FlowDocument);
            _viewModel.RefreshMemoryUsage();
        }
    }

    private void ApplyFlowDocument(System.Windows.Documents.FlowDocument fd)
    {
        // FlowDocument 는 RichTextBox 의 자식이지만 자체 시각 트리 루트라
        // RichTextBox.Foreground 가 Run 까지 전파되지 않는다.
        // SetResourceReference 로 테마 사전을 동적 바인딩 — 테마 교체 시 자동 갱신.
        fd.SetResourceReference(System.Windows.Documents.FlowDocument.ForegroundProperty, "OnSurface");
        // Background 는 PaperBorder 가 담당하므로 FlowDocument 는 투명으로 둔다.
        fd.Background = Brushes.Transparent;

        _suppressTextChanged = true;
        try
        {
            BodyEditor.Document = fd;
        }
        finally
        {
            _suppressTextChanged = false;
        }

        // 용지 크기·색상을 PaperBorder 에 반영
        var page = _viewModel?.Document.Sections.FirstOrDefault()?.Page;
        ApplyPageSettings(page);

        // 부유 객체 (글상자 등) 오버레이 재구축
        RebuildFloatingObjects();
    }

    private void ApplyPageSettings(PolyDoc.Core.PageSettings? page)
    {
        if (page is null) return;

        double padL = PolyDoc.App.Services.FlowDocumentBuilder.MmToDip(page.MarginLeftMm);
        double padT = PolyDoc.App.Services.FlowDocumentBuilder.MmToDip(page.MarginTopMm);
        double padR = PolyDoc.App.Services.FlowDocumentBuilder.MmToDip(page.MarginRightMm);
        double padB = PolyDoc.App.Services.FlowDocumentBuilder.MmToDip(page.MarginBottomMm);

        // 용지 너비·높이 (세로·가로 방향 보정).
        // Height 는 MinHeight 로 지정해 빈 문서도 한 페이지 분량으로 보이고, 본문이 길어지면 늘어난다.
        PaperBorder.Width     = PolyDoc.App.Services.FlowDocumentBuilder.MmToDip(page.EffectiveWidthMm);
        PaperBorder.MinHeight = PolyDoc.App.Services.FlowDocumentBuilder.MmToDip(page.EffectiveHeightMm);

        // 여백을 RichTextBox Padding 으로 반영 — FlowDocument.PagePadding 은 RichTextBox 컨텍스트에서 무시됨
        BodyEditor.Padding = new Thickness(padL, padT, padR, padB);

        // 여백 안내선 위치 갱신
        MarginGuideRect.Margin     = new Thickness(padL, padT, padR, padB);
        MarginGuideRect.Visibility = page.ShowMarginGuides
            ? System.Windows.Visibility.Visible
            : System.Windows.Visibility.Collapsed;

        // 용지 배경색 — 지정된 경우 SolidColorBrush, 없으면 테마 Surface 동적 리소스
        if (!string.IsNullOrEmpty(page.PaperColor))
        {
            try
            {
                // PolyDoc.Core.Color 와 충돌하므로 WpfMedia alias 로 명시.
                var c = (WpfMedia.Color)WpfMedia.ColorConverter.ConvertFromString(page.PaperColor)!;
                PaperBorder.Background = new SolidColorBrush(c);
            }
            catch
            {
                PaperBorder.SetResourceReference(System.Windows.Controls.Border.BackgroundProperty, "Surface");
            }
        }
        else
        {
            PaperBorder.SetResourceReference(System.Windows.Controls.Border.BackgroundProperty, "Surface");
        }
    }

    private void OnEditorTextChanged(object sender, TextChangedEventArgs e)
    {
        if (_suppressTextChanged) return;
        _viewModel?.MarkDirty();
    }

    private void OnFindReplaceRequested(object? sender, EventArgs e)
    {
        var dlg = new FindReplaceWindow(BodyEditor) { Owner = this };
        dlg.Show();
    }

    private void OnSettingsRequested(object? sender, EventArgs e)
    {
        var dlg = new SettingsWindow { Owner = this };
        dlg.ShowDialog();
    }

    private void OnOutlineStyleRequested(object? sender, EventArgs e)
    {
        var current = (_viewModel is { } vm2 && vm2.Document?.OutlineStyles != null)
            ? vm2.Document.OutlineStyles
            : PolyDoc.Core.OutlineStyleSet.CreateDefault();
        var dlg = new OutlineStyleWindow(current) { Owner = this };
        dlg.StyleApplied += (_, styleSet) => _viewModel?.ApplyOutlineStyles(styleSet);
        dlg.ShowDialog();
    }

    private void OnFormatChar(object sender, RoutedEventArgs e)
    {
        var dlg = new CharFormatWindow(BodyEditor) { Owner = this };
        if (dlg.ShowDialog() == true)
        {
            _viewModel?.MarkDirty();
            BodyEditor.Focus();
        }
    }

    private void OnFormatPara(object sender, RoutedEventArgs e)
    {
        var dlg = new ParaFormatWindow(BodyEditor) { Owner = this };
        if (dlg.ShowDialog() == true)
        {
            _viewModel?.MarkDirty();
            BodyEditor.Focus();
        }
    }

    private void OnFormatPage(object sender, RoutedEventArgs e)
    {
        var current = _viewModel?.Document.Sections.FirstOrDefault()?.Page
                      ?? new PolyDoc.Core.PageSettings();
        var dlg = new PageFormatWindow(current) { Owner = this };
        if (dlg.ShowDialog() == true)
        {
            if (_viewModel?.Document.Sections.FirstOrDefault() is { } section)
                section.Page = dlg.ResultSettings;
            _viewModel?.RebuildFlowDocument();
            _viewModel?.MarkDirty();
        }
    }

    // 편집 > 지우기: RichTextBox 는 ApplicationCommands.Delete 를 자체 바인딩하지 않으므로
    // 메뉴에서 직접 호출 시 동작하도록 선택 영역을 지운다. 선택이 비어 있으면 캐럿 직후
    // 한 글자(EditingCommands.Delete) 를 지우는 일반 워드프로세서 동작을 따른다.
    private void OnEditDelete(object sender, RoutedEventArgs e)
    {
        if (!BodyEditor.Selection.IsEmpty)
        {
            BodyEditor.Selection.Text = string.Empty;
        }
        else
        {
            System.Windows.Documents.EditingCommands.Delete.Execute(null, BodyEditor);
        }
        BodyEditor.Focus();
    }

    // 쓰기 보호: IsReadOnly 가 켜져 있으면 RichTextBox 가 자체적으로 입력을 막지만,
    // 사용자에게 "왜 안 되지?" 의 침묵 대신 비밀번호 프롬프트를 띄워 즉시 잠금 해제 흐름으로 안내한다.
    // 검증 성공 시 ViewModel 이 IsWriteProtected=false 로 풀고, 다음 키부터 바로 편집된다.
    private void OnEditorPreviewTextInput(object sender, TextCompositionEventArgs e)
    {
        if (_viewModel?.IsWriteProtected != true) return;
        e.Handled = true;
        _viewModel.TryUnlockForEditing();
    }

    private void OnEditorPreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (_viewModel?.IsWriteProtected != true) return;
        if (!IsEditingIntent(e)) return;
        e.Handled = true;
        _viewModel.TryUnlockForEditing();
    }

    /// <summary>
    /// 입력 키가 "편집 의도" 인지 판별. 화살표·Home/End·Ctrl+C/Ctrl+A 등 읽기 전용 동작은 false.
    /// (PreviewTextInput 이 일반 문자 입력은 별도로 처리하므로 여기서는 특수 편집 키만 본다.)
    /// </summary>
    private static bool IsEditingIntent(KeyEventArgs e)
    {
        var ctrl = (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control;
        return e.Key switch
        {
            Key.Back or Key.Delete or Key.Enter or Key.Tab => true,
            Key.V or Key.X when ctrl                       => true,  // 붙여넣기 / 잘라내기
            _                                              => false,
        };
    }

    private void OnDragOver(object sender, DragEventArgs e)
    {
        if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;
        e.Effects = DragDropEffects.Copy;
        e.Handled = true;
    }

    private void OnDrop(object sender, DragEventArgs e)
    {
        if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;
        if (e.Data.GetData(DataFormats.FileDrop) is string[] files && files.Length > 0)
        {
            _viewModel?.OpenFile(files[0]);
        }
        e.Handled = true;
    }

    // ── 확대/축소 조절기 ──────────────────────────────────────────────────────

    private void OnFitToWidth(object sender, RoutedEventArgs e)
    {
        if (_viewModel is null) return;
        // 뷰포트 너비에서 PaperStackPanel 좌우 여백(32+32) 을 빼고 용지 논리 너비로 나눈다.
        const double hMargin = 64;
        double paperW = PaperBorder.Width;
        double viewW  = EditorScrollViewer.ViewportWidth;
        if (paperW <= 0 || viewW <= 0) return;
        _viewModel.ZoomPercent = (viewW - hMargin) / paperW * 100;
    }

    private void OnFitToPage(object sender, RoutedEventArgs e)
    {
        if (_viewModel is null) return;
        const double hMargin = 64; // StackPanel 좌우 여백
        const double vMargin = 76; // StackPanel 상(28)+하(48) 여백
        double paperW = PaperBorder.Width;
        double paperH = PaperBorder.MinHeight; // 콘텐츠가 길어도 한 페이지 분량 기준
        double viewW  = EditorScrollViewer.ViewportWidth;
        double viewH  = EditorScrollViewer.ViewportHeight;
        if (paperW <= 0 || paperH <= 0 || viewW <= 0 || viewH <= 0) return;
        double fitW = (viewW - hMargin) / paperW;
        double fitH = (viewH - vMargin) / paperH;
        _viewModel.ZoomPercent = Math.Min(fitW, fitH) * 100;
    }

    private void OnZoomTextBoxKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key != Key.Enter) return;
        ApplyZoomFromTextBox();
        BodyEditor.Focus();
        e.Handled = true;
    }

    private void OnZoomTextBoxLostFocus(object sender, RoutedEventArgs e)
        => ApplyZoomFromTextBox();

    private void ApplyZoomFromTextBox()
    {
        if (_viewModel is null) return;
        var text = TxtZoom.Text.Trim().TrimEnd('%');
        if (double.TryParse(text, out double val))
            _viewModel.ZoomPercent = val;
        else
            TxtZoom.Text = _viewModel.ZoomPercent.ToString("0");
    }

    // ── 글상자 (입력 > 글상자) ──────────────────────────────────────────────

    private void OnInsertTextBoxRequested(object? sender, TextBoxShape shape)
    {
        _drawingTextBox = true;
        _drawingShape   = shape;
        Mouse.OverrideCursor = Cursors.Cross;
        if (_viewModel is not null)
            _viewModel.StatusMessage = SR.StatusDrawTextBox;
    }

    private void EndDrawingMode()
    {
        _drawingTextBox = false;
        _drawingInProgress = false;
        Mouse.OverrideCursor = null;
        DrawPreviewRect.Visibility = Visibility.Collapsed;
        if (PaperBorder.IsMouseCaptured) PaperBorder.ReleaseMouseCapture();
        if (_viewModel is not null) _viewModel.StatusMessage = SR.StatusReady;
    }

    private void OnPaperPreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (!_drawingTextBox) return;

        var pos = e.GetPosition(PaperBorder);
        pos.X = Math.Clamp(pos.X, 0, PaperBorder.ActualWidth);
        pos.Y = Math.Clamp(pos.Y, 0, PaperBorder.ActualHeight);

        _drawStart = pos;
        _drawingInProgress = true;

        Canvas.SetLeft(DrawPreviewRect, pos.X);
        Canvas.SetTop(DrawPreviewRect, pos.Y);
        DrawPreviewRect.Width = 0;
        DrawPreviewRect.Height = 0;
        DrawPreviewRect.Visibility = Visibility.Visible;

        PaperBorder.CaptureMouse();
        e.Handled = true;
    }

    private void OnPaperPreviewMouseMove(object sender, MouseEventArgs e)
    {
        if (!_drawingInProgress) return;

        var pos = e.GetPosition(PaperBorder);
        pos.X = Math.Clamp(pos.X, 0, PaperBorder.ActualWidth);
        pos.Y = Math.Clamp(pos.Y, 0, PaperBorder.ActualHeight);

        double x = Math.Min(_drawStart.X, pos.X);
        double y = Math.Min(_drawStart.Y, pos.Y);
        double w = Math.Abs(pos.X - _drawStart.X);
        double h = Math.Abs(pos.Y - _drawStart.Y);

        Canvas.SetLeft(DrawPreviewRect, x);
        Canvas.SetTop(DrawPreviewRect, y);
        DrawPreviewRect.Width  = w;
        DrawPreviewRect.Height = h;
    }

    private void OnPaperPreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (!_drawingInProgress) return;

        var pos = e.GetPosition(PaperBorder);
        pos.X = Math.Clamp(pos.X, 0, PaperBorder.ActualWidth);
        pos.Y = Math.Clamp(pos.Y, 0, PaperBorder.ActualHeight);

        double x = Math.Min(_drawStart.X, pos.X);
        double y = Math.Min(_drawStart.Y, pos.Y);
        double w = Math.Abs(pos.X - _drawStart.X);
        double h = Math.Abs(pos.Y - _drawStart.Y);

        EndDrawingMode();

        // 너무 작은 드래그(클릭에 가까움) 는 생성 안 함.
        const double minDip = 10;
        if (w < minDip || h < minDip)
        {
            e.Handled = true;
            return;
        }

        var model = new TextBoxObject
        {
            Shape    = _drawingShape,
            XMm      = x / TextBoxOverlay.DipsPerMm,
            YMm      = y / TextBoxOverlay.DipsPerMm,
            WidthMm  = w / TextBoxOverlay.DipsPerMm,
            HeightMm = h / TextBoxOverlay.DipsPerMm,
            Status   = NodeStatus.Modified,
        };
        _viewModel?.AddFloatingObjectToCurrentSection(model);
        var overlay = AddTextBoxOverlay(model);
        SelectOverlay(overlay);
        overlay.BeginEditing();

        e.Handled = true;
    }

    /// <summary>현재 섹션의 FloatingObjects 를 캔버스에 다시 채워 그린다 (문서 로드 시).</summary>
    private void RebuildFloatingObjects()
    {
        FloatingCanvas.Children.Clear();
        _selectedOverlay = null;
        var section = _viewModel?.Document.Sections.FirstOrDefault();
        if (section is null) return;
        foreach (var obj in section.FloatingObjects.OfType<TextBoxObject>())
        {
            AddTextBoxOverlay(obj);
        }
    }

    private TextBoxOverlay AddTextBoxOverlay(TextBoxObject model)
    {
        var overlay = new TextBoxOverlay(model);
        Canvas.SetLeft(overlay, model.XMm * TextBoxOverlay.DipsPerMm);
        Canvas.SetTop(overlay,  model.YMm * TextBoxOverlay.DipsPerMm);
        overlay.Width  = model.WidthMm  * TextBoxOverlay.DipsPerMm;
        overlay.Height = model.HeightMm * TextBoxOverlay.DipsPerMm;

        overlay.Selected += (_, _) => SelectOverlay(overlay);

        overlay.BringForwardRequested += (_, _) =>
        {
            int idx = FloatingCanvas.Children.IndexOf(overlay);
            if (idx < FloatingCanvas.Children.Count - 1)
            {
                FloatingCanvas.Children.RemoveAt(idx);
                FloatingCanvas.Children.Insert(idx + 1, overlay);
                model.ZOrder++;
                _viewModel?.NotifyFloatingObjectChanged();
            }
        };

        overlay.SendBackRequested += (_, _) =>
        {
            int idx = FloatingCanvas.Children.IndexOf(overlay);
            if (idx > 0)
            {
                FloatingCanvas.Children.RemoveAt(idx);
                FloatingCanvas.Children.Insert(idx - 1, overlay);
                model.ZOrder--;
                _viewModel?.NotifyFloatingObjectChanged();
            }
        };

        overlay.AppearanceChangedCommitted += (_, _) => _viewModel?.NotifyFloatingObjectChanged();

        overlay.GeometryChangedCommitted += (_, _) =>
        {
            // Canvas DIP → mm 동기화. NaN 방어 (SetLeft 직후라 정상이지만 안전).
            double left = Canvas.GetLeft(overlay); if (double.IsNaN(left)) left = 0;
            double top  = Canvas.GetTop(overlay);  if (double.IsNaN(top))  top  = 0;
            model.XMm      = left / TextBoxOverlay.DipsPerMm;
            model.YMm      = top  / TextBoxOverlay.DipsPerMm;
            model.WidthMm  = overlay.ActualWidth  / TextBoxOverlay.DipsPerMm;
            model.HeightMm = overlay.ActualHeight / TextBoxOverlay.DipsPerMm;
            model.Status   = NodeStatus.Modified;
            _viewModel?.NotifyFloatingObjectChanged();
        };

        overlay.ContentChangedCommitted += (_, _) => _viewModel?.NotifyFloatingObjectChanged();

        overlay.DeleteRequested += (_, _) =>
        {
            FloatingCanvas.Children.Remove(overlay);
            _viewModel?.RemoveFloatingObject(model);
            if (ReferenceEquals(_selectedOverlay, overlay)) _selectedOverlay = null;
            BodyEditor.Focus();
        };

        FloatingCanvas.Children.Add(overlay);
        return overlay;
    }

    private void SelectOverlay(TextBoxOverlay overlay)
    {
        if (!ReferenceEquals(_selectedOverlay, overlay) && _selectedOverlay is not null)
            _selectedOverlay.IsSelected = false;
        _selectedOverlay = overlay;
        overlay.IsSelected = true;
    }

    private void DeselectAllOverlays()
    {
        if (_selectedOverlay is null) return;
        _selectedOverlay.IsSelected = false;
        _selectedOverlay = null;
    }
}

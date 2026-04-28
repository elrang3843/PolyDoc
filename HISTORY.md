# 변경 이력 (Change History)

PolyDonky의 모든 의미 있는 변경 사항을 이 파일에 기록합니다.

이 문서는 [Keep a Changelog](https://keepachangelog.com/ko/1.1.0/) 규칙을 따르고,
버전 번호는 [Semantic Versioning](https://semver.org/lang/ko/) 을 따릅니다.

---

## 작성 규칙

- **변경이 발생하면 같은 PR/커밋에서 `## [Unreleased]` 섹션에 항목을 추가**합니다.
- 항목은 다음 카테고리로 분류합니다.
  - **Added** — 새 기능 추가
  - **Changed** — 기존 기능 동작 변경
  - **Deprecated** — 곧 제거될 기능 표시
  - **Removed** — 제거된 기능
  - **Fixed** — 버그 수정
  - **Security** — 보안 관련 수정
  - **Docs** — 문서 변경 (사용자에게 영향 있는 경우만)
  - **Internal** — 내부 리팩터링·빌드·CI 등 사용자 비가시 변경
- 한 줄로 *무엇이 바뀌었는지* 적고, 필요하면 괄호로 *왜* 또는 관련 이슈/PR 번호를 답니다.
  - 예: `- HWPX 표 셀 병합 import 지원 (#42)`
- 릴리스 시 `## [Unreleased]` 의 내용을 새 버전 헤더로 옮기고, 비어 있는 `[Unreleased]` 를 다시 만듭니다.
- 버전 헤더 형식: `## [1.0.0] - 2026-MM-DD` (정식), `## [1.0.0-test.1] - 2026-MM-DD` (테스트 빌드).
- 날짜는 `YYYY-MM-DD` (KST 기준).

### 버전 정책 (중요)

- **저장소의 모든 빌드는 사용자가 명시적으로 "릴리즈" 를 지시하기 전까지 테스트 버전으로 관리합니다.**
- **테스트 빌드 태그**: `1.0.0-test.<n>` (SemVer pre-release 식별자).
  - 예: `1.0.0-test.1`, `1.0.0-test.2`, ...
  - 테스트 빌드의 변경 내역도 동일한 헤더(`## [1.0.0-test.N] - YYYY-MM-DD`)로 기록할 수 있고, 별도 컷이 필요 없는 작은 변경은 `[Unreleased]` 에 누적합니다.
- **최초 정식 릴리스는 `1.0.0`** 입니다.
  - 사용자의 명시적 릴리스 지시(예: "릴리즈하자", "1.0.0 으로 컷하자")가 있을 때만,
    `[Unreleased]` 의 내용을 `## [1.0.0] - YYYY-MM-DD` 로 승격하고 git tag `v1.0.0` 을 생성합니다.
  - 자동화/AI/기여자가 임의로 정식 버전 헤더를 만들거나 `v1.0.0` 태그를 붙이지 않습니다.
- **1.0.0 이후**는 일반 [SemVer](https://semver.org/lang/ko/) 규칙(`1.0.1` / `1.1.0` / `2.0.0` ...)을 따릅니다.
- 정식·테스트 빌드 태그가 만들어지기 전, 순수 사양·문서 단계의 작업은 본 파일 하단의 `## [Pre-release]` 섹션에 날짜별로 누적 기록합니다.

---

## [Unreleased]

> 다음 릴리스에 들어갈 변경 사항을 여기에 기록합니다.

### Fixed
- **Fixed** — **인라인(본문 흐름) 모드 그림·도형이 멀티-선택 복사에서 누락되던 버그**. `CopyMultiSelectedToClipboard` 의 dedup 가드가 `ExtractCoreSelection` 결과에서 `ShapeObject`/`ImageBlock`/`Table` 을 모조리 스킵해, 오버레이 캔버스가 아닌 FlowDocument 내부 BlockUIContainer 로 존재하는 인라인 이미지/도형까지 함께 클립보드에서 떨어졌음. 수정: `_multiSelectedControls` 에서 수집한 Core 객체 참조를 `HashSet<object>(ReferenceEqualityComparer.Instance)` 로 기억해 두고, 본문 텍스트 선택 추출 시 **WPF Block.Tag 가 collectedRefs 와 동일한 참조** 인 경우만 스킵 — 인라인 그림·도형은 별도 Core 객체이므로 정상 포함된다. `ExtractCoreSelection` 호출이 아닌 동등 로직을 인라인으로 재구현해 `wpfBlock.Tag` 비교로 dedup.
- **Fixed** — **직선(Line) 그리기 직후 끝점에서 더블클릭 시 크래시**. 직선은 2점 입력이 완료되면 `ClickCount=1` 핸들러 내부에서 `FinishPolylineShape()`(→ `_drawingPolyline_active = false`)가 자동 호출된다. 이후 WPF가 같은 클릭 이벤트의 `ClickCount=2` 이벤트를 별도 라우팅으로 발생시킬 때 `_drawingPolyline_active == false` 상태에서 예기치 않은 핸들러 분기로 크래시가 발생했다. `_suppressNextClickAfterLineFinish` 플래그를 추가해 자동마감 직후의 ClickCount≥2 이벤트를 `PaperBorder.PreviewMouseLeftButtonDown` 진입 즉시 소비·차단하도록 수정.
- **Fixed** — **멀티-선택 복사 시 텍스트 상자가 클립보드에 포함되지 않던 버그**. `CopyMultiSelectedToClipboard`가 `_multiSelectedControls` 내 `TextBoxOverlay`를 `FloatingObjectsClipboardFormat`("PolyDonky.FloatingObjects.v1") 리스트 포맷으로 직렬화해 저장. 붙여넣기(`TryPasteSelectedObject`) 시 `FlowSelectionClipboardFormat`과 함께 `FloatingObjectsClipboardFormat`도 읽어 모든 글상자를 +5mm 오프셋으로 재생성.
- **Fixed** — **멀티-선택 복사 붙여넣기 시 도형·이미지가 원본 위치에 겹쳐 붙여넣어지던 버그**. `TryPasteFlowSelection` 내부에서 `InFrontOfText`/`BehindText` 모드인 `ShapeObject`·`ImageBlock`의 `OverlayXMm`/`OverlayYMm`에 +5mm 오프셋을 적용.
- **Fixed** — **Ctrl+클릭으로 글상자를 선택한 뒤 Ctrl+V 붙여넣기 시 글상자 자체가 아닌 내용만 본문에 삽입되던 버그**. `TryPasteSelectedObject`가 `FlowSelectionClipboardFormat` 우선 처리 후 `FloatingObjectsClipboardFormat`/`FloatingObjectClipboardFormat`을 순차 처리하도록 재구성. `TryPasteFloatingObject`의 과도한 `BodyEditor.Selection.IsEmpty` 가드를 제거해, 캐럿만 위치한 경우에도 글상자 클립보드 데이터가 최우선 적용되도록 수정.
- **Fixed** — **멀티-선택 복사 시 도형·이미지 개수가 증가(중복)되던 버그**. `CopyMultiSelectedToClipboard`에서 `ExtractCoreSelection()`이 반환한 항목 중 `ShapeObject`·`ImageBlock`·`Table`을 skip해 `_multiSelectedControls` 루프에서 이미 수집한 항목이 중복 직렬화되지 않도록 수정.

### Added
- **Added** — **도구 > 사전 — WebView2 미니 웹뷰어**. 네이버 국어사전·영한사전, 다음 사전, 표준국어대사전을 ComboBox로 선택해 사전 사이트를 임베드 브라우저(Microsoft.Web.WebView2)로 표시하는 비모달 플로팅 창 추가. 에디터에서 텍스트를 선택한 채 메뉴를 열면 선택 단어가 자동 검색어로 입력됨. 창 닫기는 숨김 처리(Hide)로 WebView2 세션 유지, 앱 종료 시 실제 해제.
- **Added** — **편집용지 페이지 구분선 시각화**. 본문 내용이 한 페이지 높이를 초과하면 `PageBreakCanvas`에 페이지 경계마다 파선 구분선과 "─── N페이지 ───" 레이블을 표시. `PaperBorder.SizeChanged` 이벤트로 내용 길이 변화 시 자동 갱신.

- **Fixed** — **오버레이 개체(글상자/도형/그림/표) Ctrl+클릭 멀티-선택이 작동하지 않던 버그 + Ctrl+A 통합 선택 추가**. 원인: 오버레이 컨트롤은 BodyEditor 의 형제(같은 Grid)이므로 BodyEditor.PreviewMouseLeftButtonDown 의 tunneling 경로에 들어오지 않아, 거기에 등록한 Ctrl+클릭 토글 핸들러가 절대 호출되지 않았음. 수정: Ctrl+클릭 핸들러를 PaperBorder.PreviewMouseLeftButtonDown(공통 부모) 으로 이동 — 모든 자식(BodyEditor + 오버레이 Canvas)을 포괄하는 tunneling 경로. 추가로 Ctrl+A 를 가로채 본문 텍스트 + 모든 오버레이를 한 번에 선택하는 SelectAllIncludingOverlays 구현.

### Added
- **Added** — **페이지 범위 마퀴(범위 드래그) 멀티-선택 + Ctrl+클릭 오버레이 토글**: 용지 여백 영역을 드래그하거나 Ctrl+드래그로 마퀴 사각형을 그리면 그 안의 텍스트 블록·오버레이 이미지/도형/표/글상자가 한꺼번에 선택됨. Ctrl+클릭으로 오버레이 개체를 선택에 추가/제거(토글). 선택된 개체들은 Ctrl+C/X(복사/잘라내기) 로 PolyDonky.FlowSelection.v1 + PolyDonky.FloatingObject.v1 포맷으로 클립보드 저장, Delete 로 일괄 삭제. Escape 로 선택 해제. 재구축(RebuildOverlayImages/Shapes/Tables/FloatingObjects) 시 stale 참조 방지.
- **Added** — **오버레이 표 드래그 이동 (F단계)**: InFrontOfText·BehindText·Fixed 배치 모드인 표를 캔버스 위에서 드래그해 위치를 변경할 수 있음. 마우스를 떼면 Core.Table.OverlayXMm/OverlayYMm 에 반영. Block 모드 표는 드래그 불가.
- **Added** — **표 열 너비 드래그 리사이즈 (E단계)**: FlowDocument 표의 열 경계선 위로 마우스를 이동하면 커서가 ↔(SizeWE)로 바뀌고, 드래그하면 인접 두 열의 너비가 실시간으로 조정됨. TextPointer 상향 탐색으로 TableCell 위치를 확인하고 GetCharacterRect 로 경계선 X 좌표를 추정. 드래그 완료 시 Core.Table.Columns[i].WidthMm 에 반영.
- **Added** — **멀티 셀 선택 우클릭 메뉴 (D단계)**: 여러 셀을 드래그 선택 후 우클릭하면 전용 컨텍스트 메뉴 표시. 선택 셀 병합(직사각형 범위), 선택 셀 내용 지우기, 선택 셀 배경색 일괄 변경(OS 색상 선택기), 선택 셀 속성 일괄 적용, 표 속성 진입 제공.
- **Added** — **표 배치 모드 (C단계)**: 표를 본문 흐름 외에 텍스트 위(InFrontOfText)·텍스트 아래(BehindText)·위치 고정(Fixed)으로 배치 가능. Core.TableWrapMode enum + OverlayXMm/OverlayYMm 모델 추가. 오버레이 모드일 때 Grid 기반 표 시각 요소를 OverlayTableCanvas/UnderlayTableCanvas 에 절대 위치로 렌더. 표 속성 다이얼로그에 배치 방식 섹션 추가.
- **Added** — **ColorPickerBox 공용 색상 입력 컨트롤**: 직접 입력(hex/이름) + OS 색상 선택기(`...` 버튼) + 없음(지우기) 버튼을 하나의 UserControl로 제공. 투명/없음 상태는 체커보드 패턴으로 표시. 셀 속성·도형 속성 다이얼로그에 적용해 기존 단순 TextBox 대체.
- **Added** — **표 삽입 기능 (1단계)**: 입력 → 표 메뉴에서 표 삽입 다이얼로그(행 수·열 수·너비·머리행 여부)를 통해 본문 커서 위치에 표를 삽입. 머리행(첫 행)은 회색 배경·굵은 글꼴로 표시.
- **Added** — **표·셀 속성 편집 (3단계)**: 표 우클릭 메뉴에서 "셀 속성" / "표 속성" 다이얼로그 제공. 셀 속성 — 배경색(헥스+스와치), 텍스트 정렬(왼쪽·가운데·오른쪽·양쪽 맞춤), 여백(상하좌우 mm), 테두리(두께 pt + 색상). 표 속성 — 수평 정렬(왼쪽·가운데·오른쪽). 변경 즉시 FlowDocument에 반영.
- **Added** — **표 구조 편집 (2단계)**: 표 우클릭 메뉴에서 행·열 삽입/삭제, 셀 병합(우측·아래), 셀 분할, 표 삭제 기능 제공. 핵심 모델(Core.Table)과 WPF FlowDocument 표를 동시에 업데이트해 별도 재빌드 없이 즉시 반영.

### Fixed
- **Fixed** — **본문 텍스트를 복사해 (글상자 등) RichTextBox 에 붙여넣으면 "System.Byte[]" 라는 문자열이 들어가던 버그**. 클립보드에 XamlPackage / RTF 를 넣을 때 `dataObj.SetData(format, byte[])` 로 byte 배열을 설정. 이 포맷들의 약속 타입은 각각 `Stream` / `string` 이라 받는 쪽이 byte[] 를 string 으로 변환하다 `byte[].ToString()` = `"System.Byte[]"` 를 텍스트로 받아들였음. XamlPackage 는 `MemoryStream` 그대로(dispose 하면 클립보드에서 못 읽으므로 using 제거), RTF 는 `Encoding.ASCII.GetString(bytes)` 로 string 변환해 SetData.
- **Fixed** — **단락 내 부분 텍스트 선택을 복사하면 단락 전체 + 가짜 줄바꿈이 붙여넣어지던 버그**. 두 가지 근본 수정: ① `FlowDocumentParser.ParseParagraphClipped` 추가 — 선택이 단락을 부분적으로 덮는 경우 선택 범위[clipStart, clipEnd] 내 글자만 추출하고, 범위 내 InlineUIContainer·Floater 등 내장 객체도 회피 없이 그대로 포함. `ExtractCoreSelection` 이 부분 선택 단락에는 `ParseParagraphClipped` 를 사용해 선택 글자·객체만 직렬화. ② `InsertParagraphInline` 추가 — 단일 단락 클립보드를 붙여넣을 때 새 Block 삽입 대신 캐럿 위치에 인라인 삽입. 캐럿이 Run 중간에 있으면 Run 을 분리하고 사이에 끼워 넣음. 다중 블록(표·이미지·도형) 은 기존 Block 단위 삽입 경로 유지.
- **Fixed** — **오버레이 표(InFrontOfText/BehindText/Fixed 배치 모드)를 복사-붙여넣기 하면 표가 사라지던 버그**. `BuildWpfBlockFromCore` 에서 `WrapMode != Block` 인 `Core.Table` 을 `null` 로 처리해 붙여넣기 시 드롭됐음. `FlowDocumentBuilder.BuildTableAnchor` 를 `internal` 로 승격하고, 비 Block 모드 표도 앵커 단락으로 삽입한 뒤 `RebuildOverlayTables()` 를 호출해 캔버스 시각 요소를 재구성. 오버레이 이미지·도형도 붙여넣기 후 `RebuildOverlayImages()` / `RebuildOverlayShapes()` 를 함께 호출해 모든 오버레이 객체의 캔버스 재구성을 보장.
- **Added** — **본문 RichTextBox 의 멀티-블록 복사/잘라내기/붙여넣기를 PolyDonky.FlowSelection.v1 커스텀 클립보드 포맷으로 전면 교체**. 기존: WPF XamlPackage 직렬화는 `Tag`(Core 객체 참조)를 보존하지 못해 표·이미지·도형이 붙여넣기 후 속성을 잃거나 저장 시 무음 삭제. 신규: `CommandManager.AddPreviewExecutedHandler` 로 Copy/Cut/Paste 를 가로채 ① 선택 영역의 모든 Block 을 `FlowDocumentParser.ParseSingleBlock` 으로 Core.Block 리스트로 추출 → JSON 직렬화 → 클립보드 커스텀 포맷에 저장, ② 붙여넣기 시 동일 포맷을 deserialize 후 `FlowDocumentBuilder.BuildTable/BuildImage/BuildShape/BuildParagraph` 로 WPF 트리를 재생성 — Tag·스타일·테두리·여백·정렬·표 구조·이미지 데이터 등 **모든 속성을 무손실로 라운드트립**. 외부 앱 호환을 위해 plain text / RTF / XamlPackage 도 함께 클립보드에 적재. ID 는 paste 시 재발급해 중복 방지. 기존 부유 객체 전용 포맷(`PolyDonky.Block.v1`, `PolyDonky.FloatingObject.v1`) 은 그대로 유지 (도형/글상자 단일 객체 복사 경로).
- **Fixed** — **이미지(BlockUIContainer/Paragraph+Floater)를 복사-붙여넣기 하면 저장 시 이미지가 무음 삭제되던 데이터 손실 버그**. 원인: WPF 클립보드 직렬화(XAML)는 `Tag`(임의 object)를 보존하지 않아, 붙여넣기 된 `BlockUIContainer` / `Paragraph` 의 Tag 가 null 로 초기화됨. `FlowDocumentParser.ParseInto` 의 switch 에서 `container.Tag is ImageBlock` / `Paragraph.Tag is ImageBlock` 조건이 실패해 어떤 case 에도 매칭되지 않고 조용히 드롭. 수정: ① `TryRecoverImageFromBUC` — BitmapSource(XamlPackage 포맷에 보존)에서 Inline 모드 `ImageBlock` 재구성. ② `TryRecoverImageFromPara` — Floater 포함 단락(WrapLeft/WrapRight)과 InlineUIContainer 단락(AsText)에서 이미지 재구성. ③ `BuildRecoveredImageBlock` — 크기·정렬·여백·테두리·PNG 인코딩. ④ `DataObject.AddPastingHandler` 등록 — 붙여넣기 직후 새 표에 `EnsureCoreTable` 일괄 적용. ※ 도형(ShapeObject) 은 Canvas 기하를 역산할 수 없어 이번 수정 범위 외 (붙여넣기 후 저장 시 유실은 기존 동작 유지).
- **Fixed** — **표를 복사-붙여넣기 한 뒤 우클릭하면 잘라내기/복사/붙여넣기 메뉴만 뜨고 표 편집 메뉴가 사라지던 버그**. 원인: WPF 클립보드 직렬화(XAML)는 `Wpf.Table.Tag` 에 부착된 `PolyDonky.Core.Table` 참조를 보존하지 못해, 붙여넣기 된 표는 Tag 가 null 인 채 FlowDocument 에 삽입됨. 표 우클릭 메뉴·열 너비 리사이즈·멀티 셀 검출이 모두 `Tag is PolyDonky.Core.Table` 검사를 단축 평가로 사용해 실패. 수정: WPF 표 구조(컬럼·행·셀·병합)를 미러링한 신선한 `Core.Table` 을 즉석에서 만들어 부착하는 `EnsureCoreTable` 헬퍼 추가, 위 세 진입점에서 호출. 저장 경로(`FlowDocumentParser.ParseTable`)도 Tag 가 없을 때 동일한 신선 Core.Table 시드를 사용하므로 라운드트립과 모순 없음.
- **Fixed** — **맨 오른쪽 열을 포함해 셀을 선택하면 표 편집 메뉴가 뜨지 않고 셀 수가 +1 로 잘못 집계되던 버그**. `Selection.End` 가 마지막 셀의 오른쪽 경계(또는 표 다음 Paragraph) 에 위치할 때 `FindAncestorCell` 이 null 을 반환하거나 다음 셀을 반환해 off-by-one 발생. `Selection.End` 를 한 위치 뒤로 물러난 지점(`GetPositionAtOffset(-1, Backward)`) 에서 셀을 찾도록 수정 — 항상 마지막으로 선택된 셀 안쪽에 위치하게 됨.
- **Fixed** — **표 셀 우클릭 메뉴가 선택 형태에 따라 다르게 뜨던 버그 + 우클릭 메뉴 빌드 단일화**. 가로로 여러 셀을 드래그 선택했을 때 multi-cell 검출이 실패해 잘라내기/복사/붙여넣기 기본 메뉴만 떴던 문제. 원인: `Selection.Start.Paragraph?.Parent is TableCell` 직접 검사 — Paragraph 가 List/Span 등으로 감싸여 있거나 선택 끝점이 셀 경계에 걸치면 매칭 실패. 수정: TextPointer 조상 체인을 끝까지 거슬러 올라가는 `FindAncestorCell`. 또한 BodyEditor 우클릭 메뉴를 XAML 정적 정의 + ContextMenuOpening 추가의 두 곳에서 빌드하던 구조를 ContextMenuOpening 한 곳으로 통합 — 매번 메뉴를 비우고 잘라내기/복사/붙여넣기부터 통째로 다시 빌드해 동일 선택 → 동일 메뉴 보장.
- **Fixed** — **직선(Line) 도형이 저장→불러오기 후 사각형으로 표시되던 버그**. `ShapeKind.Line = 0` 이 enum 의 default(0) 이라 `JsonIgnoreCondition.WhenWritingDefault` 정책에 의해 `"kind"` 필드가 JSON 에서 누락되었고, 역직렬화 시 ShapeObject.Kind 의 C# 기본값(Rectangle) 으로 복원되었음. `ShapeObject.Kind` 에 `[JsonIgnore(Condition = JsonIgnoreCondition.Never)]` 를 명시해 항상 직렬화하도록 수정.

### Changed
- **Changed** — **직선(Line) 입력 방식을 폴리곤선과 동일한 click-click 으로 통일**. 시작점 클릭 → 끝점 클릭 → 자동 마감. 기존 드래그 방식 제거. Polyline / Spline / Polygon / ClosedSpline 과 동일한 입력 코드 경로를 사용 — 코드 단순화 + UX 일관성.
- **Changed** — **부유 객체(도형·글상자) 의 키보드/클립보드 처리를 객체 타입 무관 핸들러로 일반화**. 기존 코드는 `_selectedOverlay` (TextBoxOverlay) 한정이었음 — 도형이나 다른 객체를 추가할 때마다 핸들러를 손봐야 했음. 이제 `TryDeleteSelectedObject` / `TryCopySelectedObject` / `TryCutSelectedObject` / `TryPasteSelectedObject` 가 선택된 객체 종류를 보고 자동 분기. **도형 선택**: 오버레이 도형 클릭 시 시각 표시(드롭섀도 글로우) + 키보드 라우팅. **Delete 키**: 선택된 부유 객체가 있으면 삭제, 없으면 본문 텍스트 삭제로 양보. **Ctrl+C/X/V**: 도형은 `BlockJsonConverter` 를 통해 `PolyDonky.Block.v1` 클립보드 포맷으로 직렬화 — 향후 다른 Block 타입(이미지 등) 도 같은 포맷으로 자동 처리 가능.

### Fixed
- **Fixed** — **저장→불러오기 후 모든 도형이 사각형으로 표시되던 버그**. `BlockJsonConverter.Write()` 에서 레거시 `"kind"` discriminator 제거를 위해 `"kind"` 이름의 JSON 필드를 전부 스킵했는데, `ShapeObject.Kind` 와 `OpaqueBlock.Kind` 도 camelCase 직렬화 시 `"kind"` 가 되어 함께 누락됨. `"$type"` 중복만 방지하도록 수정 (읽기 경로에서 `"$type"` 이 `"kind"` 보다 우선하므로 레거시 호환 유지).
- **Fixed** — **그림 "본문 흐름" 모드에서 정렬(왼쪽/가운데/오른쪽)을 바꿔도 위치가 변하지 않던 버그**. `BlockUIContainer.TextAlignment`는 텍스트 Glyph 정렬이며 UIElement 위치에 무효. 명시적 `Width`가 있는 `Image`의 기본 `HorizontalAlignment`(Stretch)가 WPF 레이아웃에서 중앙 배치처럼 동작하므로, 정렬을 Left로 바꿔도 이미지가 항상 가운데에 그려졌다. 이미지(Image) 및 테두리 래퍼(Border) 생성 시 `HorizontalAlignment = imgHA`를 명시적으로 설정해 수정.
- **Fixed** — **그림 속성 다이얼로그에서 "오른쪽 배치" 선택 시 그림 사라지던 버그** (근본 원인 수정). 기존 코드는 `FlowDocument.PageWidth = 종이 전체 폭`으로 설정하고 `BodyEditor.Padding`으로 좌우 여백을 별도 적용했다. 이로 인해 FlowDocument 오른쪽 끝이 가시 영역 밖(`padLeft + padRight` 만큼)으로 밀려, `HorizontalAlignment.Right` Floater를 비롯한 모든 우측 정렬 객체가 클리핑되어 사라졌다. `FlowDocumentBuilder.ComputeContentWidthDip` 헬퍼를 추가하고 `PageWidth = 종이 폭 − 좌여백 − 우여백(본문 폭)`으로 수정. 여백 변경 시에도 `ApplyPageSettings`가 `Document.PageWidth`를 즉시 갱신. anchor `Run`의 문자도 빈 문자열 → 비줄바꿈 공백(U+00A0)으로 교체해 WPF TextFormatter가 라인을 유효하게 처리하도록 보강.
- **Fixed** — **그림을 "텍스트 앞으로/뒤로" 모드로 처음 전환 시 페이지 좌상단(0,0) 으로 점프하던 버그**. 모드 전환 직전 `imgControl.TransformToVisual(PaperBorder)` 로 현재 화면 좌표를 캡처하고, 전환 후 `OverlayXMm/OverlayYMm` 가 기본값(0,0) 인 경우에만 캡처한 좌표를 mm 로 변환해 적용 — 사용자가 의도적으로 0,0 을 입력한 경우와 충돌하지 않는다.
- **Fixed** — **"텍스트 뒤로" 그림이 BodyEditor 뒤에 있어 우클릭·드래그·속성 변경 모두 불가능하던 버그**. UnderlayImageCanvas 자식은 BodyEditor 가 마우스를 가로채 직접 hit-test 가 불가능. 두 가지 라우팅 추가: 1) **우클릭** — `BodyEditor.ContextMenuOpening` 에서 마우스 좌표를 underlay canvas bbox 와 비교해 그림 발견 시 컨텍스트 메뉴에 "그림 속성" 항목 동적 추가 → 속성 다이얼로그로 모드 변경/위치 조정 가능. 2) **Alt+드래그** — `Alt` 키와 함께 클릭하면 underlay 그림 드래그 시작; BodyEditor 가 mouse capture 를 가지고 캔버스 좌표 변화량을 underlay 그림에 직접 라우팅. 일반 클릭은 본문 텍스트 선택에 양보해 텍스트 편집과 그림 조작이 충돌하지 않는다.
- **Fixed** — **페이지 설정 변경 / 개요 스타일 변경 시 저장 이후 편집한 이미지·글상자가 사라지던 객체 손실 버그** 2건. 공통 원인: `RebuildFlowDocument()` 가 저장 시점의 stale `_document` 모델로 FlowDocument 를 통째로 재구성하면서 그 이후 삽입·편집된 모든 블록을 덮어씌웠음. 수정: 1) **페이지 설정** — FlowDocument 내용은 용지 크기·여백·배경색과 무관하므로 `RebuildFlowDocument()` 를 완전히 제거하고 `ApplyPageSettings()` 로 레이아웃 속성만 갱신. 2) **개요 스타일** — 재빌드 전에 `FlowDocumentParser.Parse(liveDoc, _document)` 로 live FlowDocument 를 `_document` 에 먼저 동기화한 뒤 재빌드 — 이미지·글상자가 sync 된 상태로 보존된 채 스타일만 업데이트됨. `RebuildFlowDocument()` 는 `private` 으로 격리해 외부에서 직접 호출 불가.
- **Fixed** — **WrapLeft/WrapRight 그림 드래그 시 속성(ImageBlock 모델 참조) 이 사라지고 위치 제어 불가하던 핵심 버그**. 원인: WPF Floater 드래그가 클립보드 직렬화를 경유해 `Tag`(ImageBlock 참조)를 통째로 지워버렸음. `PreviewMouseMove e.Handled=true` 로 WPF 드래그를 차단하고 독자 드래그 로직 구현: MouseDown 시 `Block` + `ImageBlock` 참조를 직접 캡처 → MouseMove 에서 8px 임계거리 초과 시 커서를 SizeAll 로 변경 → MouseUp 에서 드롭 X 좌표를 3등분해 WrapLeft(왼쪽 1/3) / Inline 가운데(중앙) / WrapRight(오른쪽 1/3) 로 `WrapMode`·`HAlign` 갱신 후 Block 을 재구축. 기존 `InsertBefore/Remove` 패턴으로 모델 참조가 단 한 번도 클립보드를 통하지 않으므로 이후 우클릭 속성 편집도 정상. Inline 모드 그림은 드래그 후 `HAlign` 만 변경 (WrapMode 유지).
- **Fixed** — 본문에서 우클릭 시 `System.InvalidOperationException: 'FlowDocument' is not a Visual or Visual3D` 크래시. `TryWalkUpToImage` 의 비주얼 트리 탐색 루프가 `BodyEditor.InputHitTest` 가 반환한 `ContentElement`(Run·Paragraph·FlowDocument 등 — Visual 이 아님) 에 대해 `VisualTreeHelper.GetParent` 를 호출해 예외 발생. Visual 여부를 먼저 확인해 non-Visual 이면 즉시 null 반환하도록 수정.
- **Fixed** — 이모지·그림 위에서 우클릭/더블클릭해도 **속성 다이얼로그가 열리지 않고 RichTextBox 기본 메뉴(잘라내기/복사/붙여넣기) 만 뜨던 버그**. 두 가지 원인을 함께 수정: 1) `PreviewMouseRightButtonDown` 으로 e.Handled = true 를 세팅해도 RichTextBox 의 기본 컨텍스트 메뉴는 `ContextMenuOpening` 시점에 결정되므로 막을 수 없었음 — 핸들러를 `ContextMenuOpening` 로 옮기고 그 시점에 e.Handled = true 로 기본 메뉴 억제. 2) `e.OriginalSource` 가 RichTextBox 내부 호스팅으로 Image 까지 닿지 않는 경우가 있었음 — `BodyEditor.InputHitTest(마우스위치)` 폴백 추가해 두 경로 모두에서 Image 를 찾도록 보강.

### Added
- **Added** — **표 속성·셀 속성 편집** (3단계). 표 안에서 우클릭 시 컨텍스트 메뉴에 "셀 속성(_C)..." / "표 속성(_T)..." 항목 자동 추가. 셀 속성: 배경색(hex/이름)·텍스트 정렬(왼쪽/가운데/오른쪽/양쪽 맞춤)·여백(상하좌우 mm)·테두리(두께 pt + 색상); 확인 즉시 WPF 미리보기에 반영. 표 속성: 수평 정렬(왼쪽/가운데/오른쪽)은 모델에 저장(HWPX/DOCX 변환 시 적용). `CellTextAlign` / `TableHAlign` 열거형 + `TableCell.BackgroundColor` / `TextAlign` / `PaddingMm` / `BorderColor` / `BorderThicknessPt` / `Table.HAlign` 모델 확장. `FlowDocumentBuilder.ApplyCellPropertiesToWpf` 공유 헬퍼로 라운드트립 시 셀 속성 완전 보존.
- **Added** — **표 삽입** (입력 > 표 메뉴 활성화). 행 수·열 수·너비(mm)·머리글 행 여부를 지정해 표를 캐럿 위치에 삽입하는 `TableInsertDialog` 추가. 셀 텍스트 편집은 WPF RichTextBox 기본 편집으로 즉시 가능. `TableRow.IsHeader` 모델 프로퍼티 추가 — 머리글 행은 배경색·세미볼드로 강조 렌더링. `FlowDocumentParser.ParseTable` 이 `IsHeader` 를 라운드트립 시 보존하도록 수정.
- **Added** — **폴리곤면(Polygon)·스플라인면(ClosedSpline) 입력 추가**. 폴리곤선/스플라인선과 동일한 클릭 방식이되 마지막에 시작점으로 닫히는 닫힌 면 도형. 최소 3점 필요. 채우기 색상 기본 적용(다른 닫힌 도형과 동일). 닫힌 스플라인은 wrap-around Catmull-Rom Bezier로 이음매 없이 매끄럽게 닫힘.
- **Added** — **폴리곤선·스플라인선 다중 점 클릭 입력 모드**. 기존 드래그(2점 고정) 대신 클릭마다 점을 추가하고 더블클릭 또는 우클릭 메뉴(완료 / 마지막 점 취소 / 시작점에 닫기 / 전체 취소) 또는 Enter 키로 마침. 입력 중 마우스 위치까지 점선 고무줄 미리보기(DrawPreviewCanvas)를 실시간 표시. Esc 로 전체 취소.
- **Added** — **도형 입력 구현** (`ShapeObject`). 입력 > 도형 메뉴에 직선·폴리곤선·스플라인선·사각형·둥근사각형·타원·삼각형·정다각형·별 9종 추가. 그림과 동일한 5모드 배치 체계(`ImageWrapMode`) 공유 — 본문 흐름(블록), 왼쪽/오른쪽 배치(텍스트 감싸기), 텍스트 위/뒤(캔버스 오버레이). 레이블 텍스트 입력 지원. 오버레이 도형은 드래그 이동·더블클릭 속성창 지원. 도형 속성 다이얼로그(`ShapePropertiesWindow`)로 선/채우기/크기/위치/레이블 편집. `ShapeObject : Block` 코어 모델, `BlockJsonConverter` `"shape"` discriminator, `FlowDocumentBuilder.BuildShape()` 추가.
- **Added** — **그림을 텍스트처럼 처리** (`ImageWrapMode.AsText`). 그림 속성 → 배치에 6번째 옵션 '텍스트로 처리 — 글자처럼 한 줄에 인라인' 추가. `BlockUIContainer` 대신 `Paragraph` 안의 `InlineUIContainer` 로 빌드되어 그림이 본문 텍스트와 같은 줄에 글자처럼 흐른다 (사용자가 같은 단락에 텍스트를 추가하면 그림 옆/뒤에 텍스트가 나란히 배치). FlowDocumentParser 가 `Paragraph.Tag = ImageBlock` 케이스를 이미 인식하므로 라운드트립 자동 보장.
- **Added** — **오버레이 그림(InFrontOfText / BehindText) 드래그 이동**. 캔버스 위 그림을 마우스로 끌어 위치를 바꿀 수 있다 — 드래그 종료 시 `OverlayXMm`/`OverlayYMm` 가 업데이트되고 dirty 마킹. 단일 클릭 = 드래그 시작, 더블클릭 = 속성 다이얼로그, 우클릭 = 속성 컨텍스트 메뉴로 분기. 글상자(`TextBoxOverlay`) 는 기존 드래그 동작 유지. `FlowDocumentBuilder` 에 `DipToMm` 헬퍼 추가.
- **Added** — **그림 텍스트 배치 모드 5종**. ImageBlock 에 `WrapMode` enum 신설 (Inline / WrapLeft / WrapRight / InFrontOfText / BehindText). 그림 속성 다이얼로그에 '배치' 콤보박스 추가:
  1. **본문 흐름** — BlockUIContainer (자체 줄, 가로 정렬만 적용; 기본값)
  2. **왼쪽 배치** — Paragraph + Floater(Left), 텍스트가 오른쪽으로 흐름
  3. **오른쪽 배치** — Paragraph + Floater(Right), 텍스트가 왼쪽으로 흐름
  4. **텍스트 앞으로** — `OverlayImageCanvas` (BodyEditor 위)에 절대 위치(`OverlayXMm`/`OverlayYMm`) 로 배치, 본문 위에 겹침
  5. **텍스트 뒤로** — `UnderlayImageCanvas` (BodyEditor 뒤)에 절대 위치로 배치, 본문 아래에 깔림
  오버레이 모드 그림은 본문 흐름에는 빈 placeholder Paragraph(Tag=ImageBlock) 만 두고 실제 렌더링은 캔버스에서 처리. 모드 전환 시 컨테이너 종류가 BUC↔Paragraph↔Placeholder 로 자유롭게 바뀌며, FlowDocumentBuilder.BuildImage 가 적절한 Block 타입을 반환해 일괄 교체. 오버레이 그림은 우클릭/더블클릭으로 속성 편집 가능 (X/Y 좌표 조정 포함). FlowDocumentParser 도 Paragraph.Tag = ImageBlock 케이스 인식해 라운드트립 보장.
- **Added** — **이모지·그림 속성 편집** (`EmojiPropertiesWindow`, `ImagePropertiesWindow`). 이모지/그림 우클릭 → 속성 메뉴, 더블클릭 → 속성 다이얼로그 직접 열기. 이모지: 크기(pt) + 기준선 정렬(위/가운데/아래/기준선) — 변경 즉시 Image 컨트롤에 반영, `Run.EmojiAlignment` 에 저장해 라운드트립 보장. 그림: 크기(mm·비율 유지), 정렬(왼쪽/가운데/오른쪽), 위아래 여백(mm), 테두리(색상·두께 pt), 설명 — 확인 시 `ImageBlock` 갱신 후 `BlockUIContainer` 재빌드해 FlowDocument 교체. `ImageBlock` 에 `HAlign`/`MarginTopMm`/`MarginBottomMm`/`BorderColor`/`BorderThicknessPt` 추가; `Run` 에 `EmojiAlignment` 추가.
- **Added** — **그림(이미지) 삽입** 기능 구현. 입력 → 그림 메뉴로 `ImageWindow` 다이얼로그 열기: 파일 선택(PNG/JPEG/BMP/GIF/TIFF/WebP), 미리보기, 너비·높이(mm) 입력(비율 유지 옵션), 설명(alt-text) 입력. 삽입 시 SHA-256 해시를 자동 계산해 `ImageBlock` 에 기록, 캐럿 위치 블록 직후에 `BlockUIContainer`(`Tag = ImageBlock`) 로 삽입 — `FlowDocumentParser` 의 기존 라운드트립 경로로 IWPF 저장/재로드 정상 동작. A4 본문 너비(160mm) 초과 이미지는 자동 축소. `FlowDocumentBuilder.BuildImage` 를 `internal` 로 공개.
- **Added** — 이모지 삽입 다이얼로그에 **크기 선택 콤보박스** 추가 (12/16/20/24/32/48pt, 기본 16pt). 선택한 크기를 `Run.Style.FontSizePt` 에 저장해 저장→재로드 후에도 동일 크기로 복원(라운드트립 보장). `FlowDocumentBuilder.BuildEmojiInline` 의 보정 계수(`× 1.4`) 를 제거해 삽입 크기와 로드 크기를 일치시킴.
- **Added** — **이모지(Emoji) 입력** 활성화. 8개 섹션 × 10개 = 80개의 문서용 PNG 이모지(상태/반응/동작/우선순위/사람/도구/화살표/장식)를 카테고리별 또는 키워드 검색으로 골라 캐럿 위치에 삽입. 라운드트립용으로 `Run.EmojiKey` 를 신설 (수식 `LatexSource` 와 동일 패턴) — `FlowDocumentBuilder.BuildEmojiInline` 이 pack URI 로 PNG 를 로드해 `InlineUIContainer { Image }` 로 렌더링, `FlowDocumentParser` 가 `Tag` 의 EmojiKey 를 보존. PNG 80개는 `..\..\Resources\Emojis\**\*.png` 글롭으로 WPF Resource 임베드, 빌드 산출물에 포함. 메뉴 라우팅은 `GetActiveTextEditor()` 를 통해 글상자 안쪽 편집 중에도 정상 동작.
- **Added** — 글상자 모양에 **타원(Ellipse)** 과 **파이(Pie/부채꼴)** 추가. 타원은 박스 비율에 따라 자동 스케일되는 4-Bézier 근사 타원; 파이는 시작 각도(°)·호 범위(5~355°) 를 글상자 속성 창에서 조정 가능한 부채꼴 모양. 두 모양 모두 인셋 자동 계산 적용 (타원 15%, 파이 10%). **입력 → 글상자** 메뉴와 `InsertTextBox` 커맨드 디스패처에도 두 모양 항목 추가 (타원 _E_, 파이 _I_).

### Fixed
- **Fixed** — **글자 속성 다이얼로그가 로드/복사된 글상자에서 색·폰트·볼드 등 변경을 시각적으로 반영하지 않던 핵심 버그**. `CharFormatWindow.ApplyToSelection` 의 마지막 단계인 `ApplyTypographicProps` 가 인라인의 `Tag` 가 가리키는 stale `PolyDonky.Run.Style` 로 `BuildInline` → `ReplaceInline` 을 수행하면서, 직전 `ApplyPropertyValue` 가 Wpf.Run 에 적용한 색·폰트·볼드 결과를 통째로 덮어 씌우고 있었다. 새로 그린 글상자는 WPF 가 자연 생성한 `Wpf.Run` 에 Tag 가 없어 `ExtractPolyRun(r)` 경로로 가서 현재 Wpf 속성을 정확히 추출 → 회귀가 발생하지 않았다. 두 단계 보강: 1) 글자폭/자간이 기본값(100%, 0px) 이면 `ApplyTypographicProps` 가 즉시 return — 재구성 자체를 건너뜀. 2) `Run` 케이스는 Tag 유무와 관계없이 `ExtractPolyRun(r)` 으로 현재 Wpf 속성을 우선 추출 — Tag.pr 의 stale 색·폰트가 ApplyPropertyValue 결과를 가리지 않도록.
- **Fixed** — 메뉴/우클릭 **서식 → 글자 속성/문단 속성** 다이얼로그가 글상자 안쪽 selection 이 비어 있을 때 시각적 변화를 만들지 않던 문제.
- **Fixed** — `GetActiveTextEditor()` 우선순위를 `_selectedOverlay?.InnerEditor → _lastTextEditor → BodyEditor` 로 강화. 사용자가 chrome 만 클릭해 글상자를 선택했을 때 (안쪽 본문에 한 번도 포커스가 들어간 적 없음) 도 정확히 그 글상자의 InnerEditor 로 포맷 다이얼로그를 라우팅. 이전 구현은 `_lastTextEditor` 만 보아, chrome-only 클릭 케이스에서 `BodyEditor` 로 폴백하던 회귀.
- **Fixed** — 메뉴 **서식 → 글자 속성 / 문단 속성** 을 글상자 편집 중에 열어도 본문(`BodyEditor`) 에만 적용되던 버그(2차 수정). 1차 수정에서 `IsKeyboardFocusWithin` 검사를 추가했지만, 메뉴를 클릭하는 순간 포커스가 메뉴로 이동하면서 InnerEditor 의 `IsKeyboardFocusWithin` 이 false 로 떨어져 여전히 `BodyEditor` 로 폴백하던 회귀. `BodyEditor.GotKeyboardFocus` / 각 `TextBoxOverlay.InnerEditor.GotKeyboardFocus` 에 훅을 걸어 **마지막으로 키보드 포커스를 가졌던 RichTextBox** 를 `_lastTextEditor` 에 추적하고, `GetActiveTextEditor()` 가 이 값을 우선 사용하도록 변경. 메뉴 클릭으로 인한 일시적 포커스 이전이 있어도 직전 편집 컨텍스트가 보존된다. 글상자 삭제·문서 재로드 시 stale 참조는 자동 정리.
- **Fixed** — `TryPasteFloatingObject` 가 `BodyEditor.IsKeyboardFocusWithin` 단독 검사로 본문 편집기에 캐럿이 있는 모든 경우 일반 텍스트 붙여넣기에 양보 → 사용자가 글상자를 복사한 직후 Ctrl+V 를 눌러도 plain-text fallback 만 본문에 들어가던 문제. **텍스트 선택이 있을 때만** 일반 붙여넣기에 양보하도록 수정 — 캐럿만 위치한 경우(BodyEditor / InnerEditor 무관) 는 글상자 클립보드 데이터를 우선 적용해 새 글상자 한 개를 캔버스에 띄운다.

- **Fixed** — 글상자 안쪽 텍스트의 **글자 속성(폰트·크기·볼드·이탤릭·밑줄·글자색·배경색 등) 이 저장→재로드 시 사라지고, 로드 후 글자 속성 변경이 반영되지 않던 버그**. `TextBoxOverlay.LoadModelTextToEditor` 가 만드는 `Wpf.Run` 에 원본 `PolyDonky.Run` 을 `Tag` 로 심어, `FlowDocumentParser.ParseInline` 이 Tag 우선 시드(`Clone(original.Style)`) 로 라운드트립 — WPF 의 inheritance 로 인한 속성 drift(폰트 패밀리 자동 stamping 등) 와 비-Wpf 속성(LetterSpacingPx, WidthPercent) 까지 정확히 복원. 더불어 본문 라운드트립용 `FlowDocumentBuilder.BuildInline(Run)` / `FlowDocumentParser.Parse(FlowDocument)` 재사용 (이전 plain-text 전용 변환 → 글자 속성 통째 유실 문제도 해결).
- **Fixed** — 글상자 chrome 을 선택한 뒤 **Ctrl+C / Ctrl+X / Ctrl+V** 가 안쪽 본문(InnerEditor)에 캐럿이 있다는 이유로 글상자가 아닌 안쪽 텍스트만 처리하던 문제. 안쪽에 **텍스트 선택이 있을 때만** 일반 텍스트 클립보드에 양보하고, 선택이 비어 있는 경우(=캐럿만 위치) 는 글상자 자체 복사/잘라내기/붙여넣기로 처리. Word/PowerPoint 와 동일한 mental model.
- **Added/UX** — 글상자 안쪽 본문 편집 중 **Esc** 동작을 2단계로 변경: 1) 안쪽 편집 중 → chrome 만 선택된 상태로 전환(글상자 자체에 포커스), 2) chrome 선택 상태 → 완전 해제. 1단계에서 바로 Ctrl+C 로 글상자 통째 복사가 가능 — 사용자 발견성 개선.
- **Fixed** — 글상자 안쪽에 **불필요한 세로 스크롤바**가 표시되던 문제. `RichTextBox.VerticalScrollBarVisibility` 를 `Auto` → `Hidden` 으로 변경. 글상자는 그래픽 객체이므로 내용이 박스보다 길어지면 스크롤이 아니라 박스를 키워야 한다 (Word·PowerPoint 와 동일한 UX).
- **Fixed** — 비사각형 글상자(말풍선·구름·가시별·번개) 에서 안쪽 텍스트가 **외곽 돌출부(꼬리·뭉게·가시) 영역까지 침범**하던 문제. `ComputeShapeInset()` 이 박스 크기에 비례해 꼬리·돌출부 영역을 보호하는 추가 인셋을 계산 (말풍선 꼬리 15%, 구름 10%, 가시별 22%, 번개 15/20%). `SizeChanged` 시 자동 재계산. 사용자 padding 위에 더해진다.
- **Fixed** — 글상자(부유 객체) 가 저장 후 다시 열면 사라지던 데이터 손실 버그. `FlowDocumentParser.Parse` 가 새 `Section` 을 만들면서 원본 섹션의 `FloatingObjects` 컬렉션을 인계하지 않아 발생. 글상자는 본문 흐름(`Section.Blocks`) 과 별도 레이어이므로 FlowDocument 에서 파싱되지 않는데, 저장 직전 rebuilt 문서에 누락되면 IWPF 직렬화 단계에서 통째로 사라졌다. `originalForMerge.Sections[0].FloatingObjects` 를 새 섹션으로 복사하도록 수정. 같은 경로에서 `Watermark` 와 `OutlineStyles` (문서 수준 상태) 도 누락되던 것을 함께 인계하도록 보정.

### Changed
- **Changed** — 새 글상자의 **기본 텍스트 정렬을 좌상단 → 가운데 정렬**(가로 Center / 세로 Middle) 로 변경. 글상자(특히 말풍선·구름·가시 등 비사각형 모양) 의 일반적 사용 패턴은 짧은 텍스트를 박스 한가운데 두는 것 — Word · PowerPoint 의 기본 텍스트 상자 동작과 일치. **마이그레이션 주의**: 이전에 기본값(Top/Left) 으로 저장된 IWPF 파일은 JSON 에 `hAlign`/`vAlign` 필드가 생략되므로, 새 코드로 다시 열면 가운데 정렬로 표시된다 (pre-1.0 단계라 허용).

### Added
- **Added** — 글상자(부유 객체) 자체의 복사/잘라내기/붙여넣기 지원. 글상자 chrome 이 선택된 상태(안쪽 본문 편집 중이 아님)에서 **Ctrl+C / Ctrl+X / Ctrl+V** 를 누르면 `TextBoxObject` 전체(모양·여백·정렬·색·내용 포함) 를 사용자 정의 클립보드 포맷 `PolyDonky.FloatingObject.v1` 로 직렬화/복원. 붙여넣기 시 위치를 (+5mm, +5mm) 오프셋해 원본 위에 겹치지 않게 복제. 안쪽 본문 또는 본문 편집기에 포커스가 있으면 가로채지 않고 일반 텍스트 클립보드 동작에 양보. plain-text 폴백을 함께 실어 다른 앱으로의 붙여넣기 시 안쪽 텍스트만 전달.

### Docs
- **Docs** — `PolyDonky` 작명 유래(**Poly**(gon) + **Donky**(당나귀): 다각형으로 거칠게 빚은 당나귀처럼 외형은 엉성해도 어떤 문서 포맷이든 가리지 않고 먹어치운다는 뜻) 를 한국어·영어 병기로 명문화. `README.md` 에 신규 섹션 "이름의 유래 (Name origin)" 추가 + 상단 인트로 직하단에 한·영 한 줄 요약. `CLAUDE.md` 의 프로젝트 개요 인용 블록으로 추가 — 향후 세션의 코드/UI 톤 일관성 가이드. `AboutWindow` 에 "Poly(gon) + Donky(당나귀)" 헤더 + 한·영 한 줄 설명을 새 섹션으로 노출.

### Changed
- **Changed** — 앱 아이콘 및 About 다이얼로그 히어로 이미지를 신규 PolyDonky 브랜드 아트(`assets/PolyDonky.jpg` 원본 → 자동 크롭/멀티사이즈 ICO)로 교체. `assets/PolyDonky.ico`(16/24/32/48/64/128/256 멀티사이즈) 와 `assets/PolyDonky_1024.png`(1024×1024 RGBA) 추가. `PolyDonky.App.csproj` 의 `<ApplicationIcon>` 을 `Handtech.ico` → `PolyDonky.ico` 로 전환. `MainWindow` 와 `AboutWindow` 에 `Icon="pack://application:,,,/Assets/PolyDonky.ico"` 명시 — 작업 표시줄·Alt+Tab·다이얼로그 타이틀바 모두 동일 아이콘으로 통일. About 다이얼로그의 140×140 히어로 이미지가 `Handtech_1024.png` → `PolyDonky_1024.png` 로 변경되어 앱 아이덴티티(IWPF + 도키) 가 강조됨. 회사 로고(`Handtech_1024.png`)는 리소스로 유지 — 추후 다른 위치(예: 라이선스 다이얼로그)에서 재사용 가능.

### Internal
- **Internal** — 프로젝트 전체 이름을 `PolyDoc` → `PolyDonky` 로 변경. 네임스페이스, 어셈블리명, `.csproj` 파일명, 솔루션 파일명, 소스 디렉터리명 일괄 변경. GitHub 레포지토리 이름은 사용자가 직접 Settings → General → Repository name 에서 변경 필요.

### Added
- **Added** — 입력 > **특수문자** 다이얼로그(`SpecialCharWindow`). 13개 카테고리(자주 쓰는 / 라틴 보충 / 그리스 / 키릴 / 화살표 / 수학기호 / 통화 / 숫자·위첨자 / 도형 / 표시 / 한글 자모 / 괄호·구두점 / 박스 그리기). 글자 그리드 클릭 시 미리보기(48pt) + 유니코드 코드포인트(`U+XXXX`) + Unicode 블록명 표시. 더블클릭 시 즉시 삽입. 유니코드 hex 직접 입력으로 임의 코드포인트 검색·삽입 가능(예: `03B1`, `2603`, `AC00`). 본문 캐럿 위치(또는 선택 영역 대체)에 단일 문자 삽입.
- **Added** — 입력 > **수식** 다이얼로그(`EquationWindow`). LaTeX 소스 입력 + 인라인(`\(...\)`) / 별행(`\[...\]`) 모드 토글 + 빠른 삽입 팔레트 16개(α/β/γ/π/Σ/∏/∫/√/\frac/위첨자/아래첨자/≤/≥/≠/∞/→). 빠른 삽입 시 `{}` 그룹 안으로 자동 캐럿 이동. 미리보기는 **WpfMath FormulaControl** 로 실시간 LaTeX 렌더링(파싱 오류 시 붉은 메시지 표시). 삽입 시 FormulaControl 을 오프스크린(`RenderTargetBitmap`) 렌더링 후 `Image`로 변환해 `InlineUIContainer`로 삽입 — `XamlWriter.Save()` undo 직렬화 충돌 방지. `Run.LatexSource` / `Run.IsDisplayEquation` 필드로 IWPF 라운드트립 보장.
- **Fixed** — 수식 `InlineUIContainer` 포함 선택 영역에 글자 속성 적용 시 발생하던 `XamlParseException`(`AdornedElementPlaceholder는 Template의 일부로서만 사용할 수 있습니다`) 수정. WPF undo 시스템이 `XamlWriter.Save()` 호출 시 `FormulaControl` 템플릿의 `AdornedElementPlaceholder`에서 crash — `Image(BitmapSource)` 로 교체해 해결. 글자 속성 창에서 수식 IUC 를 글자폭/자간 재구성 대상에서 제외하는 안전 가드 추가.

### Security
- **Security** — 라이선스를 Apache 2.0 에서 **Mozilla Public License 2.0** 으로 전환. 파일 수준 약(弱) 카피레프트로 PolyDonky 파일 자체의 수정 사항을 공유하도록 보장하면서 독점 확장 추가는 허용.
- **Security** — `THIRD_PARTY_NOTICES.md` 신규 생성. 배포 포함 의존성(CommunityToolkit.Mvvm MIT / DocumentFormat.OpenXml MIT / Markdig BSD-2-Clause / .NET 10 MIT) 전문(全文) 라이선스 텍스트 포함. 테스트 전용 의존성(xUnit Apache-2.0 / coverlet MIT)은 별도 섹션으로 구분.
- **Security** — `NOTICE` 파일 갱신. MPL-2.0 참조로 업데이트, 실제 배포 의존성 목록 반영, 파일 형식 명세 출처 섹션 추가(ECMA-376, [MS-DOC], OWPML/KS X 6101, HWP 5.0 공개 문서 — 한컴 권장 표기 문구 포함), 상표 섹션에 Hancom Inc.·Microsoft Corporation 명시.

### Added
- **Added** — 도움말 > **라이선스 및 참조** 다이얼로그 신규 추가(`LicenseInfoWindow`). 탭 4개: ① PolyDonky 라이선스(MPL-2.0 전문), ② 서드파티 라이선스 전문, ③ 의존성 목록(배포 포함/테스트 전용 분리), ④ 파일 형식 명세 참조(ECMA-376/[MS-DOC]/OWPML·KS X 6101/HWP 5.0). LICENSE 및 THIRD_PARTY_NOTICES.md 가 어셈블리 임베드 리소스로 내장되며, 개발 환경에서는 저장소 루트 파일로 폴백.

### Changed
- **Changed** — 글자 방향(세로쓰기 / 왼쪽으로 진행) 기능을 "추후 지원 예정"으로 전환. 페이지 서식 및 글상자 속성의 글자 방향 콤보박스를 비활성화(Opacity 0.45, IsEnabled=False)하여 선택 불가 상태로 표시. 내부적으로는 항상 LTR 가로쓰기로 동작하며, 모델의 `TextOrientation`/`TextProgression` 필드는 값을 보존하여 추후 구현 시 데이터 마이그레이션 없이 재개 가능. U+202E RLO 오버라이드 마커 및 관련 FlowDirection 분기 코드 전체 제거.

### Added
- **Added** — 페이지 속성 / 글상자 속성에 글자 방향 설정 추가. `TextOrientation`(가로/세로 enum) + `TextProgression`(오른쪽으로/왼쪽으로 진행 enum) 을 `PageSettings` 와 `TextBoxObject` 양쪽에 도입. **가로쓰기 LTR/RTL** 은 `FlowDocument.FlowDirection` / `RichTextBox.FlowDirection` 으로 즉시 시각 반영 (왼쪽으로 진행 = RTL, 아랍어/히브리어 또는 옛 한글 가로). **세로쓰기** 는 모델만 보존 — WPF FlowDocument 가 native CJK 세로조판을 지원하지 않아 다음 사이클에서 커스텀 레이아웃으로 도입 예정. 페이지 속성 다이얼로그(용지 탭 → "글자 방향" 두 ComboBox) 와 글상자 속성 다이얼로그 모두 동일한 입력. IWPF 직렬화 자동 포함.
- **Added** — 글상자에 회전 각도 (`RotationAngleDeg`, 기본 0, -360~360) 추가. 속성 대화상자의 "회전 / 각도(°)" 입력으로 박스를 시계방향으로 회전. 모양·테두리·본문 텍스트가 모두 박스 중심을 피벗으로 함께 회전. 번개상자를 15~30° 정도 기울이면 더 번개 같은 실루엣이 된다. (구현은 `RenderTransform` 이라 레이아웃 슬롯은 회전 전 사각형 — 회전 상태에서 핸들 리사이즈는 마우스 이동 방향과 변 방향이 어긋나 약간 어색하므로 회전 0으로 되돌려 조절을 권장.)
- **Added** — 글상자 4종 모양에 형태 파라미터 추가. `TextBoxObject` 모델에 `SpeechDirection`(`SpeechPointerDirection` enum: 8방향), `CloudPuffCount`(6~32, 기본 10), `SpikeCount`(5~24, 기본 12), `LightningBendCount`(1~5, 기본 2) 추가. `TextBoxOverlay` 의 정적 PathGeometry 상수를 동적 생성기(`BuildSpeechPath`/`BuildCloudPath`/`BuildSpikyPath`/`BuildLightningPath`)로 교체 — 100×100 정규화 좌표 위에서 파라미터에 따라 path 문자열을 매번 빌드. 말풍선 꼬리는 8방향(상하좌우+4모서리) 어디든, 구름은 베지어 호 N개, 가시별은 N각 별, 번개는 N개 꺽임 템플릿. 속성 대화상자(`TextBoxPropertiesWindow`)에 모양 선택값에 따라 표시되는 conditional 패널 4개 — 말풍선 꼬리 방향(3×3 RadioButton 그리드), 구름 뭉게 개수, 가시 개수, 번개 꺽임 개수.
- **Added** — 글상자 속성 대화상자 확장. 안쪽 여백 4방향(위·아래·왼·오른, mm) 입력란 추가. 가로 정렬(왼쪽/가운데/오른쪽/양쪽) + 세로 정렬(위/가운데/아래) ComboBox 추가. 테두리·배경색 hex 입력란 옆에 Color Picker 버튼(클릭 시 `System.Windows.Forms.ColorDialog` 전체보기 모드로 열림) 추가.
- **Added** — 글상자 우클릭 메뉴에 "글자 속성..." / "문단 속성..." 항목 추가. `CharFormatWindow` / `ParaFormatWindow` 를 글상자 내 `RichTextBox` (`InnerEditor`) 에 연결해 기존 글자·문단 서식 대화상자를 그대로 재사용.
- **Added** — 글상자 `TextBoxObject` 모델에 `PaddingTopMm`/`PaddingBottomMm`/`PaddingLeftMm`/`PaddingRightMm` (기존 단일 `PaddingMm` 대체), `HAlign`(`TextBoxHAlign` enum), `VAlign`(`TextBoxVAlign` enum) 추가. IWPF 직렬화 자동 포함.
- **Fixed** — 글상자 크기 조절 핸들 클릭 시 이동만 되고 리사이즈가 안 되던 버그. `PreviewMouseLeftButtonDown`(tunneling) 이벤트에서 UserControl 루트가 먼저 발화해 `e.Handled = true` 를 세팅, 핸들 Rectangle 의 `OnHandleMouseDown` 이 호출되지 않던 원인. `OnRootMouseDown` 에서 핸들 클릭(`Tag: "TL"/"TR"/"BL"/"BR"`)을 조기 감지해 즉시 리턴, 이벤트가 핸들까지 터널링되도록 수정.
- **Added** — 글상자 4종 추가 모양 렌더링 (Speech/Cloud/Spiky/Lightning). `TextBoxOverlay` 에 `Path Stretch=Fill` 방식으로 각 모양의 `PathGeometry` 문자열(100×100 정규화 좌표)을 정의 — 말풍선(하단 중앙 삼각 꼬리), 구름풍선(베지어 곡선 다중 돌기), 가시풍선(12각 별형), 번개상자(번개 볼트 실루엣). 테두리 색·두께·배경색을 모양별로 일관 적용.
- **Added** — 글상자 속성 대화상자 (`TextBoxPropertiesWindow`). 우클릭 컨텍스트 메뉴 → "속성..." 으로 열림. 테두리 색(hex), 테두리 두께(pt), 배경색(hex) 입력란 + 실시간 색상 미리보기. 확인 시 모델 갱신 + `AppearanceChangedCommitted` 이벤트 발행 → Dirty 플래그 갱신.
- **Added** — 글상자 우클릭 컨텍스트 메뉴. 속성/앞으로 가져오기/뒤로 보내기/삭제 항목. 앞/뒤 이동은 `FloatingCanvas` 자식 순서(ZOrder) 조정.
- **Added** — 글상자 내부 편집기를 `TextBox` → `RichTextBox` 로 교체. 볼드·이탤릭 등 WPF 기본 서식 단축키(Ctrl+B/I/U) 사용 가능. 단락별 plain-text 는 `FlowDocument.Blocks` 에서 `TextRange` 로 읽어 `TextBoxObject.Content` 에 동기화.
- **Added** — 입력 > 글상자 (사각형). 메뉴 클릭 후 종이 위에 마우스 드래그로 위치/크기를 직접 지정해 생성. 모델은 `Section.FloatingObjects` 에 `TextBoxObject`(IWPF 직렬화: `FloatingObjectJsonConverter`, `$type="textbox"`) 로 추가. 본문 흐름과 별도 레이어 (`FloatingCanvas`) 에 렌더링되어 본문 편집과 충돌하지 않음. 4코너 핸들로 리사이즈, 외곽 클릭으로 이동, 내부 클릭으로 본문 편집. 선택 시 점선(파랑) chrome 표시, Esc 로 선택 해제 또는 드래그 모드 취소, Delete 로 삭제.
- **Added** — 툴바 오른쪽에 확대/축소 조절기 추가. "−/+" 버튼(10% 단계), 배율 직접 입력 TextBox(Enter·LostFocus 로 적용), "폭 맞춤"(뷰포트 너비 기준)·"쪽 맞춤"(너비·높이 중 작은 쪽 기준) 버튼. 배율은 `ZoomPercent`(10–500 %) 에 저장되며 `PaperStackPanel.LayoutTransform(ScaleTransform)` 으로 편집창 전체에 적용. 쪽 맞춤은 빈 문서에서도 `PaperBorder.MinHeight`(한 페이지 분량) 기준으로 계산.
- **Added** — 메뉴 바로 아래에 메인 툴바 추가. 1단계로 파일 그룹(새 파일/불러오기/저장/다른 이름으로 저장/미리보기/인쇄/닫기) 7개 버튼 노출. 아이콘은 `Resources/ToolbarIcons.xaml` 의 `DrawingImage` 리소스를 `App.xaml` 에 머지해 사용하며, SVG 원본 80개는 `Resources/Icons/svg/` 에 임베드(추후 `Svg.Skia`/`SharpVectors` 로 동적 로딩 예정). 미리보기·인쇄는 메뉴와 동일하게 비활성.
- **Added** — `PageSettings.ShowMarginGuides` (기본 `true`) — 편집창 용지 위에 점선(파랑) 여백 안내선을 표시. 페이지 서식 > 여백 탭에 "여백 안내선 표시" 체크박스 추가.
- **Added** — 편집 영역 페이지 뷰. `ScrollViewer`(스크롤바 상시 표시) 위에 종이(`PaperBorder`)가 떠 있는 형태로 재구성. 그림자(`DropShadowEffect BlurRadius=18`) 적용, 캔버스 배경은 테마별 `EditorCanvasBg` 리소스. `PageSettings` 가 변경될 때마다 `PaperBorder.Width`(방향 보정 포함)와 배경색을 자동 갱신. 본문 들여쓰기는 `RichTextBox.Padding` 으로 페이지 여백을 반영 (RichTextBox 는 FlowDocument.PagePadding 을 무시).
- **Added** — 서식 > 페이지 모양 다이얼로그 (`PageFormatWindow`). 용지 크기(ISO A·B, JIS/KS B, 미국·국제, 신문 판형, 한국/국제 도서 판형 전체), 용지 색상, 방향(세로/가로), 여백(위·아래·좌·우·머리말·꼬리말), 다단(N단+단간격), 페이지 번호 시작값, 미리보기(여백·단 구분선 포함). 머리말·꼬리말 입력 UI 는 다음 사이클에서 추가. `PageSettings` 를 `FlowDocumentBuilder` 에서 읽어 `FlowDocument.PageWidth/PagePadding/ColumnWidth/ColumnGap` 및 용지 배경색에 반영.
- **Added** — 서식 > 개요 서식 다이얼로그 (`OutlineStyleWindow`). Body/H1~H6 7개 수준의 글자 모양·문단 모양·번호 매기기·테두리·배경색을 수준별로 편집. 글자/문단 편집은 기존 `CharFormatWindow` / `ParaFormatWindow` 의 단독 모드(에디터 없이 RunStyle / ParagraphStyle 직접 받는 생성자)를 재사용해 구현. 내장 4가지 프리셋(기본/학술/비즈니스/모던)을 제공하며 수준별 초기화 버튼 포함. `FlowDocumentBuilder` 가 `PolyDonkyument.OutlineStyles` 를 읽어 제목/본문 단락에 사용자 정의 스타일을 반영. `PolyDonkyument.OutlineStyles` (`OutlineStyleSet`) 는 IWPF 직렬화 시 자동 포함. `OutlineLevelStyle` — RunStyle + ParagraphStyle + OutlineNumbering + OutlineBorder + BackgroundColor(hex).

### Fixed
- **Fixed** — 왼쪽으로 진행(RTL) 모드에서 새로 입력한 글자가 기존 글자의 *왼쪽* 이 아니라 *오른쪽* 에 붙어 기존 글자가 좌측으로 밀려나던 버그 (배치만 우측 정렬되고 입력 방향은 LTR). WPF `FlowDirection.RightToLeft` 는 paragraph base direction 만 RTL 로 만들 뿐, 한글/라틴 같은 Bidi-약방향 문자는 Unicode Bidi 알고리즘에 의해 LTR run 으로 묶여 시각 순서가 좌→우로 유지된다. paragraph 시작에 **U+202E (RLO, Right-to-Left Override)** 마커를 자동 삽입해 문단 내 모든 글자를 강제로 RTL 시각 흐름으로 override 하도록 수정 — 이제 새 글자가 시각적으로 기존 글자의 왼쪽에 붙으며, 기존 글자가 우측으로 밀리지 않는다. `FlowDocumentBuilder.Build` 가 RTL FlowDocument 의 모든 paragraph (List/Table/Section 안 포함) 시작에 RLO 를 박고, `MainWindow.OnEditorTextChanged` / `TextBoxOverlay.OnInnerTextChanged` 가 새 paragraph (Enter 등) 가 생길 때마다 RLO 자동 보충. `TextBoxOverlay.SyncEditorToModel` 은 모델로 동기화할 때 RLO 마커 제거 (디스플레이 전용). LTR 복귀 시 `RemoveRloFromAllParagraphs` 로 마커 제거.
- **Fixed** — 왼쪽으로 진행(RTL) 모드에서 글자 입력 시 초기 문자들이 보이지 않고 4자 이상 입력해야 나타나는 버그. `RichTextBox.FlowDirection = RightToLeft` 만으로는 컨테이너 스크롤 원점이 좌측이라 텍스트 시작점(우측)이 뷰포트 밖에 위치해 초기 글자가 가려졌다. 레이아웃 패스 직후 `Dispatcher.BeginInvoke` 로 `ScrollToRightEnd()` 를 호출해 원점을 우측으로 이동시켜 수정. (FlowDocument 에만 `FlowDirection` 을 걸면 Latin/Hangul 같은 Bidi-약방향 문자가 우측 정렬된 LTR run 으로만 표시되어 실제 입력 방향이 LTR 이 되므로, 컨테이너 자체를 RTL 로 두는 것이 필수.) `TextBoxOverlay` 글상자와 `BodyEditor` 메인 편집기 모두 동일하게 적용.
- **Fixed** — 왼쪽으로 진행(RTL) 모드에서 좌우 화살표 키가 기대 방향과 반대로 동작하는 버그(Left→시각 우, Right→시각 좌). WPF RTL 에서 Left 키 = 논리 이전(시각 오른쪽), Right 키 = 논리 다음(시각 왼쪽) 이므로 `PreviewKeyDown` 에서 Left↔Right 커맨드를 교환, 시각 방향과 일치하도록 수정. Shift(문자 선택)·Ctrl(단어 이동/선택) 조합도 함께 처리.
- **Fixed** — 개요 서식 다이얼로그: 선색·배경색 스와치 클릭 시 `System.Windows.Forms.ColorDialog` 색상 선택기가 열리도록 수정 (기존은 텍스트박스 포커스만 이동). 배경색 그룹에 "없음 (투명)" 체크박스 추가 — 체크 시 배경색 입력 비활성화 및 `BackgroundColor = null` 적용.
- **Fixed** — 글자폭·자간이 적용된 영역이 통째로 atomic 요소로 잡혀 클릭만 해도 전체 범위가 다시 선택되고 캐럿이 안으로 들어가지 못하던 문제. `FlowDocumentBuilder.BuildScaledContainer` 가 더 이상 한 Run 전체를 하나의 `InlineUIContainer` (TextBlock/StackPanel) 로 감싸지 않고, **Span 안에 문자별 InlineUIContainer 를 나열**하도록 변경. 각 IUC 와 부모 Span 모두 같은 `PolyDonky.Run` 을 Tag 로 가져 라운드트립 머지 단서가 된다. `FlowDocumentParser` 의 Span 케이스에 `TryMergePerCharSpan` 추가 — 모든 자식이 같은 Tag 의 per-char IUC 면 한 Run 으로 머지하고, 사용자 편집(중간에 Run 삽입 등)이 있으면 자식별 fallback. `CharFormatWindow.CollectLeafInlines` / `GetFirstInlineInSelection` / `ApplyTypographicProps` 도 Span 단위 교체를 인식하도록 갱신.
- **Fixed** — 글자 서식 적용 후 InlineUIContainer 가 atomic 요소로 잡혀 선택 영역이 시각상 묶여 보이고 자동 해제되지 않던 문제. `CharFormatWindow.OnOk` 에서 적용 직후 `_editor.Selection.Select(end, end)` 으로 캐럿을 끝으로 collapse.
- **Fixed** — IWPF 옛 / 누락 discriminator 호환. `BlockJsonConverter` 신설 — 읽기 시 `$type` (현행) 외에 `kind` (29c09bd 시기 옛 빌드) 도 인식하고, 둘 다 없으면 `Paragraph` 로 폴백. 사용자가 옛 빌드로 저장한 iwpf 를 신 빌드에서 그대로 열 수 있다. 쓰기는 항상 `$type` 로 출력. xUnit 회귀 테스트 2건 추가 (legacy `kind` / 누락 discriminator).
- **Fixed** — 테마 변경이 메인 윈도우에 적용되지 않던 문제. `MainWindow` / `AboutWindow` / `FindReplaceWindow` / `SettingsWindow` 의 모든 테마 리소스 참조를 `StaticResource` → `DynamicResource` 로 교체. (StaticResource 는 로드 시 한 번만 해석되어 사전 교체 후에도 옛 값을 유지.)
- **Fixed** — 설정 창 우측 잘림 — 너비 320 → 380.
- **Fixed** — 찾기/바꾸기에 "바꾸기(_E)" 단일 교체 버튼 추가, "모두 바꾸기(_A)" 유지. 상태 메시지 행을 별도 `Auto` 행으로 분리해 버튼 위치 흔들림 제거.
- **Fixed** — 설정 창 조용한 종료. `OnThemeChecked` 가 `InitializeComponent` 중 `ThemeDark`/`ThemeSoft` 가 아직 null 인 상태에서 호출돼 `NullReferenceException`. null guard 추가.
- **Fixed** — B 폴리싱 4종 빌드 오류 수정 (2226861 에서 미적용). `Properties/Resources.Designer.cs` 누락으로 `dotnet build` 가 실패해 새 기능이 배포되지 않던 문제 해결 — Designer.cs 수동 생성 (`ResourceManager` + 정적 속성). 드래그&드롭: `RichTextBox` 가 `DragOver`를 가로채 파일 드롭 이벤트가 윈도우에 도달하지 않던 문제 — Window 이벤트를 `Drop`/`DragOver` → `PreviewDrop`/`PreviewDragOver` 로 변경, 파일 드롭만 처리 후 `Handled=true`.

### Added
- **Added** — 서식 > 글자 서식 다이얼로그 (`CharFormatWindow`). RichTextBox 선택 영역의 글꼴·크기·굵게·기울임꼴·밑줄·취소선·위선·위첨자·아래첨자·글자색·배경색을 읽어 표시하고, OK 시 선택 영역에 일괄 적용. 선택이 없으면 캐럿 위치 서식을 읽어 이후 입력에 반영. 미리보기 TextBlock 실시간 갱신. 선택 혼합값(mixed)은 세 상태 체크박스(불확정)로 표시 — 값을 바꾸지 않으면 해당 속성 변경 없음.
- **Added** — 글자폭(WidthPercent)·자간(LetterSpacingPx) WPF 시각화. 글자폭 != 100% 인 Run 은 `ScaleTransform` + `InlineUIContainer(TextBlock)`으로 렌더링. 자간은 `Typography.CharacterSpacing`(1/1000 em)으로 적용. `CharFormatWindow` 글꼴 그룹에 글자폭(%) / 자간(px) 입력란 추가 및 미리보기 실시간 반영. `FlowDocumentParser` 가 `InlineUIContainer`·`CharacterSpacing` 값을 파싱해 IWPF 라운드트립 보존.
- **Added** — 서식 > 문단 서식 다이얼로그 (`ParaFormatWindow`). 선택 영역에 걸친 단락의 정렬(왼/가운데/오른/양쪽/배분), 줄 간격(%), 단락 위/아래 여백(pt), 첫 줄·왼쪽·오른쪽 들여쓰기(mm), 개요 수준(본문/H1~H6) 을 일괄 적용. `Paragraph.Tag`(PolyDonky.Paragraph)의 `Style` 도 함께 갱신해 비-FlowDocument 속성(IndentMm, SpaceBeforePt 등)까지 라운드트립 보존. `FlowDocumentScrollViewer` 기반 실시간 미리보기.

### Fixed
- **Fixed** — 다른 이름으로 저장 시 형식을 바꿔도 원본 파일명(확장자 포함)이 그대로 남아 "이미 있습니다" 경고가 뜨던 문제. `SaveFileDialog.FileName` 초기값을 `GetFileName` → `GetFileNameWithoutExtension` 으로 변경 — 선택한 필터의 확장자가 자동 적용된다.

### Added
- **Added** — 암호 설정 다이얼로그를 체크박스 기반으로 재설계. 열기·쓰기 암호를 독립적으로 설정 가능 — 같이사용 체크 시 단일 입력란, 미체크 시 열기/쓰기 각각 다른 입력란 활성화. `IwpfWriter.WritePassword` 추가로 Both 모드에서 열기와 쓰기 암호를 별도로 지정 지원.
- **Added** — 쓰기 보호 자동 잠금 해제 UX. 쓰기 보호된 IWPF 를 열면 `RichTextBox.IsReadOnly=true` 로 시작 + 상태 표시줄에 "쓰기 보호됨" 인디케이터(주황). 사용자가 첫 편집(타이핑·Backspace·Delete·Enter·Tab·Ctrl+V/X)을 시도하면 즉시 비밀번호 프롬프트. 정답이면 같은 세션 내 추가 입력·저장 시 재입력 요구하지 않음.
- **Added** — IWPF 암호 보호 3단계 모드. 단순 암호화(열기 보호)에서 열기 / 쓰기 / 둘 다 보호 모드로 확장. 쓰기 보호(`PasswordMode.Write`)는 AES 암호화 없이 PBKDF2 해시만 `security/write-lock.json` 에 저장해 저장 시 비밀번호를 검증한다; 둘 다(`Both`)는 AES-256-GCM 암호화 + inner ZIP 내 write-lock 병행. `PasswordChangeWindow` 에 보호 모드 RadioButton 추가. 문서 정보 보안 탭에서 현재 모드를 표시.
- **Added** — 편집 > 문서 정보 다이얼로그 확장: 3개 탭 (정보 / 보안 / 워터마크). 작성자(Author) 입력 가능. 저장 시 첫 저장이면 작성일자 + 수정일자, 이후 저장에서는 수정일자만 자동 갱신.
- **Added** — IWPF 문서 암호 보호. 다음 저장부터 PBKDF2-HMAC-SHA256 (200,000회) 으로 키 유도 + AES-256-GCM 으로 inner ZIP 전체 암호화 → outer envelope ZIP 봉인 (`security/envelope.json` + `security/payload.bin`). 옵션 패키지라 기존 평문 IWPF 와 100% 호환. 암호 보호 IWPF 를 열 때는 `PasswordPromptWindow` 로 입력 + 틀리면 GCM tag 불일치로 즉시 감지·재시도.
- **Added** — IWPF 워터마크 설정 (텍스트 / 색상 / 글자 크기 / 회전 / 불투명도). `PolyDonkyument.Watermark` 로 직렬화 — 옵션 필드라 기존 IWPF 와 호환. 편집기 화면 미리보기·인쇄 렌더링은 후속 단계.
- **Security** — AES-256-GCM 인증 암호화로 본문/스타일/리소스를 모두 봉인 — 잘못된 비밀번호 또는 변조 시 GCM tag 검증으로 거부. 암호는 메모리 내에서만 보관되고 비밀번호 자체는 저장하지 않는다 (envelope 에는 salt/nonce/tag 만 보관).
- **Added** — B 폴리싱 4종.
  - **드래그&드롭**: `MainWindow` 에 `AllowDrop=True` + `PreviewDrop`/`PreviewDragOver` 핸들러. 파일을 창에 끌어놓으면 `OpenFile(path)` 호출로 즉시 열기. `MainViewModel.OpenFile` 공개 메서드 추가, 내부 `OpenPath` 로 중복 제거.
  - **찾기/바꾸기**: `FindReplaceWindow` 모달리스 다이얼로그 (Ctrl+F / Ctrl+H). `FlowDocumentSearch` 헬퍼 — WPF `TextPointer` 기반 순방향 검색(`FindNext`) + `ReplaceAll`. 대소문자 구분 옵션, wrap-around, 상태 메시지 표시.
  - **i18n .resx 한/영 분리**: `Properties/Resources.resx` (한국어 기본) + `Properties/Resources.en-US.resx` (영어 병행). 메뉴 항목 / 다이얼로그 제목·버튼 / 상태 메시지 / 에러 메시지 전부 리소스 키로 분리. `MainViewModel` 의 모든 하드코딩 문자열을 `SR.Xxx` 참조로 교체.
  - **테마 다중화**: `Themes/Dark.xaml` (어두운 배경·밝은 텍스트) + `Themes/Soft.xaml` (따뜻한 아이보리 색조, 학생·장년 대상). `ThemeService` — `Application.Resources.MergedDictionaries[0]` 교체로 런타임 즉시 전환. `SettingsWindow` (도구→설정) — 라디오 버튼으로 Light/Dark/Soft 실시간 미리보기.
- **Added** — HWPX OpaqueBlock 양방향. Reader: `<hp:rect>` / `<hp:line>` / `<hp:ellipse>` / `<hp:arc>` / `<hp:polygon>` / `<hp:textBox>` / `<hp:connector>` 를 `OpaqueBlock(Format="hwpx", Kind, Xml=원본XML)` 으로 보존 (pic/tbl 처리 방식 동일). Writer: `OpaqueBlock.Format=="hwpx"` 이면 Xml 을 파싱해 `<hp:p><hp:run>` 안에 원형 재출력; 다른 포맷(docx 등)은 placeholder 단락으로 fallback. xUnit 테스트 2건 추가 (hwpx OpaqueBlock 보존 / 비hwpx fallback). 13 → 15건.
- **Fixed** — HWPX 자체 라운드트립 시 글자 크기·색상·폰트 패밀리 손실 수정. `HwpxWriter` 를 동적 charPr/paraPr/font registry 방식으로 재설계 — 문서에서 사용된 모든 `RunStyle`(폰트 크기·색·패밀리·굵게/기울임 등)을 고유 charPr 항목으로 header.xml 에 기록하고, `ParagraphStyle`(정렬·줄간격·여백)도 고유 paraPr 항목으로 기록. 기존 6개 고정 charPr ID 방식을 대체. xUnit 라운드트립 4건 추가 (FontSize / ForegroundColor / FontFamily / 복합 서식). 9 → 13건.
- **Added** — HWPX 표 / 이미지 양방향. Reader: `<hp:tbl>` → Table (rows·cells·cellSpan·cellSz, 중첩 표), `<hp:pic>` 의 `<hp:img binaryItemIDRef>` → ImageBlock (BinData/{stem}.* 파일 매칭, 바이트 추출, curSz 의 hwpunit → mm). 표 안 paragraph 는 본문 평탄화에서 제외하고 셀 본문에 모음. Writer: Table → `<hp:tbl>` (sz/outMargin/inMargin 최소 valid 정의 + tr/tc/cellAddr/cellSpan/cellSz/cellMargin), ImageBlock → `<hp:pic><hp:img>` 와 BinData/{imageN}.{ext} 추가. SHA-256 dedupe 로 같은 이미지는 한 번만 저장. xUnit 라운드트립 6 → 9건 (Table 구조, 이미지 바이트 동일성, BinData dedupe 추가). 자체 라운드트립 + 스모크 전부 그린 유지.
- **Added** — HWPX 한컴 서식 회수 — `HwpxHeader` / `HwpxHeaderReader` 신설. header.xml 의 `fontfaces` / `charPr` / `paraPr` / `style` 정의를 PolyDonky 모델로 매핑해 한컴이 만든 hwpx 의 임의 ID(0~5 약속과 다른) 도 본문 서식이 살아남는다.
  - charPr: height(0.01pt 단위) / textColor / shadeColor / `<bold>` / `<italic>` / `<underline>` / `<strikeout>` / `<sup·subscript>` / `<fontRef hangul/latin/hanja...>` → `RunStyle`.
  - paraPr: `<align horizontal>` / `<margin left/right/intent/prev/next>` (hwpunit → mm) / `<lineSpacing PERCENT>` → `ParagraphStyle`.
  - style: name/engName 이 "Heading{N}" 또는 한국어 "개요{N}" 일 때 `OutlineLevel.H1~H6` 추정. paraPrIDRef + charPrIDRef 보존.
  - HwpxReader: header context 를 ReadSectionFromDoc → ReadParagraph → ReadRun 로 전달. 우선순위는 「style 의 paraPr/charPr → paragraph 직접 paraPrIDRef → run 직접 charPrIDRef」 로 override. header 가 매핑을 못 가진 ID 는 우리 자체 codec 의 0~5 약속을 fallback 으로 사용 (자체 라운드트립 호환 유지). 자체 라운드트립 6/6 + 스모크 6/6 변동 없이 그린.
- **Added** — `PolyDonky.Core/DocumentMeasurement` — PolyDonkyument 가 차지하는 데이터 크기(텍스트 byte + ImageBlock.Data + OpaqueBlock 바이트/XML, 표는 셀 재귀)를 근사 계산하는 헬퍼. 단위 자동 (B / KB / MB / GB).
- **Added** — `MainWindow` 상태 표시줄 우측에 5칸 그룹: **파일 경로 · 문서 메모리 · 삽입/수정 · CapsLock · NumLock**. ItemsPanel 을 `DockPanel(LastChildFill=True)` 로 교체해 좌측 상태 메시지가 가변 너비를 차지하고 우측 5칸은 한 묶음으로 우측에 고정. 상태 메시지 길이가 변해도 우측 칸 위치가 흔들리지 않는다. 메모리 표시는 **앱 전체 워킹셋이 아닌 문서 콘텐츠 크기** 만 보여주도록 변경 (`DocumentMeasurement.EstimateBytes` 사용) — 새 만들기 직후엔 자연스럽게 0 가까이 떨어지고, HWPX 처럼 본문 인식이 0건이면 작은 값으로 표시되어 진단 신호가 됨.
- **Added** — HWPX reader 진단 정보 — 읽은 section 파일 수 / 인식된 paragraph 수 / 비어있지 않은 run 수 / 첫 section 경로를 `DocumentMetadata.Custom["hwpx.*"]` 에 박는다. 본문 인식이 0건이면 MainViewModel 의 상태 메시지에 "HWPX 본문 인식 0건, 섹션 파일 N개. 한컴 변종 가능 — 진단 정보를 메인테이너에게 공유 부탁" 안내. FallbackSectionPaths 검색 범위를 `Contents/section*.xml` → 폴더 무관 + 파일명 contains "section" + 대소문자 무시로 확장.
- **Changed** — `HwpxReader` 를 한컴 오피스가 만든 hwpx 변종에 robust 하게 매칭하도록 보강. `content.hpf` 와 `container.xml` 모두 LocalName 기반(namespace 무시) descendants 검색. section 파일은 manifest의 id 또는 href 파일명이 "section" 접두인 것으로 fallback. 그래도 못 찾으면 ZIP 의 `Contents/section*.xml` 직접 스캔. paragraph/run/text 모두 깊이 어디든 `Descendants` 로 흡수, `<hp:tab>` / `<hp:lineBreak>` 추가 인식. 자체 라운드트립 6/6 + 스모크 6/6 그대로 그린.
- **Added** — Phase C C3·C4 HWPX 1급 시민 codec 1차 — `src/PolyDonky.Codecs.Hwpx`. KS X 6101 사양 기반 자체 구현 (BCL + System.Xml.Linq + System.IO.Compression, 외부 의존 0).
  - 패키지 구조: `mimetype`(STORED, "application/hwp+zip") + `META-INF/container.xml` + `Contents/content.hpf` (OPF) + `Contents/header.xml` + `Contents/section{N}.xml` + `version.xml`.
  - HwpxWriter: 단락(`hp:p`), 런(`hp:run`+`hp:t`), 정렬(LEFT/CENTER/RIGHT/JUSTIFY → paraPr 0~3), 굵게/기울임/밑줄/취소선 (charPr 0~5), 헤더 H1~H6 (style 1~6). header.xml 의 charPr/paraPr/style 정의를 codec 내부 ID 약속으로 고정해 라운드트립 호환성 보장.
  - HwpxReader: container.xml → content.hpf → spine 으로 section{N}.xml 들 순회. 각 섹션의 `hp:p` 와 `hp:run` 을 PolyDonky 모델로 복원. 잘못된 mimetype 거부.
  - `tests/PolyDonky.Codecs.Hwpx.Tests` xUnit 라운드트립 6건 (단락·헤더·정렬·강조·mimetype STORED·필수 파트 존재).
  - `PolyDonky.SmokeTest` 에 HWPX 라운드트립 추가 → 6/6 그린.
  - `PolyDonky.App` 의 `KnownFormats`: `.hwpx` 가 외부 컨버터 위탁 목록에서 제거되고 직접 처리. OpenFilter/SaveFilter 의 「PolyDonky 직접 지원」 그룹에 .hwpx 포함, 「외부 컨버터 필요」 그룹에서 제외. 이제 외부 위탁은 HWP/DOC/HTML 만 남는다.
  - 한컴 오피스 호환은 G3 검증 후 다음 사이클에서 fine-tune (현재는 PolyDonky 자체 라운드트립만 보장).
- **Added** — 비텍스트 객체(표·이미지·미인식 도형) 1차 모델링 + 라운드트립. IWPF.md 의 「opaque island」 정책 본격 적용.
  - `PolyDonky.Core` 블록 추가: `Table` / `TableRow` / `TableCell` / `TableColumn`, `ImageBlock`, `OpaqueBlock` (Block 의 `JsonDerivedType` 디스크리미네이터 4종 등록).
  - `PolyDonky.Iwpf`: ImageBlock 의 binary 를 `resources/images/img-NNNN.<ext>` 로 분리 저장하고 SHA-256 dedupe. 매니페스트에 별도 part 로 기록되어 무결성 검증. 다형성 디스크리미네이터를 `kind` → `$type` 으로 변경 (`OpaqueBlock.Kind` 속성과의 충돌 회피).
  - `PolyDonky.Codecs.Docx`: Reader 가 `w:tbl` → Table (셀 너비·병합·중첩 표), `w:drawing` 의 그림 → ImageBlock (ImagePart 바이너리 추출, EMU → mm 사이즈 보존, alt text), 미인식 블록 → OpaqueBlock(Format="docx", Xml=OuterXml). Writer 는 대칭으로 `w:tbl` / `w:drawing`+`ImagePart` 등록 / OpaqueBlock OuterXml 을 임시 Body 로 파싱 후 자식만 옮겨 그대로 재출력.
  - `PolyDonky.App` (WPF): FlowDocumentBuilder 가 Table → `Wpf.Table`, ImageBlock → `BlockUIContainer + Image` (메모리 BitmapImage), OpaqueBlock → 회색 placeholder Paragraph 로 시각화. FlowDocumentParser 가 Tag 머지로 비파괴 회수 (사용자가 셀 텍스트만 편집해도 표 구조·컬럼 너비·이미지 바이너리 보존).
  - 테스트: DOCX 라운드트립 9건(표·이미지 바이트 동일성·OpaqueBlock 보존 추가), IWPF 라운드트립 9건(Table 구조·이미지 resources/images 분리·dedupe·OpaqueBlock 추가). xUnit 합계 36 → 43건. 스모크 5/5.
- **Changed** — 본문 편집기를 `TextBox` (plain string) 에서 **`RichTextBox` + `FlowDocument`** 로 업그레이드. DocxReader/IwpfReader 가 이미 가져온 폰트·크기·색·굵게·기울임·밑줄·취소선·위·아래첨자·정렬·줄간격·문단간격·들여쓰기·헤더 레벨이 화면에 그대로 표시되고, 사용자가 그대로 편집·저장 가능.
- **Added** — `src/PolyDonky.App/Services/FlowDocumentBuilder.cs` — PolyDonkyument → WPF FlowDocument 매퍼. RunStyle (폰트·크기 (pt→DIP)·색상 (PolyDonky.Color → SolidColorBrush)·강조·장식·BaselineAlignment), ParagraphStyle (정렬·간격·들여쓰기·LineHeight·Outline 헤더 시각화), ListMarker (Wpf.List/ListItem) 매핑. 원본 `Paragraph`/`Run` 을 Tag 에 보관해 Parser 가 비-FlowDocument 속성을 비파괴 보존.
- **Added** — `src/PolyDonky.App/Services/FlowDocumentParser.cs` — FlowDocument → PolyDonkyument 역매퍼. Tag 머지로 한글 조판(장평·자간) / Provenance / 페이지 설정 비파괴 보존. FontWeight/FontStyle/TextDecorations/Foreground/Background/BaselineAlignment 추출, FontSize 로 헤더 레벨 추정.
- **Added** — `tests/PolyDonky.App.Tests` (net10.0-windows + UseWPF) — FlowDocumentBuilder/Parser 라운드트립 9건. 본 환경(Linux)에선 WPF 의존이라 build/test 못 돌리고 사용자 Windows 검증 (G2.5).
- **Changed** — `MainViewModel`: `DocumentBody` 문자열 제거, `FlowDocument` ObservableProperty 노출. `LoadDocument` 헬퍼로 Open/New 통일. SaveTo 가 FlowDocument 를 Parser 에 보내 원본 `_document` 를 머지 베이스로 회수해 비파괴 저장. `MarkDirty()` public 메서드.
- **Changed** — `MainWindow`: 본문 영역을 `RichTextBox` 로 교체. code-behind 가 ViewModel `FlowDocument` 변경 시 `BodyEditor.Document` 동기화, `TextChanged` 에서 `vm.MarkDirty()` 호출. 프로그램적 변경 중에는 `_suppressTextChanged` 플래그로 dirty 회피.

- **Added** — Phase C (1/N) DOCX 1급 시민 codec — `src/PolyDonky.Codecs.Docx`. DocumentFormat.OpenXml 3.5.1 기반 reader/writer. 단락 / Heading1~6 / 정렬(좌·중·우·양쪽·균등) / 굵게·기울임·밑줄·취소선·위첨자·아래첨자 / 폰트 패밀리·크기 / 색상 / 기본 리스트 / Title·Author 코어 속성 라운드트립. xUnit 라운드트립 6건 + 스모크 1건 그린.
- **Added** — `tests/PolyDonky.Codecs.Docx.Tests` xUnit 라운드트립 테스트 6건.
- **Added** — `Directory.Packages.props` 에 DocumentFormat.OpenXml 3.5.1 (MIT) 등록.
- **Changed** — `PolyDonky.Codecs.Markdown` 의 reader 를 BCL 직접 파서에서 **Markdig 0.42.0** (BSD-2-Clause) 으로 교체. 풀 CommonMark 파싱 — 단락 / 헤더 / 리스트 / 강조에 더해 인라인 코드(monospace 힌트), 코드블록(fenced/indented), 인용(현재는 단락으로 격하), 링크(밑줄 표시) 처리. 라운드트립 호환성 유지. xUnit 11건(기존 6 + 신규 5) 그린.
- **Changed** — `PolyDonky.App` 의 `DocumentFormat` 서비스가 DOCX 를 **외부 컨버터 위탁 목록에서 제거**하고 직접 처리 대상으로 등록. 이제 메인 앱에서 `.docx` 를 native 로 읽고 쓴다. 외부 컨버터 위탁은 HWP / HWPX / DOC / HTML / HTM 만 남는다.
- **Added** — 핸텍 공식 회사 로고/아이콘 자산: `assets/Handtech_1024.png` (1024×1024 PNG), `assets/Handtech.ico` (멀티 사이즈 Windows ICO).
- **Added** — `src/PolyDonky.App` WPF 앱 첫 사이클 (Phase B1~B4 골격). `net10.0-windows` + WPF + CommunityToolkit.Mvvm 8.4.0. 메인 윈도우(파일/편집/입력/서식/도구/도움말 메뉴), TextBox 본문 편집기, 상태 바, About 다이얼로그(로고·핸텍·노진문·버전 표시). 파일 메뉴는 IWPF/MD/TXT 직접 처리, 외부 포맷은 Phase D 안내 메시지. Ctrl+N/O/S, Ctrl+Shift+S 단축키. 한국어 UI(.resx 분리는 다음 사이클로 이연).
- **Added** — Light 기본 테마 (`src/PolyDonky.App/Themes/Light.xaml`) — 핸텍 브랜드 블루 기반.
- **Added** — `Directory.Packages.props` 에 CommunityToolkit.Mvvm 8.4.0 등록.
- **Internal** — Phase A 솔루션 골격: `PolyDonky.sln` + `src/PolyDonky.Core` + `src/PolyDonky.Iwpf` + `src/PolyDonky.Codecs.Text` + `src/PolyDonky.Codecs.Markdown` + 대응 `tests/*` xUnit 프로젝트 + `tools/PolyDonky.SmokeTest` 콘솔 러너. .NET 10 + Central Package Management.
- **Added** — `PolyDonky.Core` 공통 문서 모델 1차: `PolyDonkyument`, `DocumentMetadata`, `Section`/`PageSettings`, `Block`/`NodeStatus`, `Paragraph`/`ParagraphStyle`/`Alignment`/`OutlineLevel`/`ListMarker`/`ListKind`, `Run`/`RunStyle`/`Color`, `StyleSheet`, `Provenance`/`SourceAnchor`, `IDocumentReader`/`IDocumentWriter`/`IDocumentCodec`. 한글 조판용 `WidthPercent`(장평) / `LetterSpacingPx`(자간) 포함.
- **Added** — `PolyDonky.Iwpf` 1차 codec (writer/reader). ZIP+JSON 패키지(`manifest.json`, `content/document.json`, `content/styles.json`, 선택적 `provenance/source-map.json`). 매니페스트 SHA-256 해시 검증, packageType 검사, 위변조 거부.
- **Added** — `PolyDonky.Codecs.Text` (TXT in/out, BOM 자동 감지).
- **Added** — `PolyDonky.Codecs.Markdown` 1차 codec — Markdig 의존 없이 BCL 만으로 ATX 헤더(#~######), 순서/비순서 리스트, `**굵게**`, `*기울임*` 인라인을 처리하는 실용 서브셋. 추후 Markdig 도입 시 교체 예정.
- **Added** — `PolyDonky.SmokeTest` BCL-only 콘솔 스모크 러너 (`tools/PolyDonky.SmokeTest`). 라운드트립 + 위변조 검출까지 4건 통과.
- **Added** — xUnit 테스트 1차 (Core·Iwpf·Text·Markdown). 본 환경에서 NuGet 차단으로 미실행, Windows 에서 `dotnet test` 시 검증.
- **Internal** — `Directory.Build.props` (TreatWarningsAsErrors, LangVersion=latest, Company=HANDTECH, Authors=Noh JinMoon, Version=1.0.0-test.1), `Directory.Packages.props` (xUnit 2.9.3 / Test.Sdk 17.14.1 / Markdig 0.42.0), `global.json` (SDK 10.0.107 핀), .NET 표준 `.gitignore`.

### Changed
- **Docs** — 버전 정책 명시: 사용자 명시 지시 전까지 모든 빌드는 테스트 버전(`1.0.0-test.<n>`)으로 관리하고, 최초 정식 릴리스는 `1.0.0` 으로 한다는 규칙을 `HISTORY.md` / `CLAUDE.md` / `README.md` 에 일관되게 반영.
- **Docs** — `README.md` 헤더 로고를 GitHub 아바타에서 정식 핸텍 로고(`assets/Handtech_1024.png`)로 교체.
- **Docs** — `CLAUDE.md` 의 로고/아이콘 경로를 정식 자산(`assets/Handtech_1024.png`, `assets/Handtech.ico`) 으로 갱신.
- **Docs** — `README.md` 헤더에 핸텍 회사 로고 + 핸텍/노진문 표기, "만든 사람들" 섹션, .NET 10 배지 추가, UI 를 WPF 로 확정.

### Added (docs)
- **Docs** — `WORK_PLAN.md` 신설. 다단계 작업 계획서, 환경 사실관계, 기술 스택, Phase A~H 진행표, 사용자 게이트 G0~G5, 다음 세션 인수인계 체크리스트.
- **Docs** — `NOTICE` 신설. Apache 2.0 저작권 고지(© 2026 HANDTECH — Noh JinMoon) + 향후 의존성 attribution 사전 기록.

### Resolved Limitations
- NuGet.org 차단(503)이 풀려 본 개발 환경에서도 xUnit / Markdig / OpenXml SDK 복원이 정상 동작. Phase A 의 xUnit 25건이 첫 실행에서 그린, Phase C 진입과 Markdig 교체가 가능해짐.

### Fixed
- **Internal** — 솔루션 파일을 `.slnx` (.NET 9~ XML 신형) 에서 `.sln` (전통 형식) 으로 교체. 사용자 Windows Visual Studio 빌드에서 `PolyDonky.Codecs.Docx` 프로젝트가 `.slnx` 의존성 그래프에서 누락되어 빌드 큐에 들어가지 않는 회귀가 발생 (다른 12개 프로젝트는 정상). `.sln` 형식은 모든 도구(VS, MSBuild, dotnet CLI, Rider)에서 안정적으로 인식되므로 호환성 우선.

---

## [Pre-release]

정식 버전 부여 이전, 사양 정립 및 초기 문서화 단계의 기록입니다.

### 2026-04-25
- **Docs** — `HISTORY.md` 신설. 변경 이력을 별도 파일로 분리해 Keep a Changelog 형식으로 관리하기 시작.
- **Docs** — `README.md` 를 GitHub 방문자(사용자·기여자) 안내 중심으로 재작성. 기술 사양은 `IWPF.md` / `CLAUDE.md` 로 링크.
- **Docs** — `CLAUDE.md` 신설. Claude Code 세션이 참고할 개발 가이드라인 정리 (정본 IWPF, 2계층 설계, 외부 컨버터 분리, 한글 조판 특화, 단계별 구현 전략, 회피 사항).
- **Docs** — `IWPF.md` 작성. 자체 통합 포맷의 설계 근거·패키지 구조·보존 캡슐·provenance 정책 정의.
- **Docs** — `README.md` 초안 작성. 제품 개요, 메뉴 구성, 개발 원칙 정리.
- **Internal** — 저장소 초기화, Apache License 2.0 채택.

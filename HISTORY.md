# 변경 이력 (Change History)

PolyDoc의 모든 의미 있는 변경 사항을 이 파일에 기록합니다.

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

### Added
- **Changed** — 본문 편집기를 `TextBox` (plain string) 에서 **`RichTextBox` + `FlowDocument`** 로 업그레이드. DocxReader/IwpfReader 가 이미 가져온 폰트·크기·색·굵게·기울임·밑줄·취소선·위·아래첨자·정렬·줄간격·문단간격·들여쓰기·헤더 레벨이 화면에 그대로 표시되고, 사용자가 그대로 편집·저장 가능.
- **Added** — `src/PolyDoc.App/Services/FlowDocumentBuilder.cs` — PolyDocument → WPF FlowDocument 매퍼. RunStyle (폰트·크기 (pt→DIP)·색상 (PolyDoc.Color → SolidColorBrush)·강조·장식·BaselineAlignment), ParagraphStyle (정렬·간격·들여쓰기·LineHeight·Outline 헤더 시각화), ListMarker (Wpf.List/ListItem) 매핑. 원본 `Paragraph`/`Run` 을 Tag 에 보관해 Parser 가 비-FlowDocument 속성을 비파괴 보존.
- **Added** — `src/PolyDoc.App/Services/FlowDocumentParser.cs` — FlowDocument → PolyDocument 역매퍼. Tag 머지로 한글 조판(장평·자간) / Provenance / 페이지 설정 비파괴 보존. FontWeight/FontStyle/TextDecorations/Foreground/Background/BaselineAlignment 추출, FontSize 로 헤더 레벨 추정.
- **Added** — `tests/PolyDoc.App.Tests` (net10.0-windows + UseWPF) — FlowDocumentBuilder/Parser 라운드트립 9건. 본 환경(Linux)에선 WPF 의존이라 build/test 못 돌리고 사용자 Windows 검증 (G2.5).
- **Changed** — `MainViewModel`: `DocumentBody` 문자열 제거, `FlowDocument` ObservableProperty 노출. `LoadDocument` 헬퍼로 Open/New 통일. SaveTo 가 FlowDocument 를 Parser 에 보내 원본 `_document` 를 머지 베이스로 회수해 비파괴 저장. `MarkDirty()` public 메서드.
- **Changed** — `MainWindow`: 본문 영역을 `RichTextBox` 로 교체. code-behind 가 ViewModel `FlowDocument` 변경 시 `BodyEditor.Document` 동기화, `TextChanged` 에서 `vm.MarkDirty()` 호출. 프로그램적 변경 중에는 `_suppressTextChanged` 플래그로 dirty 회피.

- **Added** — Phase C (1/N) DOCX 1급 시민 codec — `src/PolyDoc.Codecs.Docx`. DocumentFormat.OpenXml 3.5.1 기반 reader/writer. 단락 / Heading1~6 / 정렬(좌·중·우·양쪽·균등) / 굵게·기울임·밑줄·취소선·위첨자·아래첨자 / 폰트 패밀리·크기 / 색상 / 기본 리스트 / Title·Author 코어 속성 라운드트립. xUnit 라운드트립 6건 + 스모크 1건 그린.
- **Added** — `tests/PolyDoc.Codecs.Docx.Tests` xUnit 라운드트립 테스트 6건.
- **Added** — `Directory.Packages.props` 에 DocumentFormat.OpenXml 3.5.1 (MIT) 등록.
- **Changed** — `PolyDoc.Codecs.Markdown` 의 reader 를 BCL 직접 파서에서 **Markdig 0.42.0** (BSD-2-Clause) 으로 교체. 풀 CommonMark 파싱 — 단락 / 헤더 / 리스트 / 강조에 더해 인라인 코드(monospace 힌트), 코드블록(fenced/indented), 인용(현재는 단락으로 격하), 링크(밑줄 표시) 처리. 라운드트립 호환성 유지. xUnit 11건(기존 6 + 신규 5) 그린.
- **Changed** — `PolyDoc.App` 의 `DocumentFormat` 서비스가 DOCX 를 **외부 컨버터 위탁 목록에서 제거**하고 직접 처리 대상으로 등록. 이제 메인 앱에서 `.docx` 를 native 로 읽고 쓴다. 외부 컨버터 위탁은 HWP / HWPX / DOC / HTML / HTM 만 남는다.
- **Added** — 핸텍 공식 회사 로고/아이콘 자산: `assets/Handtech_1024.png` (1024×1024 PNG), `assets/Handtech.ico` (멀티 사이즈 Windows ICO).
- **Added** — `src/PolyDoc.App` WPF 앱 첫 사이클 (Phase B1~B4 골격). `net10.0-windows` + WPF + CommunityToolkit.Mvvm 8.4.0. 메인 윈도우(파일/편집/입력/서식/도구/도움말 메뉴), TextBox 본문 편집기, 상태 바, About 다이얼로그(로고·핸텍·노진문·버전 표시). 파일 메뉴는 IWPF/MD/TXT 직접 처리, 외부 포맷은 Phase D 안내 메시지. Ctrl+N/O/S, Ctrl+Shift+S 단축키. 한국어 UI(.resx 분리는 다음 사이클로 이연).
- **Added** — Light 기본 테마 (`src/PolyDoc.App/Themes/Light.xaml`) — 핸텍 브랜드 블루 기반.
- **Added** — `Directory.Packages.props` 에 CommunityToolkit.Mvvm 8.4.0 등록.
- **Internal** — Phase A 솔루션 골격: `PolyDoc.sln` + `src/PolyDoc.Core` + `src/PolyDoc.Iwpf` + `src/PolyDoc.Codecs.Text` + `src/PolyDoc.Codecs.Markdown` + 대응 `tests/*` xUnit 프로젝트 + `tools/PolyDoc.SmokeTest` 콘솔 러너. .NET 10 + Central Package Management.
- **Added** — `PolyDoc.Core` 공통 문서 모델 1차: `PolyDocument`, `DocumentMetadata`, `Section`/`PageSettings`, `Block`/`NodeStatus`, `Paragraph`/`ParagraphStyle`/`Alignment`/`OutlineLevel`/`ListMarker`/`ListKind`, `Run`/`RunStyle`/`Color`, `StyleSheet`, `Provenance`/`SourceAnchor`, `IDocumentReader`/`IDocumentWriter`/`IDocumentCodec`. 한글 조판용 `WidthPercent`(장평) / `LetterSpacingPx`(자간) 포함.
- **Added** — `PolyDoc.Iwpf` 1차 codec (writer/reader). ZIP+JSON 패키지(`manifest.json`, `content/document.json`, `content/styles.json`, 선택적 `provenance/source-map.json`). 매니페스트 SHA-256 해시 검증, packageType 검사, 위변조 거부.
- **Added** — `PolyDoc.Codecs.Text` (TXT in/out, BOM 자동 감지).
- **Added** — `PolyDoc.Codecs.Markdown` 1차 codec — Markdig 의존 없이 BCL 만으로 ATX 헤더(#~######), 순서/비순서 리스트, `**굵게**`, `*기울임*` 인라인을 처리하는 실용 서브셋. 추후 Markdig 도입 시 교체 예정.
- **Added** — `PolyDoc.SmokeTest` BCL-only 콘솔 스모크 러너 (`tools/PolyDoc.SmokeTest`). 라운드트립 + 위변조 검출까지 4건 통과.
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
- **Internal** — 솔루션 파일을 `.slnx` (.NET 9~ XML 신형) 에서 `.sln` (전통 형식) 으로 교체. 사용자 Windows Visual Studio 빌드에서 `PolyDoc.Codecs.Docx` 프로젝트가 `.slnx` 의존성 그래프에서 누락되어 빌드 큐에 들어가지 않는 회귀가 발생 (다른 12개 프로젝트는 정상). `.sln` 형식은 모든 도구(VS, MSBuild, dotnet CLI, Rider)에서 안정적으로 인식되므로 호환성 우선.

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

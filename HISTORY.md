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
- **Added** — 핸텍 공식 회사 로고/아이콘 자산: `assets/Handtech_1024.png` (1024×1024 PNG), `assets/Handtech.ico` (멀티 사이즈 Windows ICO).
- **Added** — `src/PolyDoc.App` WPF 앱 첫 사이클 (Phase B1~B4 골격). `net10.0-windows` + WPF + CommunityToolkit.Mvvm 8.4.0. 메인 윈도우(파일/편집/입력/서식/도구/도움말 메뉴), TextBox 본문 편집기, 상태 바, About 다이얼로그(로고·핸텍·노진문·버전 표시). 파일 메뉴는 IWPF/MD/TXT 직접 처리, 외부 포맷은 Phase D 안내 메시지. Ctrl+N/O/S, Ctrl+Shift+S 단축키. 한국어 UI(.resx 분리는 다음 사이클로 이연).
- **Added** — Light 기본 테마 (`src/PolyDoc.App/Themes/Light.xaml`) — 핸텍 브랜드 블루 기반.
- **Added** — `Directory.Packages.props` 에 CommunityToolkit.Mvvm 8.4.0 등록.
- **Internal** — Phase A 솔루션 골격: `PolyDoc.slnx` + `src/PolyDoc.Core` + `src/PolyDoc.Iwpf` + `src/PolyDoc.Codecs.Text` + `src/PolyDoc.Codecs.Markdown` + 대응 `tests/*` xUnit 프로젝트 + `tools/PolyDoc.SmokeTest` 콘솔 러너. .NET 10 + Central Package Management.
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

### Known Limitations
- 본 개발 환경에서 NuGet.org(`api.nuget.org`)이 503 으로 차단되어 xUnit / Markdig 등 패키지 복원 불가. xUnit 테스트와 Markdig 도입은 Windows 환경(또는 NuGet 접근 가능한 환경)에서 검증·전환 필요. **사용자 게이트 G2** 에서 `dotnet test` 정상 동작 확인 예정.

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

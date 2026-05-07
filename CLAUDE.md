# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

Claude Code가 이 저장소에서 작업할 때 참고하는 가이드. 자세한 내용은 `README.md`(사용자 안내), `IWPF.md`(파일 포맷 사양), `HISTORY.md`(변경 이력), `WORK_PLAN.md`(다단계 작업 계획)를 본다.

## 프로젝트 개요

**PolyDonky** — HWP, HWPX, DOC, DOCX, HTML/HTM, MD, TXT 문서를 읽어 자체 포맷 **IWPF**로 저장하는 워드프로세서 앱.

> **이름의 유래**: **Poly**(gon) + **Donky**(당나귀). 다각형으로 거칠게 빚어 외형은 엉성해도,
> 당나귀처럼 어떤 짐(문서 포맷)이든 가리지 않고 먹어치우고 운반한다는 뜻. 멀티 포맷 ingest 가
> 프로젝트의 정체성이라는 점을 그대로 담은 작명이며, 코드·문서·UI 톤은 이 정체성(거칠지만
> 강건한 다중 포맷 처리기)을 일관되게 유지한다.
> _English: "A donkey roughly sculpted from polygons — clumsy on the outside,_
> _with a voracious appetite for any document format."_

- 회사: **핸텍 (HANDTECH)**
- 메인테이너 / 저작권자: **노진문 (Noh JinMoon)** — GitHub `@elrang3843`
- 언어/UI: **C# + WPF** (WinUI 3 가 아니라 WPF 로 확정)
- 대상 OS: **Windows 10 이상**
- 다국어: 한국어(기본), 영어
- 라이선스: Apache 2.0 — `LICENSE` / `NOTICE`
- 회사 로고: `assets/Handtech_1024.png` (1024×1024 PNG)
- 앱 아이콘: `assets/Handtech.ico` (Windows 멀티 사이즈 ICO)

## 핵심 아키텍처 원칙

### 1. 정본은 항상 IWPF
- HWP/DOC/HWPX/DOCX/HTML 등은 **import/export edge format**일 뿐, canonical master는 IWPF다.
- 내부 편집·검색·분석은 모두 IWPF의 공통 모델 위에서 수행한다.
- "외부 앱이 저장해도 보존 캡슐이 살아남는다"고 기대하지 말 것.

### 2. 2계층 설계 (반드시)
1. **의미 계층(공통 문서 모델)** — 편집/검색/분석용. 최소공배수가 아니라 **최대상한(superset)** 으로 잡는다.
2. **충실도 계층(fidelity capsule + 원본 내장 + provenance map)** — 역변환·무손실 보장용.

공통 모델만으로는 포맷 고유 기능을 못 살리므로 두 층 모두 필수.

### 3. 외부 변환 모듈 분리 (CLI)
- IWPF / MD / TXT 만 메인 앱에서 직접 read/write.
- HWP, HWPX, DOC, DOCX, HTML, HTM 의 read/write 는 **별도의 command-line 컨버터 앱/모듈**로 분리해 호출한다.
- 다른 형식으로 저장 시 사용자에게 **항상 한 번 더 확인**.

### 4. 변환 품질 목표
- 외부 포맷 import 시 원본 레이아웃 **99% 보존** 목표.
- HWPX ↔ IWPF ↔ DOCX 라운드트립이 1급 시민.

## IWPF 패키지 구조

ZIP 기반 OPC 유사 컨테이너. UTF-8, SHA-256 해시, 리소스 분리 저장.

```
iwpf/
  manifest.json
  content/         document.xml, styles.xml, numbering.xml,
                   sections.xml, annotations.xml
  resources/       images/, ole/, fonts/
  fidelity/
    original/      source.doc, source.hwp, ...   ← 원본 파일 내장(무손실 보증)
    capsules/      msdoc/, hancom/, ooxml/, hwpx/  ← 포맷별 보존 캡슐
  provenance/      source-map.json   ← 노드 ↔ 원본 위치 매핑
  render/          layout-snapshot.json, preview.pdf
  signatures/
```

### 공통 모델 필수 항목
문서 메타데이터 / 섹션·페이지 / 문단·문자 런 / 스타일 계층 / 번호·개요·목차 /
표(병합·셀 속성) / 이미지·도형·텍스트박스 / 머리말·꼬리말 / 각주·미주 /
책갈피·하이퍼링크·필드 / 주석·변경추적 / 수식 / 임베디드 개체 / 양식·컨트롤 /
숨김·보호 / 호환성 옵션 / **활성 콘텐츠는 격리 저장**.

### 한글 조판 특화 (절대 빼지 말 것)
줄격자, 문자단위 배치, 장평, 자간, 정렬 미세 규칙, 리스트/개요 번호 생성 규칙,
탭·들여쓰기·문단 줄바꿈 세부 동작, 개체 anchoring, 폰트 대체 규칙.

### Provenance / Dirty Tracking
각 노드에 `clean` / `modified` / `opaque` / `degraded` 상태를 기록.
수정되지 않은 노드는 원본 조각을 재사용하는 **하이브리드 export** 가 가능해야 함.

### Opaque Island 정책
완전히 이해 못 한 개체는 버리지 말고 **opaque object**로 보존 — 에디터에서는 read-only,
HWPX export 시 원형에 가깝게 직렬화, DOCX export 시 시각 대체(이미지·텍스트박스).

## 개발 단계 (구현 순서)

1. **1단계** — DOCX, HWPX **1급 시민** 지원. import/export + IWPF save/load + 기본 편집.
2. **2단계** — DOC, HWP **ingest 전용** 추가. 내부 포맷으로 정규화 후 출력은 HWPX/DOCX.
3. **3단계** — 고급 기능: 변경추적, 주석, 수식, 텍스트박스/도형, 필드/목차, 고급 표, 특수 조판.

## 메뉴 구조 (요약, 자세한 건 README.md)

파일 / 편집 / 입력(글상자·표·그래프·특수문자·수식·이모지·도형·그림·사인) /
서식(글자·문단·페이지) / 도구(설정·사전·맞춤법·사인 만들기) / 도움말.

- 테마 다수 지원(대상 연령: 학생~장년).
- 룰러·눈금·편집용지 보기 옵션.
- 도움말/라이선스 문서는 **언어별 별도 파일**.

## 절대 피할 것

1. 내부 포맷을 **DOCX XML 그대로** 쓰지 말 것 — HWPX/HWP 특성을 담을 수 없다.
2. 내부 포맷을 **HTML 기반**으로 잡지 말 것 — 워드프로세서 기능이 계속 새어나간다.
3. 외부 앱이 저장해도 앱 전용 확장이 보존될 거라고 기대하지 말 것.

## 변환 품질을 좌우하는 4가지

1. **폰트 통제** — fingerprint, 대체 테이블, 기준 폰트 세트 고정.
2. **Deterministic serializer** — 같은 입력·옵션이면 항상 같은 바이트.
3. **Source mapping / dirty tracking** — exporter가 똑똑해지는 전제.
4. **테스트 코퍼스** — 일반 본문, 복잡한 표, 머리말/꼬리말, 각주, 도형, 변경추적,
   목차/필드, 다단, 수식, 한글 조판 강한 문서, 레거시 DOC/HWP 샘플.
   검증 항목: 텍스트 동일성, 스타일/구조 동일성, 페이지 수, 위치 오차, PDF 렌더 비교.

## 활성 콘텐츠·서명

- 매크로/스크립트(VBA 등)는 **격리 저장**, 실행 정책은 별도.
- 전자서명은 수정 시 효력 깨짐 — **보존 증적**으로만 별도 관리.

## 솔루션 구조

`PolyDonky.sln` — .NET 10 (`global.json`이 SDK `10.0.107` 핀, 라이브러리는 `net10.0`,
WPF 앱은 `net10.0-windows`). 중앙 패키지 관리(`Directory.Packages.props`),
`TreatWarningsAsErrors=true`, `Nullable=enable` (`Directory.Build.props`).

```
src/
  PolyDonky.Core/             공통 문서 모델 — Document/Section/Paragraph/Run/Block/Table/
                              FloatingObject/StyleSheet/Provenance, IDocumentCodec, JSON 직렬화
  PolyDonky.Iwpf/             IWPF ZIP 패키지 reader/writer, manifest, 암호화, write-lock
  PolyDonky.Codecs.Text/      TXT codec
  PolyDonky.Codecs.Markdown/  MD codec (Markdig)
  PolyDonky.Codecs.Docx/      DOCX codec (DocumentFormat.OpenXml) — 1급 시민, 메인 앱 직접 처리
  PolyDonky.Codecs.Hwpx/      HWPX codec (자체 구현, KS X 6101) — 1급 시민
  PolyDonky.Codecs.Html/      HTML5 codec (AngleSharp) — 메인 앱 직접 처리
  PolyDonky.Codecs.Xml/       XML / XHTML5 codec (Codecs.Html 위에 polyglot serializer)
  PolyDonky.App/              WPF 데스크톱 앱 (AssemblyName=PolyDonky)
                              MVVM (CommunityToolkit.Mvvm), Views/ViewModels/Services/Themes
                              FlowDocument 기반 에디터 — FlowDocumentBuilder/Parser/Search,
                              PageViewBuilder, PerPageEditorHost
                              i18n: Properties/Resources.resx + Resources.en-US.resx
tests/                        프로젝트별 xUnit 테스트 (.Tests 짝)
tools/PolyDonky.SmokeTest/    콘솔 스모크 — 모든 codec + IWPF round-trip 통합 검증
```

CLI 변환 모듈 분리 원칙(아키텍처 §3) — **메인 앱은 IWPF/MD/TXT 만 직접 read/write**,
그 외 모든 포맷은 **포맷별로 별도 CLI 실행 파일**로 분리한다.

현재 구현된 분리 대상:
- `tools/PolyDonky.Convert.Html` — HTML ↔ IWPF
- `tools/PolyDonky.Convert.Xml`  — XML/XHTML ↔ IWPF
- `tools/PolyDonky.Convert.Docx` — DOCX ↔ IWPF
- `tools/PolyDonky.Convert.Hwpx` — HWPX ↔ IWPF

메인 앱은 `Codecs.Html`/`Codecs.Xml`/`Codecs.Docx`/`Codecs.Hwpx` 를
ProjectReference 하지 **않으며**, 빌드 시 CLI 출력(.dll + 부속 파일) 이 메인 앱
출력 디렉터리로 복사되고, 런타임에 `Services/ExternalConverter.cs` 가 spawn 한다.
열기: 입력 → 같은 이름의 정식 `.iwpf` 변환 후 메인 앱이 IWPF 를 읽음 (CurrentFilePath = .iwpf).
저장: 같은 이름의 IWPF 정본 저장 → CLI 가 외부 포맷으로 변환 (두 파일 모두 디스크에 남음).
2단계 추가 대상인 HWP/DOC 도 동일한 패턴으로 `tools/PolyDonky.Convert.Hwp` / `Doc` 로 분리할 예정.

## 빌드·테스트 명령

`PolyDonky.App` 은 `net10.0-windows`(WPF) 라 비-Windows 환경에서는 빌드되지 않는다.
라이브러리·코덱·테스트는 크로스 플랫폼 빌드 가능.

```powershell
# 전체 복원 / 빌드 / 테스트
dotnet restore PolyDonky.sln
dotnet build   PolyDonky.sln -c Debug
dotnet test    PolyDonky.sln -c Debug

# 단일 프로젝트
dotnet build src/PolyDonky.Iwpf/PolyDonky.Iwpf.csproj
dotnet test  tests/PolyDonky.Iwpf.Tests/PolyDonky.Iwpf.Tests.csproj

# 단일 테스트 (xUnit FQN 또는 DisplayName 필터)
dotnet test tests/PolyDonky.Iwpf.Tests --filter "FullyQualifiedName~IwpfWriterTests.RoundTrip"
dotnet test tests/PolyDonky.Iwpf.Tests --filter "DisplayName~round-trip"

# 앱 실행 (Windows 전용)
dotnet run --project src/PolyDonky.App

# 모든 codec + IWPF 통합 스모크 (콘솔)
dotnet run --project tools/PolyDonky.SmokeTest
```

**Linux / CI 환경에서 WPF 프로젝트 빌드 시** `EnableWindowsTargeting=true` 플래그가 필요하다.
라이브러리·코덱·테스트는 해당 플래그 없이도 빌드된다.

```bash
# Linux 에서 전체 솔루션 빌드 (WPF 포함)
dotnet build PolyDonky.sln -c Debug -p:EnableWindowsTargeting=true
# 라이브러리·테스트만 (플래그 불필요)
dotnet build src/PolyDonky.Core/PolyDonky.Core.csproj -c Debug
dotnet test  PolyDonky.sln -c Debug  # App.Tests 제외하면 Linux 에서도 통과
```

CI 워크플로: `.github/workflows/dotnet.yml`(라이브러리/테스트),
`.github/workflows/dotnet-desktop.yml`(WPF 앱). 분석기 경고는 빌드 오류로 승격되므로
무시하지 말고 수정한다.

## WPF 앱 내부 구조

### PaperHost 캔버스 레이어 스택 (z-순서, 아래→위)

PaperHost(`Grid`) 는 아래 Canvas 들을 겹쳐 단일 좌표 공간을 구성한다. 히트 테스트는 위에서 아래로 전달.

```
PageBackgroundCanvas   (IsHitTestVisible=false) — 페이지 테두리·그림자·여백 가이드
HeaderFooterCanvas     (IsHitTestVisible=false) — 머리말/꼬리말 레이블
UnderlayImageCanvas    — BehindText 이미지
UnderlayShapeCanvas    — BehindText 도형
UnderlayTableCanvas    — BehindText 표
WatermarkCanvas        (IsHitTestVisible=false) — 워터마크
PerPageEditorHost      — 페이지별 RichTextBox (FloatingCanvas 포함)
FloatingCanvas         — 본문 내 Float 개체
OverlayImageCanvas     — InFrontOfText 이미지
OverlayShapeCanvas     — InFrontOfText 도형
OverlayTableCanvas     — InFrontOfText 표 / 고정 표
DrawPreviewCanvas      (IsHitTestVisible=false) — 도형 그리기 미리보기
TypesettingMarksCanvas (IsHitTestVisible=false) — 조판 기호
```

모든 오버레이 컨트롤의 `.Tag`는 해당 Core 모델(`ImageBlock`, `ShapeObject`, `Table` 등)을 담는다.
오버레이 좌표계는 PaperHost 기준 DIP 단위(`FlowDocumentBuilder.MmToDip`/`DipToMm`).

### 마우스 우클릭 통합 핸들러

`OnPaperPreviewMouseRightButtonDown` (MainWindow.xaml.cs) 이 PaperHost 의 **PreviewMouseRightButtonDown**(터널링) 을 잡아 **모든** 우클릭을 처리한다. 우선순위 체인:

```
① 폴리선/스플라인 입력 모드 → OpenPolylineInputMenu()
①-b 도형 편집 핸들(Rectangle, _shapeEditHandles 소속) → OnShapeEditHandleRightClicked()
② 오버레이 도형 (Overlay→Underlay ShapeCanvas) → BuildShapeMenu()
③ 오버레이 이미지 (Overlay→Underlay ImageCanvas) → BuildOverlayImageMenu()
④ 오버레이 표   (Overlay→Underlay TableCanvas)  → BuildOverlayTableMenu()
⑤ BodyEditor 본문 → ContextMenuOpening 에 위임 (e.Handled = false)
```

새 우클릭 동작을 추가할 때는 이 체인에 else-if/priority 블록을 추가하고
**개별 ContextMenu 를 오버레이 컨트롤에 직접 붙이지 않는다** (주석 참조).

### MainWindow 부분 클래스 분리

- `Views/MainWindow.xaml.cs` — 문서 로드/저장, 오버레이 배치, 마우스 핸들러, 메뉴 빌더 전체
- `Views/MainWindow.ShapeEdit.cs` — 도형 편집 핸들(`_shapeEditHandles`, 정점/세그먼트) + `OnShapeEditHandleRightClicked`

도형 편집 코드 수정 시 두 파일을 함께 본다.

### 핵심 이름 주의사항

- 공통 문서 모델 클래스명은 **`PolyDonkyument`** (`PolyDonky.Core` 네임스페이스). `Document`가 아님.
- 모델 ↔ FlowDocument 변환 단위 함수: `FlowDocumentBuilder.MmToDip`, `DipToMm`, `PtToDip`, `DipToPt`.

## 작업 시 유의사항

- 문서/UI 문자열은 i18n 가능하게 분리(한국어 기본, 영어 병행).
  `PolyDonky.App/Properties/Resources.resx` 와 `Resources.en-US.resx` 를 짝으로 갱신.
  사용자 노출 문자열 하드코딩 금지.
- 파일 I/O 단위 테스트는 위 "테스트 코퍼스" 카테고리를 기준으로 설계.
- 새 외부 포맷 지원 추가 시: ① 공통 모델 매핑 ② fidelity capsule 정의 ③ provenance 기록
  ④ 라운드트립 테스트 추가 — 네 단계를 한 세트로 진행.
- 새 codec 은 `PolyDonky.Core.IDocumentCodec` 을 구현하고
  `tools/PolyDonky.SmokeTest` 에 round-trip 케이스를 추가한다.
- 에디터는 WPF `FlowDocument` 기반이다. 모델 ↔ FlowDocument 변환은
  `Services/FlowDocumentBuilder` · `FlowDocumentParser` 가 담당하므로,
  `PolyDonky.Core` 에 새 블록/런 타입을 추가하면 두 곳을 모두 갱신해야 하고
  검색 경로(`FlowDocumentSearch`) 도 함께 본다.

## 변경 이력 관리

- 사용자·기여자에게 영향 있는 모든 변경은 같은 PR/커밋에서 **`HISTORY.md` 의 `## [Unreleased]` 섹션**
  에 한 줄을 추가한다.
- 카테고리: `Added` / `Changed` / `Deprecated` / `Removed` / `Fixed` / `Security` / `Docs` / `Internal`.
- 형식·릴리스 절차는 `HISTORY.md` 상단의 "작성 규칙" 을 따른다.

### 버전·릴리스 정책 (절대 어기지 말 것)

- **사용자가 명시적으로 "릴리즈" 를 지시하기 전까지 모든 빌드는 테스트 버전이다.**
  테스트 빌드 태그 형식: **`1.0.0-test.<n>`** (예: `1.0.0-test.1`).
- **최초 정식 릴리스는 `1.0.0`** 이며, 그 이후로는 일반 SemVer (1.0.1 / 1.1.0 / 2.0.0 ...) 를 따른다.
- 사용자의 명시적 지시 없이 다음을 절대 수행하지 않는다.
  - `## [Unreleased]` 를 `## [1.0.0]` 등 정식 버전 헤더로 승격
  - `v1.0.0` 등 정식 릴리스 태그 생성·푸시
  - GitHub Release 생성
- 사용자 지시가 있을 때의 절차:
  1. `[Unreleased]` 내용을 `## [1.0.0] - YYYY-MM-DD` 헤더로 승격하고 비어 있는 `[Unreleased]` 를 다시 만든다.
  2. 커밋 후 `git tag v1.0.0` (또는 테스트 컷이면 `v1.0.0-test.N`) 생성.
  3. 사용자가 푸시·릴리스 노트 작성을 별도 지시하면 그때 진행.

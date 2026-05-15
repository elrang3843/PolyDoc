# Work Plan — PolyDonky

이 문서는 **세션을 가로질러 작업을 이어받기 위한 운영 계획서**입니다.
한 세션이 끝나도 다음 세션이 이 문서를 읽고 동일한 맥락에서 이어 갈 수 있도록 갱신합니다.

- 사용자: 노진문 (Noh JinMoon)
- 회사: 핸텍 (HANDTECH)
- 메인 브랜치(작업 대상): `claude/install-dotnet-10-K5KaU`
- 정책: 사용자 명시 지시 전까지 모든 빌드는 **테스트 버전(`1.0.0-test.<n>`)**, 최초 정식 릴리스는 `1.0.0`.
- 운영 가정: 한 세션당 일일 사용량의 ~90% 까지 사용을 허용하되, 위험·비가역 결정은 사용자 게이트에서 멈춘다.

---

## 환경 사실관계

| 항목 | 상태 |
|---|---|
| 호스트 OS | Ubuntu 24.04 (root) |
| .NET SDK | **10.0.107** (apt: `dotnet-sdk-10.0`) |
| GUI/Display | 없음 — WPF UI 빌드·실행·스크린샷 불가 |
| 타겟 OS | Windows 10/11 x64 |
| Git remote | `elrang3843/PolyDonky` |
| **NuGet.org** | **정상**. xUnit·Markdig·OpenXml SDK·AngleSharp 등 외부 패키지 복원 가능 |

**핵심 제약**:
- WPF 앱(`Microsoft.NET.Sdk.WindowsDesktop`)은 Windows 에서만 빌드 가능.
- 이 환경에서는 **WPF 외 모든 라이브러리·코덱·xUnit 테스트** 검증 가능. UI 시각 검증만 사용자 책임.
- Linux 에서 전체 솔루션 빌드: `dotnet build PolyDonky.sln -p:EnableWindowsTargeting=true`

---

## 기술 스택 (확정)

| 항목 | 값 |
|---|---|
| .NET | **10.0** (LTS) |
| UI | **WPF** |
| MVVM | CommunityToolkit.Mvvm (MIT) |
| 테스트 | xUnit (Apache 2.0) — Assertion 은 xUnit 내장만 사용 (FluentAssertions v8 라이선스 회피) |
| Markdown | Markdig 0.42.0 (BSD-2-Clause) |
| DOCX | DocumentFormat.OpenXml 3.5.1 (MIT) |
| HTML/XML | AngleSharp (MIT) |
| HWPX | 자체 구현 (KS X 6101) |
| HWP·DOC | v1.0.0 이후 자체 CLI 파서 예정 |
| 직렬화 | System.Text.Json |

라이선스는 모두 Apache 2.0 호스트 프로젝트와 호환.

---

## 솔루션 레이아웃

```
PolyDonky/
├── PolyDonky.sln
├── Directory.Build.props
├── Directory.Packages.props      # central package management
├── global.json                   # SDK pin (10.0.x)
├── src/
│   ├── PolyDonky.Core/             # 공통 문서 모델 (POCO)
│   ├── PolyDonky.Iwpf/             # IWPF 패키지 codec (ZIP+JSON)
│   ├── PolyDonky.Codecs.Text/      # TXT codec
│   ├── PolyDonky.Codecs.Markdown/  # MD codec (Markdig)
│   ├── PolyDonky.Codecs.Docx/      # DOCX codec (1급 시민)
│   ├── PolyDonky.Codecs.Hwpx/      # HWPX codec (1급 시민)
│   ├── PolyDonky.Codecs.Html/      # HTML5 codec (AngleSharp) — 완전 구현
│   ├── PolyDonky.Codecs.Xml/       # XML/XHTML5 codec (Html 위 polyglot) — 완전 구현
│   └── PolyDonky.App/              # WPF 데스크톱 앱 (net10.0-windows)
├── tests/
│   ├── PolyDonky.Core.Tests/
│   ├── PolyDonky.Iwpf.Tests/
│   ├── PolyDonky.Codecs.Text.Tests/
│   ├── PolyDonky.Codecs.Markdown.Tests/
│   ├── PolyDonky.Codecs.Docx.Tests/
│   ├── PolyDonky.Codecs.Hwpx.Tests/
│   ├── PolyDonky.Codecs.Html.Tests/
│   ├── PolyDonky.Codecs.Xml.Tests/
│   └── PolyDonky.App.Tests/        # net10.0-windows — Windows 전용
└── tools/
    ├── PolyDonky.SmokeTest/        # 콘솔 스모크 — 전 codec + IWPF 통합 검증
    ├── PolyDonky.Convert.Html/     # HTML ↔ IWPF CLI 컨버터
    ├── PolyDonky.Convert.Xml/      # XML/XHTML ↔ IWPF CLI 컨버터
    ├── PolyDonky.Convert.Docx/     # DOCX ↔ IWPF CLI 컨버터
    └── PolyDonky.Convert.Hwpx/     # HWPX ↔ IWPF CLI 컨버터
```

---

## Phase 진행표

체크박스: ☐ 미진행 / ◑ 진행중 / ✅ 완료

### Phase A — Core 라이브러리 (Linux 전수 가능)
- ✅ A1 솔루션 스캐폴딩 (PolyDonky.sln, .NET 10, CPM, src + tests + tools/SmokeTest)
- ✅ A2 PolyDonky.Core (공통 문서 모델: Block/Paragraph/Run/Table/ImageBlock/ShapeObject/FloatingObject/StyleSheet/PageSettings/Provenance 등)
- ✅ A3 PolyDonky.Iwpf (reader/writer, ZIP+JSON, SHA-256 검증, 위변조 거부, 암호화, write-lock)
- ✅ A4 PolyDonky.Codecs.Text (TXT in/out, BOM 감지)
- ✅ A5 PolyDonky.Codecs.Markdown (Markdig 0.42.0, CommonMark 풀 파싱)
- ✅ A6 단위 테스트 — 전 프로젝트 그린
- ✅ A7 커밋·푸시 / G1 사용자 확인 완료

### Phase B — WPF UI 셸 (Windows 필수, UI 시각 검증은 사용자 책임)
- ✅ B1 PolyDonky.App 스캐폴딩 (net10.0-windows + WPF + CommunityToolkit.Mvvm 8.4.0, Handtech.ico, Handtech_1024.png 임베드)
- ✅ B2 메인 메뉴 6단 (파일/편집/입력/서식/도구/도움말) + 상태 바 + About + LicenseInfo + UserGuide 다이얼로그
- ✅ B3 FlowDocument 기반 본문 편집기 — FlowDocumentBuilder/Parser/Search, PolyDonkyument 양방향 동기화
- ✅ B4 PaperHost 캔버스 레이어 스택 (PageBackgroundCanvas ~ TypesettingMarksCanvas) + PageViewBuilder + PerPageEditorHost (페이지 분할 편집)
- ✅ B5 Undo/Redo (UndoRedoManager)
- ✅ B6 찾기/바꾸기 (FindReplaceWindow)
- ✅ B7 테마 시스템 (ThemeService, 다중 테마 — 학생/청년/장년 대상)
- ✅ B8 i18n — LanguageService + LocalizedStrings, 한국어 기본 / 영어 병행
- ✅ B9 설정 다이얼로그 (SettingsWindow)
- ✅ B10 인쇄 미리보기 (PrintPreviewWindow)
- ✅ B11 IWPF 암호화 — PasswordPromptWindow / PasswordChangeWindow
- ✅ G2 Windows 빌드·실행 통과 (사용자 보고)

### Phase C — 코덱 완성 (1급 시민 포맷 + HTML/XML)
- ✅ C1 DOCX reader — 단락·헤더·정렬·강조·폰트·색상·리스트·표·이미지·OpaqueBlock
- ✅ C2 DOCX writer + 라운드트립 테스트
- ✅ C3 HWPX reader — KS X 6101 자체 구현, ZIP entry 정규화, BOM-aware, style/charPr/paraPr 매핑
- ✅ C4 HWPX writer + 라운드트립 테스트
- ✅ C5 HWPX 표·이미지 양방향 (`<hp:tbl>` ↔ Table, `<hp:pic>`+BinData ↔ ImageBlock, SHA-256 dedupe)
- ✅ C6 HTML5 codec — AngleSharp 기반 완전 구현
  - 전 블록/인라인 타입 (Paragraph, Table, ImageBlock, ShapeObject, TocBlock, OpaqueBlock 등)
  - CSS 클래스 기반 직렬화 (pd-* 클래스), 인라인 스타일 fallback
  - SVG 도형 양방향 — ShapeObject ↔ `<svg>` (rect/ellipse/circle/line/polyline/polygon/path)
  - 경로 파싱: M/C/Z 명령어, Catmull-Rom→cubic Bezier 역변환
  - MathML → 수식 OpaqueBlock, dl/dt/dd, details/summary, form 요소, page-break div, TOC nav
  - 편집용지 직렬화 — pd-page-* meta 태그 + CSS `@page` 규칙 (size, margin)
  - 용지 정보 없을 때 기본값: A4 세로, 기본 여백
- ✅ C7 XML/XHTML5 codec — Html codec 위 polyglot serializer 완전 구현
  - XHTML5 자체 닫기 태그, XmlReader → HtmlReader 위임, 편집용지 동일 처리
- ✅ G3 DOCX/HWPX/HTML/XML 검증 — 테스트 전 통과, HWPX 한컴 오피스 시각 검증 대기

### Phase D — 외부 CLI 컨버터 분리
- ✅ D1 tools/PolyDonky.Convert.Html — HTML ↔ IWPF CLI
- ✅ D2 tools/PolyDonky.Convert.Xml  — XML/XHTML ↔ IWPF CLI
- ✅ D3 tools/PolyDonky.Convert.Docx — DOCX ↔ IWPF CLI
- ✅ D4 tools/PolyDonky.Convert.Hwpx — HWPX ↔ IWPF CLI
- ✅ D5 Services/ExternalConverter.cs — 메인 앱 ↔ CLI IPC (spawn, 인자/표준입출력/exit code)

### Phase E — 편집 기능 (WPF App)
- ✅ E1 표 삽입/편집 (TableInsertDialog, TablePropertiesWindow, CellPropertiesWindow)
- ✅ E2 이미지 삽입/편집 (ImageWindow, ImagePropertiesWindow)
- ✅ E3 도형/폴리선 (ShapePropertiesWindow, MainWindow.ShapeEdit.cs, shape edit handles, 오버레이 레이어)
- ✅ E4 텍스트박스 (TextBoxOverlay, TextBoxPropertiesWindow, TextBoxColumnHost/Layout)
- ✅ E5 각주·미주 (FootnoteEditorPanel, 플로팅 캔버스 통합)
- ✅ E6 하이퍼링크 (HyperlinkDialog)
- ✅ E7 특수 문자 (SpecialCharWindow)
- ✅ E8 이모지 (EmojiWindow, EmojiPropertiesWindow)
- ✅ E9 수식 (EquationWindow — LaTeX/MathML 기반)
- ✅ E10 편집용지/여백 (PageFormatWindow)
- ✅ E11 글자 서식 (CharFormatWindow)
- ✅ E12 단락 서식 (ParaFormatWindow)
- ✅ E13 개요 서식 (OutlineStyleWindow)
- ✅ E14 문서 정보 (DocumentInfoWindow)
- ✅ E15 사전 (DictionaryWindow)
- ✅ E16 페이지 나누기 (PageBreakPadding)
- ✅ E17 머리말/꼬리말 (PerPageEditorHost, PageViewBuilder 통합)
- ✅ E19 목차 자동 생성 (전 페이지 스캔 + 페이지 번호 삽입 구현)
- ✅ E20 필드 코드 자동 갱신 — FieldRenderContext 도입, 페이지네이션 시 Page/NumPages/Author/Title 실제 값 반영

### Phase F — RTF/HWP ingest

**결정**: LibreOffice 미사용. RTF 는 자체 구현(`DocWriter`), HWP·DOC 는 v1.0.0 이후 자체 CLI 파서로 진행.

**F1 RTF export/import** ✅
- `tools/PolyDonky.Convert.Doc` CLI: IWPF ↔ RTF (`DocWriter`/`DocReader` 자체 구현)
- `ExternalConverter.GetConverter("rtf")` 연결, `KnownFormats.OpenFilter`/`SaveFilter` 갱신
- 글자/단락 서식·위첨자/아래첨자·표·이미지·메타데이터 완전 지원
- 도형(`\shp`) 아웃라인 지원: 위치·크기·종류·채우기/선 색상 ✅
- OLE 개체(`\object`) 아웃라인 지원: OpaqueBlock 보존 ✅

**F1-후속 RTF 도형/OLE 전체 지원** ☐ (v1.0.0 이후)
- `\shp` 전체 속성: 그림자·3D·곡선 경로(polyline/spline)·텍스트 레이아웃 등
- `\object` OLE 데이터 완전 복원: 바이너리 역직렬화 + 뷰어 연동

**F2 HWP import/export** ☐ (v1.0.0 이후)
- `tools/PolyDonky.Convert.Hwp` CLI: 자체 HWP 파서 구현 예정

**F3 Opaque island 정책 전면 적용** ☐
- 이해 못 한 개체를 read-only 보존, HWPX export 원형 직렬화

### Phase G — 고급 기능
- ◑ G-Themes 다중 테마 — ThemeService 구현 완료, 테마 파일 추가 필요

### Phase H — 인스톨러 / `1.0.0` 릴리즈
- ☐ H1 MSIX 인스톨러 패키징
- ☐ G5 사용자 명시 지시 후 `1.0.0` 컷 — `[Unreleased]` → `[1.0.0]` 승격, `v1.0.0` 태그

---

## 현재 테스트 현황 (Linux 환경, 2026-05-07 기준)

| 프로젝트 | 테스트 수 | 상태 |
|---|---|---|
| PolyDonky.Core.Tests | 40 | ✅ |
| PolyDonky.Iwpf.Tests | 13 | ✅ |
| PolyDonky.Codecs.Text.Tests | 5 | ✅ |
| PolyDonky.Codecs.Markdown.Tests | 40 | ✅ |
| PolyDonky.Codecs.Docx.Tests | 24 | ✅ |
| PolyDonky.Codecs.Hwpx.Tests | 52 | ✅ |
| PolyDonky.Codecs.Html.Tests | 95 | ✅ |
| PolyDonky.Codecs.Xml.Tests | 30 | ✅ |
| **Linux 합계** | **299** | **All green** |
| PolyDonky.App.Tests | (WPF — Windows 전용) | 별도 검증 |

---

## 사용자 디버깅·컨펌 게이트

| 게이트 | 시점 | 사용자 작업 | 상태 |
|---|---|---|---|
| G0 | 시작 전 | 기술 스택 승인 | ✅ 완료 |
| G1 | Phase A 종료 | PR 리뷰 후 머지 | ✅ 완료 |
| G2 | Phase B 시작 | Windows build/run 검증 | ✅ 완료 |
| G3 | Phase C 종료 | HWPX·DOCX·HTML·XML 시각 검증 | ◑ HWPX writer 한컴 검증 대기 |
| G4 | Phase F 진입 시 | HWP/DOC 변환 노선 확정 | ✅ **LibreOffice 미사용, 자체 CLI 파서로 결정** |
| G5 | Phase H 진입 시 | "릴리즈하자" 명시 — `1.0.0` 컷 | ☐ 미진입 |

---

## 현재 인수인계

### 완료 요약
- **Phase A** ✅ — Core 모델, IWPF, TXT, Markdown 코덱, xUnit 전체 그린, G1 통과.
- **Phase B** ✅ — WPF 앱 완전 구현. 메인 윈도우·메뉴 6단·FlowDocument 편집기·PaperHost 레이어·Undo/Redo·찾기/바꾸기·테마·i18n·설정·인쇄 미리보기·암호화·전 편집 다이얼로그. G2 통과.
- **Phase C** ✅ — DOCX/HWPX 1급 시민, HTML5/XML 코덱 완전 구현 (SVG 도형 양방향, 편집용지 직렬화 포함). 274개 Linux 테스트 전 그린.
- **Phase D** ✅ — CLI 컨버터 4종 (Html/Xml/Docx/Hwpx) 분리, ExternalConverter.cs IPC 서비스.
- **Phase E (부분)** ✅ — 표·이미지·도형·텍스트박스·각주·하이퍼링크·특수문자·이모지·수식·편집용지·글자서식·단락서식·머리말/꼬리말·페이지나누기·문서정보·개요서식·사전 Shell 구현.

### 미완료 / 다음 작업 후보

1. **G3 완료** — 사용자가 PolyDonky 가 만든 `.hwpx` 를 한컴 오피스에서 열어 시각 확인. 문제가 있으면 `C6` (HWPX writer 한컴 호환 향상 — header.xml 동적 charPr/paraPr 생성) 진행.

2. ~~**E19 목차 자동 생성**~~ ✅ 완료.

3. **F2 HWP ingest** — v1.0.0 이후. 자체 CLI 파서(`tools/PolyDonky.Convert.Hwp`) 구현.

4. **H1 MSIX 인스톨러** — 1.0.0 릴리즈 전 패키징.

### 알려진 한계·주의사항

- **HWPX writer 한컴 호환**: writer 는 스펙에 맞게 구현되어 있으나, 한컴 오피스의 비공개 확장 처리가 완벽하지 않을 수 있음. header.xml 의 charPr/paraPr 동적 생성이 부족한 경우 문서가 깨져 보일 수 있음.
- **Block 다형성**: `JsonDerivedType` 등록은 `src/PolyDonky.Core/Block.cs`. 새 블록 타입 추가 시 반드시 등록.
- **FlowDocument 표현 한계**: 장평·자간·Provenance 등은 FlowDocument 표현 불가. `FlowDocumentParser.Parse(fd, originalForMerge: _document)` 로 비파괴 보존 — 편집 후 Save 경로에서 반드시 호출.
- **WPF 빌드 검증**: Linux 환경에서 App 코드 변경 후 `csproj ProjectReference` 누락 및 namespace 충돌(`DocumentFormat.OpenXml` 과 충돌한 선례 있음) 을 반드시 점검.
- **이모지 유니코드**: EmojiWindow 에서 삽입한 이모지는 `Run.Text` 에 직접 저장. HWPX export 시 UTF-8 직렬화 범위 확인 필요.

### 사용자 작업이 필요한 항목

- [ ] **G3 HWPX writer 검증**: Windows 에서 PolyDonky 로 문서 작성 → `.hwpx` 저장 → 한컴 오피스에서 열어 본문·서식·표·이미지가 정상 표시되는지 확인.
- [ ] **G3 HTML/XML import 검증**: 외부 HTML 문서(웹 기사, 메모장 HTML 등)를 드래그&드롭 또는 열기 → 화면 표시 + IWPF 저장 + 재열기 라운드트립 확인.
- [ ] **App.Tests 검증**: Windows 에서 `dotnet test tests/PolyDonky.App.Tests` — WPF 전용 단위 테스트 그린 확인.

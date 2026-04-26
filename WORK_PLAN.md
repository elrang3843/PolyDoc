# Work Plan — PolyDoc

이 문서는 **세션을 가로질러 작업을 이어받기 위한 운영 계획서**입니다.
한 세션이 끝나도 다음 세션이 이 문서를 읽고 동일한 맥락에서 이어 갈 수 있도록 갱신합니다.

- 사용자: 노진문 (Noh JinMoon)
- 회사: 핸텍 (HANDTECH)
- 메인 브랜치(작업 대상): `claude/create-claude-guide-VK1Pz`
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
| Git remote | `elrang3843/PolyDoc` |
| **NuGet.org** | **정상** (이전 503 차단 풀림). xUnit·Markdig·OpenXml SDK 등 외부 패키지 복원 가능 |

**핵심 제약**:
- WPF 앱(`Microsoft.NET.Sdk.WindowsDesktop`)은 Windows 에서만 빌드 가능.
- 이 환경에서는 **WPF 외 모든 라이브러리·코덱·xUnit 테스트** 검증 가능. UI 시각 검증만 사용자 책임.

---

## 기술 스택 (확정)

| 항목 | 값 |
|---|---|
| .NET | **10.0** (LTS) |
| UI | **WPF** |
| MVVM | CommunityToolkit.Mvvm (MIT) |
| 테스트 | xUnit (Apache 2.0) — Assertion 은 xUnit 내장만 사용 (FluentAssertions v8 라이선스 회피) |
| Markdown | Markdig (BSD-2-Clause) |
| DOCX | DocumentFormat.OpenXml (MIT) |
| HTML | AngleSharp (MIT) |
| HWPX | 자체 구현 (KS X 6101) |
| HWP·DOC | LibreOffice headless 위탁 노선 우선, 어려우면 자체 |
| 직렬화 | System.Text.Json (Phase A), 필요 시 IWPF 일부는 XML 로 전환 |

라이선스는 모두 Apache 2.0 호스트 프로젝트와 호환.

---

## 솔루션 레이아웃

```
PolyDoc/
├── PolyDoc.sln
├── Directory.Build.props
├── Directory.Packages.props      # central package management
├── global.json                   # SDK pin (10.0.x)
├── .gitignore                    # .NET 표준 + IDE
├── src/
│   ├── PolyDoc.Core/             # 공통 문서 모델 (POCO)
│   ├── PolyDoc.Iwpf/             # IWPF 패키지 codec (ZIP+JSON)
│   ├── PolyDoc.Codecs.Text/      # TXT codec
│   ├── PolyDoc.Codecs.Markdown/  # MD codec (Markdig)
│   ├── PolyDoc.Codecs.Docx/      # (Phase C) DOCX codec
│   ├── PolyDoc.Codecs.Hwpx/      # (Phase C) HWPX codec
│   └── PolyDoc.App/              # (Phase B) WPF UI shell — Windows-only
├── tests/
│   ├── PolyDoc.Core.Tests/
│   ├── PolyDoc.Iwpf.Tests/
│   ├── PolyDoc.Codecs.Text.Tests/
│   └── PolyDoc.Codecs.Markdown.Tests/
└── samples/
    └── corpus/                   # 골든 테스트 코퍼스
```

---

## Phase 진행표

체크박스: ☐ 미진행 / ◑ 진행중 / ✅ 완료

### Phase A — Core 라이브러리 (Linux 전수 가능)
- ✅ A1 솔루션 스캐폴딩 (PolyDoc.sln, .NET 10, CPM, 4 src + 4 tests + tools/SmokeTest)
- ✅ A2 PolyDoc.Core (공통 문서 모델)
- ✅ A3 PolyDoc.Iwpf (reader/writer, ZIP+JSON, SHA-256 검증, 위변조 거부)
- ✅ A4 PolyDoc.Codecs.Text (TXT in/out, BOM 감지)
- ✅ A5 PolyDoc.Codecs.Markdown (Markdig 없이 BCL 서브셋: 헤더·리스트·강조)
- ◑ A6 단위 테스트 — xUnit 코드 작성 완료, **NuGet 차단으로 본 환경 미실행**. Windows 에서 G2 시 `dotnet test` 로 검증.
- ✅ A6b 콘솔 스모크 러너 4/4 통과 (라운드트립 3종 + 위변조 검출)
- ✅ A7 커밋·푸시

### Phase B — WPF UI 셸 (Windows 필수)
- ✅ B1 PolyDoc.App 스캐폴딩 (`net10.0-windows` + WPF + CommunityToolkit.Mvvm 8.4.0, ApplicationIcon=Handtech.ico, Handtech_1024.png 임베드)
- ✅ B2 메인 메뉴 6단(파일/편집/입력/서식/도구/도움말) + 본문 편집기 + 상태 바 + About 다이얼로그
- ✅ B2.5 본문 편집기를 **RichTextBox + FlowDocument** 로 업그레이드. FlowDocumentBuilder/Parser 로 PolyDocument 와 양방향 동기화, 한글 조판 등 비-FlowDocument 속성은 Tag 머지로 비파괴 보존. PolyDoc.App.Tests 프로젝트(net10.0-windows + xUnit) 신설, 9 라운드트립 테스트 작성.
- ◑ B3 i18n 한/영 — 1차 사이클은 한국어 하드코딩, `.resx` 리소스 분리는 다음 사이클로 이연
- ◑ B4 테마 시스템 — 1차 사이클은 Light 단일 테마(핸텍 브랜드 블루), 다중 테마는 다음 사이클
- ✅ B5 **G2** — Windows 머신에서 첫 `dotnet build` / `dotnet run` 통과. 메뉴 동작·About 표시·IWPF/MD/TXT/DOCX 열고 저장 정상 (사용자 보고).
- ☐ B5.5 **G2.5** — RichTextBox 업그레이드 후 Windows 검증: build/test 그린, .docx 의 폰트·크기·색·정렬이 화면에 표시되고 편집·저장이 보존되는지.

### Phase C — DOCX/HWPX 1급 시민 (M2-M3)
- ✅ C1 DOCX reader (OpenXml SDK 3.5.1) — 단락·헤더·정렬·강조·폰트·색상·리스트
- ✅ C2 DOCX writer + xUnit 라운드트립 6건 + 스모크 1건
- ✅ C2b Markdown reader 를 Markdig 로 교체 (CommonMark 풀 파싱)
- ✅ C2.5 비텍스트 객체 (표·이미지·OpaqueBlock) 1차 — Core 모델 + DOCX/IWPF 라운드트립 + WPF 시각화
- ✅ C3 HWPX reader (KS X 6101) — 자체 구현 1차 (단락·런·정렬·강조·헤더 H1~H6)
- ✅ C4 HWPX writer + 라운드트립 테스트 (xUnit 6건 + 스모크 1건)
- ✅ C4.5 한컴 hwpx 변종 호환 — ZIP entry path 정규화, BOM-aware StreamReader, OPF spine .xml 필터, header.xml 의 charPr/paraPr/style ID → PolyDoc 모델 매핑
- ✅ C5 HWPX 표·이미지 양방향 (`<hp:tbl>` ↔ Table, `<hp:pic>`+BinData ↔ ImageBlock, SHA-256 dedupe). OpaqueBlock 은 다음 사이클
- ☐ C6 HWPX writer 의 한컴 호환 향상 — header.xml 에 사용자별 RunStyle/ParagraphStyle 마다 동적 charPr/paraPr 생성, hp:linesegarray 보강
- ☐ G3 (DOCX 측): 사용자가 Word 에서 시각 검증 — 통과 (직접 처리·라운드트립 정상 보고)
- ◑ G3 (HWPX 측 reader): 한컴 hwpx 4건 본문·서식 정상 표시 — 통과 (사용자 보고 OK)
- ☐ G3 (HWPX 측 writer): PolyDoc 가 만든 .hwpx 를 한컴 오피스에서 정상 표시 — 사용자 검증 필요

### Phase D — 외부 CLI 컨버터 분리
- ☐ D1 PolyDoc.Cli.Docx 분리
- ☐ D2 PolyDoc.Cli.Hwpx 분리
- ☐ D3 메인 앱 ↔ CLI IPC (인자/표준입출력/exit code)
- ☐ G4: LibreOffice 의존 vs 자체 결정

### Phase E — 편집 기능 / M3-M4
- ☐ 표·이미지·머리말/꼬리말·각주/미주
- ☐ 변경추적·주석
- ☐ 수식·도형/텍스트박스
- ☐ 필드/목차

### Phase F — DOC/HWP ingest / M5
- ☐ DOC import (LibreOffice 또는 자체)
- ☐ HWP import (LibreOffice 또는 자체)
- ☐ Opaque island 정책 적용

### Phase G — 고급 기능 / M6-M7
- ☐ 테마 다수, 사인 만들기 독립 앱, 사전·맞춤법 외부 모듈

### Phase H — 인스톨러 / `1.0.0` 릴리즈
- ☐ MSIX 인스톨러
- ☐ G5: 사용자 명시 지시 후 `1.0.0` 컷

---

## 사용자 디버깅·컨펌 게이트

| 게이트 | 시점 | 사용자 작업 |
|---|---|---|
| G0 | 시작 전 | 기술 스택 승인 ✅ (완료) |
| G1 | Phase A 종료 | PR 리뷰 후 머지 결정 |
| G2 | Phase B 시작 | Windows에서 `dotnet build` / `dotnet run` 후 결과 보고 |
| G3 | Phase C 종료 | HWPX·DOCX 결과를 한컴/Word에서 시각 검증 |
| G4 | Phase D 진입 시 | LibreOffice 의존 노선 확정 vs 자체 구현 |
| G5 | Phase H 진입 시 | "릴리즈하자" 명시 — `1.0.0` 컷 |

---

## 다음 세션 인수인계 체크리스트

세션 종료 시점에 이 문서의 **Phase 진행표** 와 아래 항목을 갱신한다:

- [ ] HISTORY.md `[Unreleased]` 정리
- [ ] 미해결 이슈 / 알려진 버그 목록
- [ ] 다음 세션 첫 작업 후보 (구체적 파일/함수명)
- [ ] 사용자 답을 기다리는 게이트 (있다면 어떤 게이트인지)

---

## 현재 인수인계 (Phase C 진입 사이클 종료 시점)

### 완료
- Phase A: src 4 + tests 4 + tools/SmokeTest. **G1 통과 확인** (사용자 보고: build/test/smoke 모두 OK).
- Phase B 첫 사이클: PolyDoc.App WPF (메인 윈도우, 메뉴 6단, About, Light 테마). 핸텍 로고/아이콘 통합. **G2 통과 확인** (사용자 보고: build/run/UI 모두 OK).
- Phase C C1·C2: DOCX reader/writer + 라운드트립 테스트 6건 + 스모크. DOCX 가 외부 컨버터 위탁에서 직접 처리 대상으로 승격.
- Phase C C2b: Markdown reader 를 Markdig 0.42.0 로 교체. CommonMark 풀 파싱 + 추가 테스트 5건.

### 현재 테스트 현황 (Linux 환경)
| 프로젝트 | 테스트 수 | 상태 |
|---|---|---|
| PolyDoc.Core.Tests | 9 | ✅ |
| PolyDoc.Iwpf.Tests | 9 | ✅ |
| PolyDoc.Codecs.Text.Tests | 5 | ✅ |
| PolyDoc.Codecs.Markdown.Tests | 11 | ✅ |
| PolyDoc.Codecs.Docx.Tests | 9 | ✅ |
| PolyDoc.Codecs.Hwpx.Tests | 9 | ✅ |
| **합계** | **52** | **All green** |
| PolyDoc.SmokeTest 콘솔 | 6 | ✅ |

### 사용자(노진문) 작업이 필요한 항목 — RichTextBox 업그레이드 검증 (G2.5)
- [ ] Windows 에서 `git pull` 후 `dotnet restore`
- [ ] `dotnet build PolyDoc.sln` — App + 새 PolyDoc.App.Tests 까지 포함해 0 error
- [ ] `dotnet test PolyDoc.sln` — 기존 36건 + 신규 9건 = **xUnit 45건 모두 그린**
- [ ] `dotnet run --project src/PolyDoc.App`
- [ ] **시각 검증**: Word 에서 만든 `.docx` (제목·본문·굵게·색·정렬 섞인) 를 열어 → 화면에 서식이 그대로 보여야 함
- [ ] **편집 후 저장**: 본문 일부 수정 → 저장 → 다시 Word 에서 열어 서식 보존 확인 (G3 일부)
- [ ] `.iwpf` 라운드트립도 동일하게 시각 보존되는지

### 다음 단계 후보
- **개요 서식 — 색상 입력 위젯 통합** (매뉴얼·문서 최종 수정 전 처리). 현재 `OutlineStyleWindow` 의 선색·배경색은 hex 텍스트박스 + 클릭 시 ColorDialog 를 여는 스와치로 분리되어 있다. 이를 "텍스트 입력 + 드롭다운 picker" 가 결합된 단일 위젯(예: 표준 색 팔레트 + 사용자 정의 클릭 시 ColorDialog 오픈) 으로 교체 — 입력·선택을 한 컨트롤에서 모두 처리.
- **표·이미지 opaque 보존** — IWPF.md 의 「opaque island」 정책 적용. DocxReader 가 표/이미지를 OpaqueBlock 으로 보존, Writer 가 그대로 재출력. RichTextBox 본문엔 placeholder 토큰.
- **C3·C4 HWPX codec** — KS X 6101 기반 자체 구현.
- **B 사이클 폴리싱 더** — i18n .resx (한/영), 테마 다중화, 드래그&드롭, 찾기·바꾸기.
- **D 단계 진입** — HWP/DOC/HTML 외부 컨버터 (LibreOffice headless) IPC 연결.

### 다음 사이클 (G2 통과 후)
1. **i18n 분리** — `Properties/Resources.resx` (ko-KR 기본) + `Resources.en.resx`. XAML 에 `{x:Static p:Resources.MenuFile}` 바인딩.
2. **테마 다중화** — 학생/청년/장년 대상 (Soft, Vivid, HighContrast 등) — `Themes/<Name>.xaml` 추가, 도구 → 설정에서 런타임 전환.
3. **편집 기능 1차** — 찾기/바꾸기 다이얼로그, 문서 정보 (속성) 다이얼로그.
4. **드래그 & 드롭** — 파일을 윈도우에 끌어 놓으면 즉시 열기.

### 알려진 한계 (코드 베이스 전반)
- Block 다형성을 `JsonDerivedType` 로 처리 — 현재 `Paragraph` 만 등록. `Table`, `Image`, `TextBox` 등 추가 시 같은 위치에 등록 (`src/PolyDoc.Core/Block.cs`).
- IWPF document.json 본문은 Phase A 에서 JSON. 후속 단계에서 IWPF 사양에 맞춰 일부를 XML 로 전환 가능 (`document.xml`, `styles.xml`).
- 본문 편집기가 RichTextBox + FlowDocument. FlowDocument 가 표현 못 하는 모델 속성(장평·자간·Provenance)은 ViewModel 의 `_document` 머지 베이스로 비파괴 보존. 편집 후에도 유지되는지는 Save 시 `FlowDocumentParser.Parse(fd, originalForMerge: _document)` 호출이 정상 동작에 의존.
- WPF 빌드 검증을 본 환경(Linux)에서 못 함. App 코드 변경 직후엔 (1) csproj `<ProjectReference>` 누락 점검, (2) 새 NuGet 패키지의 namespace 와 우리 type 이름이 충돌하는지 점검 — 두 가지를 매번 확인 사이클로 돌릴 것 (DocumentFormat.OpenXml 충돌로 두 차례 빌드 실패한 lessons-learned).

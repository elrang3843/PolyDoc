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
| **NuGet.org** | **503 으로 차단** — `api.nuget.org` 응답 안 함 (curl/dotnet restore 모두). BCL 만 사용하는 src 프로젝트는 복원/빌드 정상. xUnit/Markdig 같은 외부 패키지는 본 환경에서 복원 불가 |

**핵심 제약**:
- WPF 앱(`Microsoft.NET.Sdk.WindowsDesktop`)은 Windows 에서만 빌드 가능.
- 이 환경에서는 **순수 라이브러리·코덱** 만 검증 가능. UI 검증은 사용자 책임.
- xUnit 테스트는 **코드는 작성하되 본 환경에서 실행 불가**. Windows(또는 NuGet 가능 환경)에서 `dotnet restore && dotnet test` 로 검증한다.
- Markdig 같은 외부 의존이 필요한 코덱은 NuGet 회복 시점에 도입한다 (현재는 BCL 서브셋 구현으로 대체).

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
- ✅ A1 솔루션 스캐폴딩 (PolyDoc.slnx, .NET 10, CPM, 4 src + 4 tests + tools/SmokeTest)
- ✅ A2 PolyDoc.Core (공통 문서 모델)
- ✅ A3 PolyDoc.Iwpf (reader/writer, ZIP+JSON, SHA-256 검증, 위변조 거부)
- ✅ A4 PolyDoc.Codecs.Text (TXT in/out, BOM 감지)
- ✅ A5 PolyDoc.Codecs.Markdown (Markdig 없이 BCL 서브셋: 헤더·리스트·강조)
- ◑ A6 단위 테스트 — xUnit 코드 작성 완료, **NuGet 차단으로 본 환경 미실행**. Windows 에서 G2 시 `dotnet test` 로 검증.
- ✅ A6b 콘솔 스모크 러너 4/4 통과 (라운드트립 3종 + 위변조 검출)
- ✅ A7 커밋·푸시

### Phase B — WPF UI 셸 (Windows 필수)
- ✅ B1 PolyDoc.App 스캐폴딩 (`net10.0-windows` + WPF + CommunityToolkit.Mvvm 8.4.0, ApplicationIcon=Handtech.ico, Handtech_1024.png 임베드)
- ✅ B2 메인 메뉴 6단(파일/편집/입력/서식/도구/도움말) + TextBox 본문 편집기 + 상태 바 + About 다이얼로그
- ◑ B3 i18n 한/영 — 1차 사이클은 한국어 하드코딩, `.resx` 리소스 분리는 다음 사이클로 이연
- ◑ B4 테마 시스템 — 1차 사이클은 Light 단일 테마(핸텍 브랜드 블루), 다중 테마는 다음 사이클
- ☐ B5 **G2** — Windows 머신에서 첫 `dotnet build` / `dotnet run` 결과 보고. 메뉴 동작·About 표시·IWPF/MD/TXT 열고 저장 검증

### Phase C — DOCX/HWPX 1급 시민 (M2-M3)
- ☐ C1 DOCX reader (OpenXml)
- ☐ C2 DOCX writer + 라운드트립 테스트
- ☐ C3 HWPX reader (KS X 6101)
- ☐ C4 HWPX writer + 라운드트립 테스트
- ☐ G3: 사용자가 Word/한컴에서 결과 시각 검증

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

## 현재 인수인계 (Phase B 첫 사이클 종료 시점)

### 완료
- Phase A: src 4 + tests 4 + tools/SmokeTest. G1 통과.
- Phase B 첫 사이클: PolyDoc.App WPF (메인 윈도우, 메뉴 6단, About, Light 테마). 핸텍 로고/아이콘 통합.

### 사용자(노진문) 작업이 필요한 항목 — G2 직전
- [ ] Windows 머신에서 최신 브랜치 pull
- [ ] `dotnet restore PolyDoc.slnx` — CommunityToolkit.Mvvm 8.4.0 자동 복원
- [ ] `dotnet build PolyDoc.slnx` — `PolyDoc.App` 까지 포함해 0 error
- [ ] `dotnet run --project src/PolyDoc.App` — 메인 윈도우가 뜨는지
- [ ] **메뉴 검증**: 파일 → 새 파일 / 불러오기(IWPF·MD·TXT) / 저장 / 다른 이름으로 저장 동작
- [ ] **단축키 검증**: Ctrl+N / Ctrl+O / Ctrl+S / Ctrl+Shift+S
- [ ] **About 검증**: 도움말 → PolyDoc 정보 — 핸텍 로고·노진문·버전 1.0.0-test.1 표시
- [ ] **외부 포맷 보호**: HWP/HWPX/DOC/DOCX 열기 시도 시 "외부 컨버터 필요" 안내가 뜨는지
- [ ] 작업창 타이틀 바: 편집 후 `*` 표시, 저장 후 사라지는지
- [ ] (가능하다면) 메뉴 동작·About 다이얼로그 스크린샷 첨부

### 알려진 위험 (Linux 에서 빌드 검증 불가)
- WPF 빌드는 Windows 전용 SDK 가 필요하므로 본 환경에서 컴파일 검증 못 함. XAML 의 binding path, namespace, pack URI 가 첫 시도에 정확해야 함. **빌드 에러가 나면 출력 그대로 보내주면 다음 응답에서 즉시 패치**.
- `pack://application:,,,/Assets/Handtech_1024.png` 는 csproj 의 `<Resource Link="Assets\Handtech_1024.png">` 와 일치시켰지만, MSBuild 가 Link 메타를 정확히 처리하는지는 Windows 빌드에서 첫 검증.
- Directory.Build.props 의 `TreatWarningsAsErrors=true` 가 WPF 코드젠 경고와 충돌하면 빌드 실패 가능. 발생 시 App 프로젝트 한정으로 완화 검토.

### 다음 사이클 (G2 통과 후)
1. **i18n 분리** — `Properties/Resources.resx` (ko-KR 기본) + `Resources.en.resx`. XAML 에 `{x:Static p:Resources.MenuFile}` 바인딩.
2. **테마 다중화** — 학생/청년/장년 대상 (Soft, Vivid, HighContrast 등) — `Themes/<Name>.xaml` 추가, 도구 → 설정에서 런타임 전환.
3. **편집 기능 1차** — 찾기/바꾸기 다이얼로그, 문서 정보 (속성) 다이얼로그.
4. **드래그 & 드롭** — 파일을 윈도우에 끌어 놓으면 즉시 열기.

### 알려진 한계 (코드 베이스 전반)
- Markdown 코덱이 Markdig 의 풀 CommonMark 가 아닌 **실용 서브셋**. 코드블록·인용·표·이미지·링크 미지원. NuGet 회복 후 Markdig 로 교체.
- Block 다형성을 `JsonDerivedType` 로 처리 — 현재 `Paragraph` 만 등록. `Table`, `Image`, `TextBox` 등 추가 시 같은 위치에 등록 (`src/PolyDoc.Core/Block.cs`).
- IWPF document.json 본문은 Phase A 에서 JSON. 후속 단계에서 IWPF 사양에 맞춰 일부를 XML 로 전환 가능 (`document.xml`, `styles.xml`).
- 본문 편집기가 `TextBox` (plain string). Phase E 에서 RichTextBox/FlowDocument 로 교체하면서 ParagraphStyle/RunStyle 양방향 동기화 도입.

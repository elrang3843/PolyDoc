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

**핵심 제약**: WPF 앱(`Microsoft.NET.Sdk.WindowsDesktop`)은 Windows에서만 빌드 가능.
이 환경에서는 **순수 라이브러리·코덱·테스트** 만 검증 가능. UI 검증은 사용자 책임.

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
- ◑ A1 솔루션 스캐폴딩
- ☐ A2 PolyDoc.Core (공통 문서 모델)
- ☐ A3 PolyDoc.Iwpf (reader/writer + manifest/document 직렬화)
- ☐ A4 PolyDoc.Codecs.Text
- ☐ A5 PolyDoc.Codecs.Markdown
- ☐ A6 단위 테스트 그린
- ☐ A7 첫 커밋·푸시

### Phase B — WPF UI 셸 (Windows 필수)
- ☐ B1 PolyDoc.App 스캐폴딩 (WPF + MVVM)
- ☐ B2 메뉴/툴바/문서 탭/룰러 골격
- ☐ B3 i18n 리소스 (한/영)
- ☐ B4 테마 시스템 (학생~장년 대상 다중 테마)
- ☐ B5 사용자 게이트 G2: Windows에서 첫 빌드/run 결과 보고

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

<p align="center">
  <img src="assets/Handtech_1024.png" alt="HANDTECH" width="160" height="160" />
</p>

<h1 align="center">PolyDoc</h1>

<p align="center">
  <b>HWP · HWPX · DOC · DOCX · HTML/HTM · MD · TXT</b> 문서를 한 곳에서 읽고 편집하고<br/>
  자체 무손실 포맷 <b>IWPF</b> 로 보관하는 데스크톱 워드프로세서.
</p>

<p align="center">
  <a href="LICENSE"><img src="https://img.shields.io/badge/License-Apache%202.0-blue.svg" alt="License: Apache 2.0"/></a>
  <a href="#시스템-요구사항"><img src="https://img.shields.io/badge/Platform-Windows%2010%2B-0078D6.svg" alt="Platform: Windows 10+"/></a>
  <a href="#프로젝트-상태"><img src="https://img.shields.io/badge/Status-Pre--alpha-orange.svg" alt="Status: Pre-alpha"/></a>
  <a href="#소스에서-빌드하기"><img src="https://img.shields.io/badge/Lang-C%23%20%2F%20WPF-512BD4.svg" alt="Language: C# / WPF"/></a>
  <a href="#소스에서-빌드하기"><img src="https://img.shields.io/badge/.NET-10.0-512BD4.svg" alt=".NET 10"/></a>
</p>

<p align="center">
  Made by <b>핸텍 (HANDTECH)</b> · 저작권자 <b>노진문 (Noh JinMoon)</b>
</p>

---

## 목차
- [PolyDoc이 해결하는 문제](#polydoc이-해결하는-문제)
- [주요 특징](#주요-특징)
- [지원 포맷](#지원-포맷)
- [프로젝트 상태](#프로젝트-상태)
- [시스템 요구사항](#시스템-요구사항)
- [설치](#설치)
- [빠른 시작](#빠른-시작)
- [메뉴 구성](#메뉴-구성)
- [소스에서 빌드하기](#소스에서-빌드하기)
- [아키텍처 개요](#아키텍처-개요)
- [로드맵](#로드맵)
- [다국어 지원](#다국어-지원)
- [기여하기](#기여하기)
- [버그 리포트 / 기능 요청](#버그-리포트--기능-요청)
- [라이선스](#라이선스)
- [참고 문서](#참고-문서)

---

## PolyDoc이 해결하는 문제

업무 문서는 흔히 **HWP / DOCX / DOC / HWPX** 가 뒤섞여 유통됩니다. 한 포맷에서
다른 포맷으로 변환할 때마다 표·머리말·번호·한글 조판 같은 미세 정보가 깨지고,
원본을 다시 받아야 하는 일이 반복됩니다.

PolyDoc은 모든 문서를 **공통 의미 모델 + 포맷별 보존 캡슐 + 원본 내장**으로
구성된 자체 포맷 **IWPF**로 정규화해 보관합니다. 그 결과:

- **편집·검색**은 IWPF의 공통 모델로 빠르게,
- **외부 포맷으로의 라운드트립(HWPX ↔ IWPF ↔ DOCX 등)** 은 원본에 가깝게,
- **원본 무손실 보존**은 패키지 안에 포함된 원본 파일로 보장합니다.

자세한 설계 근거는 [`IWPF.md`](IWPF.md) 참고.

---

## 주요 특징

- **하나의 에디터에서 다중 포맷 읽기/쓰기** — HWP, HWPX, DOC, DOCX, HTML/HTM, MD, TXT
- **IWPF — 자체 통합 포맷** — ZIP 기반 패키지에 공통 모델 + 충실도 캡슐 + 원본 + provenance map 동봉
- **원본 무손실 보장** — 가져온 원본 파일을 패키지에 그대로 보관해 byte-level 원복 가능
- **편리한 편집 기능** — 표(엑셀풍 시트), 다양한 글상자(말풍선·구름풍선·가시풍선·번개상자), 그래프, 도형, 수식, 이모지
- **다양한 테마** — 학생부터 장년까지 폭넓은 사용자층을 고려한 컬러 테마 선택
- **다국어 UI** — 한국어(기본), 영어
- **사인 만들기** — 별도 앱으로 서명 이미지 생성 → 문서 내 도형/그림으로 삽입
- **외부 변환은 분리된 CLI 모듈** — 메인 앱은 IWPF/MD/TXT 만 직접 처리하고, 그 외 포맷은 별도 컨버터로 호출

---

## 지원 포맷

| 포맷  | 읽기 | 쓰기 | 비고                                         |
|-------|:---:|:---:|----------------------------------------------|
| IWPF  | ✅  | ✅  | 자체 정본(canonical) 포맷                     |
| MD    | ✅  | ✅  | 기본 내장                                     |
| TXT   | ✅  | ✅  | 기본 내장                                     |
| DOCX  | ✅  | ✅  | 1단계 1급 시민. **메인 앱 직접 처리** (DocumentFormat.OpenXml) |
| HWPX  | ✅  | ✅  | 1단계 1급 시민. **자체 구현 (KS X 6101)** — 한컴 오피스 호환성 fine-tune 진행 중 |
| HTML / HTM | ⏳ | ⏳ | 외부 컨버터 모듈 (Phase D)                |
| DOC   | ⏳  | ⚠️  | 2단계: ingest 전용. 출력은 HWPX/DOCX 권장     |
| HWP   | ⏳  | ⚠️  | 2단계: ingest 전용. 출력은 HWPX/DOCX 권장     |

> 다른 포맷으로 저장할 때는 **항상 한 번 더 확인 다이얼로그**가 뜹니다.
> 외부 포맷에서는 일부 정보가 손실될 수 있고, 정본 보존을 위해 IWPF 저장을 권장합니다.

---

## 프로젝트 상태

> 🚧 **Pre-alpha — 설계/초기 구현 단계입니다.**
> 현재 저장소에는 사양 문서(`README.md`, `IWPF.md`, `CLAUDE.md`, `HISTORY.md`)만 포함되어 있고,
> 실행 가능한 바이너리·릴리스는 아직 제공되지 않습니다.
> 진행 상황은 [Issues](../../issues) / [Releases](../../releases) / [`HISTORY.md`](HISTORY.md) 에서 확인하세요.

### 버전 정책

- **`1.0.0` 이전의 모든 빌드는 테스트 버전입니다.** 태그 형식: `1.0.0-test.<n>` (예: `1.0.0-test.1`).
- **최초 정식 릴리스는 `1.0.0`** 이며, 메인테이너의 명시적 릴리스 결정이 있을 때만 컷합니다.
- `1.0.0` 이후에는 일반 [Semantic Versioning](https://semver.org/lang/ko/) 규칙(`1.0.1` / `1.1.0` / `2.0.0` ...)을 따릅니다.

### 개발 단계

1. **1단계 (현재 목표)** — DOCX, HWPX 1급 지원 / IWPF 저장·로드 / 기본 편집
2. **2단계** — DOC, HWP **ingest 전용** 추가, 정규화 후 출력은 HWPX/DOCX
3. **3단계** — 변경추적, 주석, 수식, 도형/텍스트박스, 필드/목차, 고급 표, 특수 조판

---

## 시스템 요구사항

| 항목       | 요구 사항                              |
|-----------|---------------------------------------|
| OS        | Windows 10 (1809) 이상, Windows 11    |
| 아키텍처   | x64                                    |
| 런타임     | .NET (런타임 포함 인스톨러 제공 예정)   |
| 디스크     | 약 200 MB 이상 권장                    |
| 메모리     | 4 GB 이상 권장                         |

> macOS / Linux 는 현재 지원하지 않습니다.

---

## 설치

### 정식 릴리스 (예정 — `v1.0.0`)

정식 릴리스(`v1.0.0`)가 준비되면 [Releases 페이지](../../releases)에서 다운로드할 수 있습니다.
설치 방법은 다음 두 가지가 제공될 예정입니다.

- **MSIX 설치 패키지** — Windows 표준 설치 마법사로 설치
- **Portable ZIP** — 압축 해제 후 `PolyDoc.exe` 실행

### 테스트 빌드 (`1.0.0-test.<n>`)

`v1.0.0` 이전의 모든 공개 빌드는 **테스트 버전** 입니다.
배포되는 경우 [Releases 페이지](../../releases)에 `Pre-release` 로 표시되며,
실험·평가 목적 외의 운영 사용은 권장되지 않습니다.

### 현재 시점에서는

실행 가능한 빌드가 아직 없습니다. 소스에서 직접 빌드해 사용해야 합니다 → [소스에서 빌드하기](#소스에서-빌드하기) 참고.

---

## 빠른 시작

```text
1. PolyDoc 실행
2. [파일] → [새 파일]   ─ 기본 IWPF 모드로 새 문서 생성
   또는
   [파일] → [불러오기]  ─ HWP/HWPX/DOC/DOCX/HTML/MD/TXT 등 불러오기
3. 본문을 편집
4. [파일] → [저장]              ─ IWPF / MD / TXT 로 저장
   [파일] → [다른 이름으로 저장] ─ HWPX, DOCX 등 다른 형식으로 내보내기
```

> **팁:** 다른 형식으로 내보낸 뒤에도 정본은 IWPF로 함께 저장해 두는 것을 권장합니다.
> 외부 앱에서 편집·저장하면 PolyDoc 전용 보존 정보가 손실될 수 있습니다.

---

## 메뉴 구성

<details>
<summary><b>파일</b></summary>

- 새 파일 (기본 IWPF)
- 불러오기 / 저장 / 다른 이름으로 저장 — IWPF·MD·TXT 는 내장, 그 외 포맷은 외부 컨버터 호출
- 미리보기 (편집용지·인쇄 색상 설정 포함)
- 인쇄
- 닫기 (저장 여부 확인)
- 종료
</details>

<details>
<summary><b>편집</b></summary>

- 복사 / 잘라내기 / 지우기 / 붙여넣기
- 가져오기 — 외부 문서·개체를 현재 위치에 삽입
- 내보내기 — 선택 영역을 외부 문서·개체로 저장
- 문서정보 — 작성/편집 정보, 암호, 워터마크
- 찾기 / 바꾸기
</details>

<details>
<summary><b>입력</b></summary>

- **글상자** — 사각형, 말풍선, 구름풍선(머릿속 생각), 가시풍선(번뜩이는 아이디어), 번개상자(임팩트 표현)
- **표(시트)** — 엑셀 같은 시트형 표
- **그래프** — 꺾은선·파이·막대·분포 등
- 특수문자 / 수식 / 이모지
- **도형** — 직선, 폴리곤·스플라인 선/면, 사각형, 삼각형, 원, 타원, 호, 화살표 등
- 그림 — PNG, BMP, JPEG, TIFF
- 사인 (불러오기)
</details>

<details>
<summary><b>서식</b></summary>

- **글자 서식** — 폰트, 크기, 글자폭(%), 자간(px), 두껍게/이탤릭/위첨자/아래첨자/밑줄/중간줄/윗줄/테두리, 글자색·배경색·줄색
- **문단 서식** — 줄 간격, 문단 간격, 들여쓰기/내어쓰기, 자동번호(Markdown 호환)
- **페이지 서식** — 편집 용지, 색상(단색·16색·256색·Full Color), 여백, 머릿글/꼬릿글, 다단
</details>

<details>
<summary><b>도구</b></summary>

- 설정 — 사용자 정보, 룰러/눈금/편집용지 표시, 읽기·쓰기 포맷 활성화 토글, 언어, 테마
- 사전 — 외부 앱/모듈 연결 (우선순위 낮음)
- 맞춤법 검사 — 외부 모듈 연결 (우선순위 낮음)
- 사인 만들기 — 별개의 독립 앱으로 실행, [입력] 메뉴의 사인과 연계
</details>

<details>
<summary><b>도움말</b></summary>

- 사용 방법(매뉴얼) — 언어별 별도 파일
- 라이선스 — 써드파티 포함, 언어별 별도 파일
- About — 저작권자·회사 로고 포함
</details>

---

## 소스에서 빌드하기

> ⚠️ 빌드 스크립트와 솔루션 구조는 1단계 진행 중 확정됩니다. 아래는 예정된 절차입니다.

### 사전 요구

- **Windows 10/11**
- **Visual Studio 2022** (워크로드: *.NET 데스크톱 개발*, 필요 시 *Windows App SDK*)
- **.NET SDK** (버전은 솔루션 확정 시 명시)
- **Git**

### 클론 & 빌드

```bash
git clone https://github.com/elrang3843/PolyDoc.git
cd PolyDoc

# (예정) 솔루션 빌드
# dotnet build PolyDoc.sln -c Release

# (예정) 실행
# dotnet run --project src/PolyDoc.App
```

빌드가 가능한 시점부터 본 문서가 갱신됩니다.

---

## 아키텍처 개요

PolyDoc은 **2계층 설계** 위에 동작합니다.

1. **의미 계층 (공통 문서 모델)** — 검색·편집·분석을 담당하는 superset 모델
2. **충실도 계층 (Fidelity Capsule + 원본 내장 + Provenance Map)** — 원본 재현·역변환 보장

```
PolyDoc.exe (WPF / WinUI 3)
   ├─ Editor / Renderer  ── 공통 의미 모델 위에서 동작
   ├─ IWPF Reader/Writer ── ZIP 패키지 입출력
   └─ External Converters (CLI)
        ├─ HWP   <-> IWPF
        ├─ HWPX  <-> IWPF
        ├─ DOC   ->  IWPF      (ingest 전용)
        ├─ DOCX  <-> IWPF
        └─ HTML  <-> IWPF
```

IWPF 패키지 구조, 보존 캡슐 설계, provenance / dirty tracking, opaque island
정책 등 자세한 내용은 [`IWPF.md`](IWPF.md) 와 개발 가이드인
[`CLAUDE.md`](CLAUDE.md) 를 참고하세요.

---

## 로드맵

- [ ] **M1** — 솔루션 골격, IWPF reader/writer 프로토타입, 기본 텍스트 편집
- [ ] **M2** — DOCX import/export, HWPX import/export
- [ ] **M3** — 표, 이미지, 머리말/꼬리말, 각주/미주
- [ ] **M4** — 변경추적, 주석, 수식, 도형/텍스트박스
- [ ] **M5** — DOC / HWP ingest, opaque island 정책 적용
- [ ] **M6** — 테마 다중화, i18n(한/영) 완성, 인쇄/미리보기
- [ ] **M7** — 사인 만들기 독립 앱, 사전·맞춤법 외부 모듈 연동
- [ ] **M8** — MSIX 인스톨러, 첫 정식 릴리스 **`v1.0.0`** (메인테이너 명시 지시 시 컷)

진행 상황은 [Projects](../../projects) / [Milestones](../../milestones) 에서 추적합니다.

---

## 다국어 지원

- 기본 언어: **한국어**
- 1단계 지원: **영어**
- UI 문자열, 매뉴얼, 라이선스 표기 모두 **언어별 분리 파일** 로 관리됩니다.

추가 언어 기여는 [기여 가이드](#기여하기)를 통해 환영합니다.

---

## 기여하기

PolyDoc은 초기 단계라 이슈 제보와 설계 토론 모두 큰 도움이 됩니다.

1. **Issue 먼저 열기** — 큰 변경(아키텍처·포맷·UX)은 PR 전에 Issue로 합의를 잡습니다.
2. **브랜치 분기** — `feature/<주제>`, `fix/<주제>` 형식 권장.
3. **커밋 메시지** — 어떤 변화인지 한 줄 요약 + 필요 시 본문에 *왜* 를 적습니다.
4. **PR** — 변경 의도, 영향 범위, 테스트 방법을 본문에 명시합니다.

> 코드 스타일·테스트·CI 가이드는 솔루션 확정 후 `CONTRIBUTING.md` 로 분리될 예정입니다.

---

## 버그 리포트 / 기능 요청

- 버그: [Issues → New issue → Bug report](../../issues/new?template=bug_report.md)
- 기능 요청: [Issues → New issue → Feature request](../../issues/new?template=feature_request.md)
- 보안 취약점은 공개 이슈로 올리지 말고 메인테이너에게 비공개로 알려 주세요.

이슈를 올릴 때는 다음 정보가 도움이 됩니다.

- PolyDoc 버전 / 빌드 해시
- Windows 버전
- 재현 절차, 입력 파일(가능하면 최소 재현 샘플)
- 기대 동작 vs 실제 동작
- 스크린샷 또는 로그

---

## 만든 사람들

<table>
  <tr>
    <td align="center" width="160">
      <img src="assets/Handtech_1024.png" width="120" height="120" alt="HANDTECH"/><br/>
      <sub><b>핸텍 (HANDTECH)</b></sub>
    </td>
    <td>
      <ul>
        <li><b>회사</b>: 핸텍 (HANDTECH)</li>
        <li><b>저작권자 / 메인테이너</b>: 노진문 (Noh JinMoon)</li>
        <li><b>GitHub</b>: <a href="https://github.com/elrang3843">@elrang3843</a></li>
        <li><b>Repository</b>: <a href="https://github.com/elrang3843/PolyDoc">elrang3843/PolyDoc</a></li>
      </ul>
    </td>
  </tr>
</table>

회사·저작권 정보의 정식 표기는 [`NOTICE`](NOTICE) 파일을 따릅니다.

---

## 라이선스

이 프로젝트는 [Apache License 2.0](LICENSE) 으로 배포됩니다.

```
Copyright (c) 2026 HANDTECH (핸텍) — Noh JinMoon (노진문)
Licensed under the Apache License, Version 2.0
```

써드파티 의존성의 라이선스 고지는 추후 [`NOTICE`](NOTICE) 와
앱 내 [도움말 → 라이선스] 메뉴에서 제공됩니다.

---

## 참고 문서

| 문서                                | 대상            | 내용                                          |
|------------------------------------|----------------|-----------------------------------------------|
| [`README.md`](README.md)           | 사용자·기여자   | 프로젝트 소개와 사용 안내 (이 문서)             |
| [`HISTORY.md`](HISTORY.md)         | 사용자·기여자   | 변경 이력(Changelog) — 버전별 추가·수정 내역    |
| [`IWPF.md`](IWPF.md)               | 설계자·개발자   | IWPF 통합 포맷 사양과 설계 근거                |
| [`WORK_PLAN.md`](WORK_PLAN.md)     | 메인테이너·AI   | 다단계 작업 계획·진행 상태·인수인계             |
| [`CLAUDE.md`](CLAUDE.md)           | Claude Code    | AI 어시스턴트가 참고할 개발 가이드라인          |
| [`NOTICE`](NOTICE)                 | 모두           | 저작권 고지·써드파티 attribution                |
| [`LICENSE`](LICENSE)               | 모두           | Apache License 2.0 본문                       |

---

<p align="center">
  <sub>PolyDoc — 한국어 문서 생태계와 글로벌 워드프로세서 포맷을 한 자리에서.</sub><br/>
  <sub>© 2026 HANDTECH (핸텍) · Noh JinMoon (노진문)</sub>
</p>

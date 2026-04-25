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
- 버전 헤더 형식: `## [0.1.0] - 2026-MM-DD`
- 날짜는 `YYYY-MM-DD` (KST 기준).
- 정식 릴리스는 **첫 실행 가능 빌드(M1 완료 시점)부터** 버전을 부여하기 시작합니다.
  그 이전의 사양·문서 작업은 본 파일 하단의 `## [Pre-release]` 섹션에 날짜별로 누적 기록합니다.

---

## [Unreleased]

> 다음 릴리스에 들어갈 변경 사항을 여기에 기록합니다.

### Added
- *(아직 없음)*

### Changed
- *(아직 없음)*

### Fixed
- *(아직 없음)*

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

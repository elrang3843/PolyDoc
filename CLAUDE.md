# CLAUDE.md

Claude Code가 이 저장소에서 작업할 때 참고하는 가이드. 자세한 내용은 `README.md`(사용자 안내), `IWPF.md`(파일 포맷 사양), `HISTORY.md`(변경 이력), `WORK_PLAN.md`(다단계 작업 계획)를 본다.

## 프로젝트 개요

**PolyDoc** — HWP, HWPX, DOC, DOCX, HTML/HTM, MD, TXT 문서를 읽어 자체 포맷 **IWPF**로 저장하는 워드프로세서 앱.

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

## 작업 시 유의사항

- 코드는 아직 없다(문서만 있는 초기 상태). 새 작업은 솔루션 구조부터 합의 후 진행.
- 문서/UI 문자열은 i18n 가능하게 분리(한국어 기본, 영어 병행).
- 파일 I/O 단위 테스트는 위 "테스트 코퍼스" 카테고리를 기준으로 설계.
- 새 외부 포맷 지원 추가 시: ① 공통 모델 매핑 ② fidelity capsule 정의 ③ provenance 기록
  ④ 라운드트립 테스트 추가 — 네 단계를 한 세트로 진행.

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

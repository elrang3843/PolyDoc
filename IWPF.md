통합문서 포맷 IWPF

1. 참조 포맷 문서
1) DOC / HWP: 레거시 바이너리 포맷
  - DOC: Microsoft의 Word Binary File Format 계열
     * 참고: Microsoft Open Specifications [MS-DOC]
  - HWP: 한글의 기존 바이너리 포맷
     * 참고: Hancom의 HWP 5.x File Structure
2) DOCX / HWPX: 현대적 XML 기반 포맷
  - DOCX: OOXML / WordprocessingML
     * 참고: ISO/IEC 29500
  - HWPX: XML 기반 개방형 포맷
     * 참고: KS X 6101

2. 독자 포맷을 만든 이유:
1) DOCX :
  - 국제 표준화(ISO/IEC 29500)된 OOXML 계열이지만
  - DOC의 일부 레거시 기능(예: 오래된 OLE, VBA 매크로 계열)은 DOCX로 그대로 안 들어갈 수 있습니다.
  - HWP/HWPX의 한글 조판 특성이나 일부 개체/수식/번호 체계는 변환 시 깨질 수 있습니다.
  - 매크로는 원래 DOCM 쪽이 더 맞습니다.
  - 다음 사항을 완전히 보장 못함
     * DOC의 레거시 바이너리 동작/호환 플래그
     * HWP/HWPX의 한글 조판 특성
     * HWP 계열 고유 개체/번호 체계/문단 배치
     * 앱별 내부 상태나 미세 호환성 옵션

2) ODT:
  - ISO/IEC 26300, 표준성과 중립성을 우선하면 가장 좋은 후보
  - 특정 벤더 종속성이 덜함
  - 장기 보존/상호운용 관점에서 명분이 좋음
  - HWP/HWPX → ODT 변환은 가능해도 레이아웃 보존률은 케이스별 차이가 큼
  - 정보 부족

3) HWPX:
  - 국내/한컴 중심이면 의미 있는 선택
  - KS X 6101(HWPX 관련 표준 계열)
  - DOC/DOCX를 HWPX로 통일하는 건 가능해도, 협업 생태계 측면에서는 덜 유리
  - 다음 사항을 완전히 보장 못함
     * Word의 레이아웃/호환 옵션
     * DOC 바이너리 레거시 개체
     * OOXML 특유의 구조와 확장 을 원형 그대로 포괄하기 어려움.

4) 결론 : 포맷을 손실 없이 완전히 포함하는 단일 문서 포맷은 없다.

3. 핵심 구조:
1) 공통 문서 모델
  - 문단, 문자, 스타일, 구역, 표, 머리말/꼬리말, 각주/미주, 이미지, 도형, 필드, 주석, 변경추적 등
2) 포맷별 보존 캡슐
  - DOC 전용 정보
  - HWP 전용 정보
  - DOCX 전용 XML/관계/호환 설정
  - HWPX 전용 XML/확장 정보
  - HTM, HTML 전용 정보
  - md 전용 정보
3) 원본 파일 내장
  - DOC, HWP의 경우 가져온 원본 파일을 아예 패키지 안에 보관
4) 소스 매핑 테이블
  - 공통 모델의 어느 노드가 원본 파일의 어느 구조와 대응하는지 기록
5) 렌더링 검증 정보
  - 폰트 해시, 페이지 분할 스냅샷, 미리보기 PDF/SVG 등

4. 왜 이 방식이면 되나
1) 형식적으로 보면
  - 통합 포맷 U가 각 포맷 F에 대해 다음을 만족하면 됨.

     * import_F : F -> U
     * export_F : U -> F

  - 이상적으로는:
     * export_F(import_F(x)) = x

    - 이 조건을 만족하려면 U 안에 F를 재생성할 만큼 충분한 정보가 있어야 함.

2) 여기서 중요한 점
  - 공통 의미 모델만으로는 부족.
     왜냐하면 각 포맷에는 서로 대응되지 않는 고유 기능이 있기 때문.

  - U는 단순 AST가 아니라:
     공통 의미 모델 + 포맷별 잔여 정보(residuals) + 필요시 원본 바이트
     를 함께 가져야 함.

5. 제안하는 통합 포맷 설계
ZIP 같은 압축 패키지 안에
패키징 계층 - 공통 문서 모델 - 포맷별 보존 캡슐 - 원본 파일 - 소스 매핑

문서 포맷은 겉보기엔 비슷해도 실제로는 차이가 큼. 
그래서 내부 포맷은 반드시 공통 문서 모델, 포맷별 보존 캡슐 두 층이어야 함.
1) 패키징 계층
    ZIP/OPC 유사 컨테이너가 가장 현실적입니다.
    ZIP + JSON(or XML) + binary resources
  - 이유:
     * 파일 묶기 쉬움
     * XML/JSON/바이너리 혼합 저장 가능
     * 원본 첨부, 썸네일, 렌더링 캐시, 서명 정보까지 담기 좋음
     * 사람이 디버깅 가능
     * 버전 관리 쉬움
     * 리소스 분리 저장 가능
     * exporter 작성 쉬움

  - 예시 구조:
     iwpf/
	manifest.json
	content/document.xml
	content/styles.xml
	content/numbering.xml
	content/sections.xml
	content/annotations.xml
	resources/images/*
	resources/ole/*
	resources/fonts/*
	fidelity/original/source.doc
	fidelity/original/source.hwp
	fidelity/capsules/msdoc/*
	fidelity/capsules/hancom/*
	provenance/source-map.json
	render/layout-snapshot.json
  	render/preview.pdf
  	signatures/*

2) 공통 문서 모델
     공통 모델은 “최소공배수”가 아니라 “최대상한(superset)” 이어야 함.

  - 반드시 들어가야 할 것
     * 문서 메타데이터
     * 섹션 / 페이지 설정
     * 문단 / 문자 런(run)
     * 스타일 계층
     * 번호/개요/목차 구조
     * 표 / 병합 / 셀 속성
     * 이미지 / 도형 / 텍스트박스
     * 머리말 / 꼬리말
     * 각주 / 미주
     * 책갈피 / 하이퍼링크 / 필드
     * 주석 / 변경추적 / 작성자 정보
     * 수식
     * 임베디드 개체(OLE 등)
     * 양식/컨트롤
     * 숨김 텍스트, 보호 설정
     * 호환성 옵션
     * 활성 콘텐츠(매크로/스크립트)는 분리 격리 저장

  - 특히 HWP/DOC 공통 통합에서 중요한 것
     * 동아시아/한글 조판 옵션
     * 줄격자 / 문자단위 배치 / 장평 / 자간 / 정렬의 미세 규칙
     * 리스트/개요 번호의 생성 규칙
     * 탭/들여쓰기/문단 줄바꿈의 세부 동작
     * 개체 anchoring/배치
     * 폰트 대체 규칙

  - 이 부분을 빼면 “문서 내용은 같지만 레이아웃이 달라지는” 손실이 발생함.

3) 포맷별 보존 캡슐
     이게 핵심
     공통 모델로 표현되지 않는 것들을 버리지 말고 캡슐로 붙잡아 둬야 함.

  - DOC 캡슐
     * 레거시 바이너리 속성
     * Word 호환성 플래그
     * OLE/ActiveX/VBA 관련 정보
     * 원본 구조상의 바이너리 조각
     * Word 고유 동작을 좌우하는 설정

  - DOCX 캡슐
     * 원본 XML 파트
     * Relationship 정보
     * Markup Compatibility 정보
     * AlternateContent
     * 커스텀 XML 파트
     * 테마/설정/호환성 파트

  - HWP 캡슐
     * 원본 바이너리 레코드
     * 한글 전용 제어문/개체
     * 문서 정보 스트림
     * 레코드 수준 속성 중 공통 모델로 못 올린 항목

  - HWPX 캡슐
     * 원본 XML 요소/속성
     * 확장 namespace 요소
     * HWPX 고유 메타/호환 정보

  - 노드 단위 연결
     공통 모델의 노드마다 이런 식으로 연결.
     {
       "nodeId": "p-184",
       "type": "paragraph",
       "core": { "...": "..." },
       "extensions": [
         { "ns": "urn:vendor:msdoc", "ref": "capsules/msdoc/p-184.bin" },
         { "ns": "urn:vendor:hancom", "ref": "capsules/hancom/p-184.bin" }
       ],
       "sourceAnchors": [
         { "format": "doc", "path": "WordDocument", "offset": 102944, "length": 88 },
         { "format": "hwp", "path": "BodyText/Section0", "recordId": 391 }
       ]
     }

4) 원본 파일 내장
     진짜 무손실을 보장하려면 원본을 보관해야 함.

  - 단순 백업이 아니라, 다음 용도로 중요:
     * byte-level 원복
     * 캡슐 설계가 미흡한 경우에도 복구 가능
     * 디지털 포렌식/감사
     * 변환기 버전 업 시 재해석 가능
  즉:
    공통 모델은 “편집/분석”용이고, 원본은 “최종 무손실 보증”용

5) 소스 매핑(provenance map)
     이것도 매우 중요

     공통 모델만 있으면 “어디서 왔는지”를 잃어버림.
       -> 원포맷으로 되돌릴 때 세밀한 보존이 어려워짐.

  - 그래서 각 노드에 대해:
     * 어느 원본 포맷에서 왔는지
     * 원본 파일의 어느 파트/스트림/레코드/요소인지
     * 아직 수정되지 않았는지
     * 수정되어 재생성이 필요한지
  를 추적해야 함.

  - 이 매핑이 있으면:
     * 수정되지 않은 부분은 원본 조각을 재사용
     * 수정된 부분만 새로 직렬화하는 하이브리드 export가 가능해짐.

6. 이 설계로 어떤 수준까지 가능한가
1) 원본 포맷으로 되돌리기 : 가능
  - 예:
	DOC → UFWP → DOC
	HWP → UFWP → HWP
	DOCX → UFWP → DOCX
	HWPX → UFWP → HWPX
	이건 가장 강하게 보장 가능.
	특히 원본 내장 + 보존 캡슐을 쓰면 됨.

2) 서로 다른 포맷으로 내보내기
  - 예:
	HWP → UFWP → DOCX
	DOC → UFWP → HWPX
	“실용적으로 매우 높게”는 가능
	하지만 순수 타겟 포맷만으로 100% 무손실은 일반적으로 어려움.
     * 이유 :
	a. HWP에만 있는 속성이 DOCX에 네이티브로 없을 수 있음
	b. DOC에만 있는 레거시 동작이 HWPX에 네이티브로 없을 수 있음

     * 생태계 내부 무손실은 가능:
	예를 들어 HWP의 고유 속성을 DOCX로 내보낼 때 DOCX 본문은 최대한 네이티브로 만들고, 
	추가로 앱 전용 보존 캡슐을 숨겨 넣으면, 
	나중에 다시 UFWP로 재수입할 때 HWP 고유 정보를 복원할 수 있음.
     즉:
	a. Word/Hanword 같은 일반 앱 기준의 무손실: 보장 어려움
        b. 내가 만든 시스템 기준의 무손실: 충분히 설계 가능

7. 중요한 한계
1) “외부 앱이 중간에 저장하면” 무손실성이 깨질 수 있음
     예를 들어 UFWP에서 DOCX를 뽑아 Word로 열고 저장했는데, Word가 앱 전용 보존 캡슐을 삭제/정리해버릴 수 있음.
  - 그래서 이 시스템은 보통:
     * UFWP를 canonical master
     * DOC/HWP/DOCX/HWPX는 import/export edge format
     으로 운용해야 함.
     이게 제일 중요.

2) byte-identical 재생성은 “편집 후”엔 다른 문제
     원본을 그대로 되돌리는 건 쉽지만, 중간에 내용을 수정했다면 편집 후 DOC/HWP를 원본 바이트와 똑같이 재생성하는 건 의미가 거의 없음.

  - 보장해야 할 것은:
     * 의미 동일성
     * 시각 동일성
     * 포맷별 기능 보존
  - 편집 후 파일 전체 바이트가 같아야 한다는 건 아님.

3) 매크로/스크립트/전자서명은 별도 취급이 필요
     * DOC는 VBA 등 활성 콘텐츠 가능
     * HWP 계열도 스크립트/확장 개체 이슈가 있을 수 있음
     * 전자서명은 수정 순간 원래 서명 효력이 깨짐

  - 따라서:
     * 활성 콘텐츠는 격리 저장
     * 서명은 보존 증적으로 별도 관리

8. 권장 아키텍처
  - 1단계: Canonical package 정의
     * ZIP 기반
     * UTF-8 manifest
     * SHA-256 해시
     * 리소스 분리 저장
     * 패키지형
       예시:
	mydoc/
	  manifest.json
	  document.json
	  styles.json
	  numbering.json
	  sections.json
	  notes.json
	  comments.json
	  changes.json
	  resources/
	    images/
	    embeddings/
	  residuals/
	    doc/
	    hwp/
	    docx/
	    hwpx/
	  provenance/
	    source-map.json
	  originals/
	    source.doc
	    source.hwp


  - 2단계: 공통 모델 정의
    처음부터 너무 넓게 잡지 말고 아래부터 시작:
     * 본문 텍스트
     * 문단/스타일
     * 표
     * 이미지
     * 머리말/꼬리말
     * 각주/미주
     * 섹션/페이지
     * 주석/변경추적
  - 3단계: fidelity capsule 정의
    우선 이 4종부터:
     * msdoc
     * ooxml-wordprocessing
     * hwp-binary
     * hwpx-xml
  - 4단계: provenance map
     * 노드 ID
     * 원본 경로
     * offset/recordId
     * dirty flag

  - 5단계: 시각 검증 계층
     * 폰트 해시
     * 렌더링 스냅샷
     * 페이지 단위 비교

9. 제일 현실적인 운영 모델
    아래 “2계층”으로 가는 게 좋음

1) 의미 계층
    검색/편집/분석용 공통 구조

2) 충실도 계층
    원본 재생성/역변환용 보존 정보

이렇게 해야
     * 시스템 내부에선 편집성과 검색성이 좋고
     * 외부 포맷 왕복에서도 손실을 최소화할 수 있음.

10. “보존 섬(opaque island)”
예를 들어 어떤 HWP 문서에 DOCX로 완전히 자연 변환하기 애매한 특수 개체가 있다고 가정하고,
그걸 내부 포맷으로 가져올 때:
     * 최대한 올리고
     * 완전히 이해 못한 부분은 opaque object로 보존
     * 에디터에서는 read-only 혹은 제한 편집
     * HWPX export 시엔 최대한 원형에 가깝게 다시 직렬화
     * DOCX export 시엔 시각 대체(이미지/텍스트박스 등) 또는 근사치 변환
이 방식이 현실적임.
즉, 모든 걸 완벽히 “이해해서 재구성”하려고 하지 말고,
이해 가능한 건 구조화하고, 애매한 건 보존 섬으로 남겨두는 것이 핵심.

12. “무리없이 HWPX나 DOCX로 변환”하려면 꼭 필요한 4가지
1) 폰트 통제
    환에서 제일 큰 오차 원인입니다.

같은 속성이라도
     * 폰트가 다르면
     * 줄바꿈이 달라지고
     * 페이지가 바뀌고
     * 표 높이가 바뀌고
     * 최종 레이아웃이 틀어짐
그래서 다음이 필요:
     * 폰트 fingerprint 저장
     * 대체 폰트 테이블
     * 서버/클라이언트 렌더링 환경 통제
     * 가능하면 기준 폰트 세트 고정
이걸 안 하면 포맷 설계를 잘해도 결과가 흔들림.

2) deterministic serializer
같은 내부 문서를 같은 옵션으로 저장하면
항상 같은 HWPX/DOCX가 나오도록 해야함.

그래야
     * 테스트가 가능하고
     * diff가 안정적이고
     * 회귀 검증이 쉬워짐

3) source mapping / dirty tracking
각 문서 노드에 대해:
     * 어디서 왔는지
     * 공통 모델로 완전 승격됐는지
     * 원본 전용 속성이 있는지
     * 사용자가 수정했는지를 기록
예:
     * clean: 가져온 뒤 안 건드림
     * modified: 수정됨
     * opaque: 구조 이해 못했지만 보존 중
     * degraded: 타 포맷 export 시 근사 변환 발생
     * 이 정보가 있어야 exporter가 똑똑해집니다.

4) 테스트 코퍼스
이런 시스템은 스펙 문서보다 테스트 문서 묶음이 더 중요.

최소한 이런 샘플이 필요합니다:

     * 일반 본문
     * 표 복잡한 문서
     * 머리말/꼬리말/쪽번호
     * 각주/미주
     * 이미지/도형/텍스트박스
     * 변경추적/주석
     * 목차/필드
     * 다단
     * 수식
     * 한글 조판 강한 문서
     * 오래된 DOC/HWP 샘플
검증은:
     * 텍스트 동일성
     * 스타일/구조 동일성
     * 페이지 수
     * 위치 오차
     * PDF 렌더 비교

13. 구현 전략 
  - 1단계
    DOCX/HWPX를 1급 시민으로 먼저 지원

    즉:
     * DOCX import/export
     * HWPX import/export
     * 내부 포맷 저장/로드
     * 기본 편집
    이걸 먼저 완성

    이유:
     * 둘 다 현대적 구조라서 훨씬 다루기 쉬움.
     * 공개 사양 기반으로 매핑 설계도 수월함.

  - 2단계
    DOC/HWP는 ingest 전용으로 추가
    즉:
     * DOC import
     * HWP import
     * 내부 포맷으로 정규화
     * 이후 출력은 HWPX/DOCX
    이렇게 가면 프로젝트 리스크가 크게 낮아짐.

  - 3단계
    고급 기능 추가
     * 변경추적
     * 주석
     * 수식
     * 텍스트박스/도형
     * 필드/목차
     * 고급 표
     * 일부 특수 조판

14. 이 프로젝트에서 꼭 피해야 할 것
1) 내부 포맷을 DOCX XML 그대로 쓰는 것
     * 처음엔 쉬워 보여도 금방 막힘.
     * HWPX/HWP 특성을 담기 어려워짐.

2) 내부 포맷을 HTML 기반으로 잡는 것
     * 페이지/워드프로세서 기능이 계속 새어나감.

3) “외부 앱 저장 후에도 내부 잔여정보가 안 깨질 것”이라고 기대하는 것
     * DOCX/HWPX에 앱 전용 확장을 넣을 수는 있어도,
     * 외부 편집기가 그걸 항상 보존한다고 기대하면 안됨.
     * 그래서 정본은 항상 내 포맷.

15. 운영 원칙
  - 우리 시스템의 정본은 내부 포맷이다.
  - HWP/DOC/HWPX/DOCX는 입력/출력 포맷이다.
  - 내부 편집은 공통 모델 기반으로 수행한다.
  - HWPX/DOCX로 내보낼 때는 지원 매트릭스 내 기능은 네이티브 변환하고,
  - 비호환 기능은 보존/근사/시각 대체 정책에 따라 처리한다.

16. 성공 기준
1) 보장 가능
  - 일반 업무 문서
  - 공공기관 스타일 문서
  - 표/이미지/각주/머리말/구역 포함 문서
  - 주석/변경추적이 있는 현대 문서
  - HWPX ↔ 내부 ↔ DOCX, DOCX ↔ 내부 ↔ HWPX의 고품질 변환
2) 별도 정책 필요
  - 매우 오래된 DOC/HWP 레거시
  - 매크로
  - OLE/ActiveX
  - 앱 특화 폼 컨트롤
  - 디지털 서명
  - 일부 복잡한 수식/차트/자동필드


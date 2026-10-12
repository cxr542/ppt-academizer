# Changelog

형식: [Keep a Changelog](https://keepachangelog.com/). 상세는 [releases/](releases/) 참고.

## [1.6.11] - 2026-10-12

Add: 경고 `PICTURE_DROPPED` — 슬라이드를 거의 덮는 큰 그림을 "디자인 배경"으로 판단해 제외했는데 그 그림이 **한 슬라이드에만 나올 때**(내용 이미지일 가능성) 알려 줍니다. 여러 슬라이드에 반복되는 배경은 기존처럼 조용히 제외합니다. 제거 동작 자체는 바꾸지 않았습니다.

## [1.6.10] - 2026-10-12

Fix: **첫 슬라이드(0번)에 표·차트나 긴 본문이 있으면 무조건 표지로 처리되어 사라지던 문제** — 이제 그런 경우 내용 슬라이드로 변환합니다. 제목만 있는 첫 슬라이드는 그대로 표지입니다.  
Add: 표가 줄어들면 글자 수와 무관하게 `CONTENT_LOSS` 경고(작은 표는 텍스트 40자 기준에 걸리지 않았음).  
운영 화면에서 합성 샘플로 검증하던 중 발견했습니다.

## [1.6.9] - 2026-10-12

Fix: **표가 변환 결과에서 사라지던 문제** — 표가 있는 슬라이드가 표지로 오분류되어(표지는 제목만 이식) 표와 본문이 버려졌습니다. 표·차트가 있거나 제목 외 텍스트가 많은(80자 초과) 슬라이드는 표지·간지로 분류하지 않습니다.  
Fix: PowerPoint 기본 "제목+내용" 레이아웃의 **본문 플레이스홀더가 통째로 버려지던 문제**, 거버닝 줄이 300자를 넘으면 같은 텍스트가 거버닝·본문 양쪽에서 제외되어 사라지던 문제.  
Add: 경고 `CONTENT_LOSS`(슬라이드 텍스트가 원본의 70% 미만으로 줄었을 때), `SLIDE_SKIPPED`(덱 종류 설정으로 슬라이드를 제외했을 때). 이전에는 조용히 버려졌습니다.  
실덱 7개 검증: 텍스트 보존 60%대 → 94~116%, 표 5개 전부 복원.

## [1.6.8] - 2026-10-12

Fix: 변환 결과 **표지·레이아웃의 로고가 "그림을 표시할 수 없습니다"(빨간 X)로 깨지던 문제**. 템플릿 로고는 PNG 대체 없이 SVG(`asvg:svgBlip`)만 참조하는데, 후처리가 확장을 지우면서 빈 `<a:blip/>`을 남겼습니다. PNG 대체가 없는 SVG 전용 그림은 그대로 보존하고, 그 SVG는 래스터화하지 않습니다(기존 래스터화는 흰 불투명 정사각형이 되어 있었음). Mac PowerPoint에서 팝업 없이 정상 표시 확인.

## [1.6.7] - 2026-10-12

Fix: 1.6.6에서 차트가 포함된 결과를 PowerPoint(Mac)로 열면 **"프레젠테이션 복구" 팝업**이 뜨던 문제 — 차트 `axId`/`crossAx` 정규화 값이 int32 범위(2^31-1)를 넘던 것을 절댓값 방식으로 수정. 실제 PowerPoint에서 수정 전 팝업·수정 후 정상 열림 확인.

## [1.6.6] - 2026-10-12

Fix: migrate 시 **차트가 깨지던 문제** — 차트 파트·임베디드 워크북을 함께 복사(이전에는 `r:id`가 끊겨 무관한 이미지 파트로 연결될 수 있었음).  
Fix: `profile=spec` 명시 선택 시 lab 덱 레이아웃이 적용되지 않던 문제, 폰트 크기 상속 덱의 표지·단일 텍스트 박스 목차 감지.  
Fix: Google PNG 슬라이드를 늘리지 않고 종횡비 유지·중앙 배치.  
Config: 덱 종류(파일명 패턴·skip 슬라이드·파트 표지·부제)를 `config/deck_kinds.json` / `PPT_ACADEMIZER_DECK_KINDS`로 추가·덮어쓰기 (`config/deck_kinds.example.json`). 암묵적 CMP 기본값 제거.

## [1.6.5] - 2026-05-21

Stability: 청크 업로드(413), API 임시 폴더 자동 삭제, 무거운 작업 `to_thread`, 서버 기본 `SKIP_PP_REPAIR`·`127.0.0.1`.  
Quality: AI·자유 텍스트 덱 → spec 라우팅 강화, `_PLACEHOLDER_GEOM` 동시 요청 누수 방지.

## [1.6.4] - 2026-05-21

UX: 웹·README에 **이미지는 아카데미 스타일로 변환하지 않음** 안내 문구 추가.

## [1.6.3] - 2026-05-21

Fix: SVG→PNG 후 `[Content_Types].xml`이 `image/svg+xml`로 남아 표지 장식 이미지가「그림을 표시할 수 없습니다」로 깨지던 문제.

## [1.6.2] - 2026-05-20

Fix: 슬라이드 레이아웃 SVG → PNG 자동 변환 — PP 자동 복구 실패 시에도 Mac「복구」팝업 완화.

## [1.6.1] - 2026-05-20

Fix: CMP 변환 후 Mac PowerPoint「복구」팝업 완화 — OOXML 검증 경로(`engine/office`), OOXML repair 후 재-sanitize, PP 자동 복구 대화상자 감지 강화.

## [1.6.0] - 2026-05-20

**입력 PPTX 전용:** Markdown·텍스트·붙여넣기 제거. 서비스 정체성을 «PPT → 아카데미 강의안 변환기»로 일원화.

## [1.5.2] - 2026-05-20

Fix: `engine/office/validators` 번들 동기화 — migrate 후 OOXML repair 시 `No module named 'validators'` 해결.

## [1.5.1] - 2026-05-20

Markdown ingest: `###` 슬라이드 분할, 목차 오인식 완화, 붙여넣기 사전 검사(반려/수정 제안).

## [1.5.0] - 2026-05-20

Phase 1 다중 입력: **Markdown·텍스트** 파일 및 **붙여넣기** → spec 파이프라인. [releases/v1.5.0-RELEASE.md](releases/v1.5.0-RELEASE.md) · [docs/md-authoring.md](docs/md-authoring.md)

## [1.4.1] - 2026-05-20

안정화: 업로드 기본 **50MB** (was 150), 표준 모드 기본 **40장** (was 80). UI `/health` 연동·클라이언트 용량 검사.

## [1.4.0] - 2026-05-20

품질 모드: **표준**(기본 최대 80장·품질 권장) / **대용량**(제한 없음·품질 미보장). UI 2단계 선택, `PPT_ACADEMIZER_MAX_SLIDES_STANDARD`.

**독립 실행:** `engine/` 번들, `sync_engine_from_ppt_test.py`, `tests/conftest.py`, CMP-like fixture·스모크 11/11. [releases/v1.4.0-RELEASE.md](releases/v1.4.0-RELEASE.md)

## [1.3.6] - 2026-05-19

migrate_cmp: 아카데미 마스터 **원본**(25년 교육자료 등) 변환 허용 — `academy-output` 재업로드만 차단.

## [1.3.5] - 2026-05-19

업로드 한도 50MB → **150MB** (환경 변수 `PPT_ACADEMIZER_MAX_UPLOAD_MB`).

## [1.3.4] - 2026-05-19

migrate_cmp: 숫자만 있는 헤더는 제목으로 보지 않음 · 제목/페이지 placeholder 세로 위치 수정 · 아카데미 출력 파일 재변환 차단.

## [1.3.3] - 2026-05-19

의존성: `requirements.txt`에 `defusedxml`, `lxml` 추가(OOXML repair 시 ModuleNotFoundError 방지).

## [1.3.2] - 2026-05-19

migrate_cmp(macOS): 변환 후 PowerPoint 자동 정리 패스 — SVG·미디어 패키지로 인한「복구」팝업 제거.

## [1.3.1] - 2026-05-19

migrate_cmp: OOXML whitespace repair 활성화 — PowerPoint「복구」팝업 완화(§7 partner 출력).

## [1.3.0] - 2026-05-19

UI: 2단계 마법사 — 1) 덱 유형 카드 선택·프로필별 미리보기, 2) 선택한 profile로 아카데미화. API `POST /wizard/preview`.

## [1.2.9] - 2026-05-19

라우팅: 파일명 대신 슬라이드 구조(도형·그림·표 밀도)로 spec vs migrate_cmp 판별.

## [1.2.8] - 2026-05-19

라우팅: CONTRABASS/CMP/PaaS 파일명 → `migrate_cmp`(§7) 우선(spec 텍스트 휴리스틱보다 먼저).

## [1.2.7] - 2026-05-19

spec: 2장 목차 오탐 수정 — `• 1976년:` 본문 불릿은 목차로 보지 않음(번호 챕터 색인만).

## [1.2.6] - 2026-05-19

spec: ph13 없음 오류 — placeholder 레이아웃 복제, 본문 ph13 삭제 금지.

## [1.2.5] - 2026-05-19

spec: 템플릿 ph12/ph13 폭 0.13in 붕괴 → 본문 밴드 전체 확장(세로 한 글자씩 나오던 현상).

## [1.2.4] - 2026-05-19

spec: 왼쪽 긴 본문→ph12/13, 짧은 라벨→거버닝, ph10 72자 초과 분리(`rebalance_content_placeholders`).

## [1.2.3] - 2026-05-19

spec/migrate: 본문이 ph10 왼쪽 제목으로 들어가던 문제 — 짧은 제목·본문 분리.

## [1.2.2] - 2026-05-19

apple-history 등 **텍스트 덱 → spec** 라우팅, 본문 2~4블록 휴리스틱, 이미지 슬라이드는 노트→본문.

## [1.2.1] - 2026-05-19

§6.6·§6.7 재수정: 차트 그림(88% 필터) 유지, `sync_raster_diagram_assets`, PICTURE XML 폴백, 본문 타이포 최종 패스.

## [1.2.0] - 2026-05-19

§6.6 본문 도형 타이포, §6.7 래스터 다이어그램 이식, 참고 메시지 한글·UTF-8 UX.

- [릴리즈 노트](releases/v1.2.0-RELEASE.md)

## [1.1.0] - 2026-05-19

migrate 품질: deck_kind, 슬라이드 1:1, 한글 제목, 표 §6.5, 차트 이식, fixture/audit 도구.

- [릴리즈 노트](releases/v1.1.0-RELEASE.md)

## [1.0.0] - 2026-05-19

첫 번째 이름 붙인 기준선 (CMP migrate 파일럿 + 웹 UI).

- [릴리즈 노트](releases/v1.0.0-RELEASE.md)
- [다음 계획 1.1](releases/v1.1.0-PLAN.md)

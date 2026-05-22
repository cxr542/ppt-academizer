# Changelog

형식: [Keep a Changelog](https://keepachangelog.com/). 상세는 [releases/](releases/) 참고.

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

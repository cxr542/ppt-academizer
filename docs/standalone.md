# ppt-academizer 독립 실행

## 결론

**독립 실행입니다.** `apps/ppt-academizer`만 복사·배포해도 웹 UI·API·변환 파이프라인을 돌릴 수 있습니다.
변환 엔진은 이 저장소의 `engine/`이 정본이며, 예전 `ppt-test`(2026-10-12 폐기)에 의존하지 않습니다.

## 엔진 경로


| 우선순위 | 경로 | 용도 |
|----------|------|------|
| 1 | `PPT_ENGINE_ROOT` / `PPT_TEST_ROOT` | 명시 오버라이드 |
| 2 | `apps/ppt-academizer/engine/` | **번들 엔진** (기본) |

둘 다 없으면 `FileNotFoundError`입니다(예전의 ppt-test 폴백은 오래된 사본으로 조용히 변환해 제거).

`core/ppt_test_path.ensure_engine_on_path()`가 `from scripts.*` import 전에 경로를 잡습니다.

## 번들 내용

- `engine/scripts/*.py` — 조립·ingest·migrate (이 저장소에서 직접 수정)
- `engine/docs/examples/*.json` — 스모크용 spec JSON
- `engine/ENGINE_VERSION` — 엔진 버전(역사적 기록. 동기화는 더 이상 하지 않음)

## ppt-test (폐기)

`ppt-test`는 2026-10-12에 보관 처리했습니다. 엔진 수정은 이 저장소의 `engine/`에서 하고
`python -m pytest tests/ -q`로 검증합니다. ppt-test에만 있던 테스트는 `tests/`로 옮겼습니다.
(`build_academy_deck.py` 같은 단독 실행 도구는 쓰지 않아 이식하지 않았습니다.)

## 외부 의존성 (ppt-test와 무관)

- Python 3.11+ · `requirements.txt`
- **아카데미 템플릿** `.pptx` — `TEMPLATE_PPTX` 또는 Spotlight/OneDrive 기본 경로 (`academy_template.py`)
- (선택) macOS PowerPoint — migrate 후 OOXML 정리

## 독립 배포 체크리스트

1. `engine/` 디렉터리 포함 (Dockerfile이 복사함)
2. `.venv` + `pip install -r requirements.txt`
3. `TEMPLATE_PPTX` 환경 변수 설정
4. `./scripts/run_server.sh`

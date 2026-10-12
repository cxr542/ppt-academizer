# ppt-academizer 독립 실행 (ppt-test 연동 검토)

## 결론

**가능합니다.** `apps/ppt-academizer`만 복사·배포해도 웹 UI·API·변환 파이프라인을 돌릴 수 있습니다.  
ppt-test는 **개발 시 엔진 소스**로 두고, 배포물에는 `engine/` 번들이 포함됩니다.

## 이전 구조 (연결 고리)

- `PYTHONPATH`에 `cursorstudy/experiments/ppt-test` 필수
- 스모크 JSON 예제·`scripts/*` import가 ppt-test에만 존재
- README가 ppt-test `.venv`의 Python을 가리킴

## 현재 구조

| 우선순위 | 경로 | 용도 |
|----------|------|------|
| 1 | `PPT_ENGINE_ROOT` / `PPT_TEST_ROOT` | 명시 오버라이드 |
| 2 | `apps/ppt-academizer/engine/` | **번들 엔진** (기본) |
| 3 | `cursorstudy/experiments/ppt-test` | monorepo 공동 개발 시 폴백 |

`core/ppt_test_path.ensure_engine_on_path()`가 `from scripts.*` import 전에 경로를 잡습니다.

## 번들 내용

- `engine/scripts/*.py` — 조립·ingest·migrate (ppt-test와 동기화)
- `engine/docs/examples/*.json` — 스모크용 spec JSON
- `engine/ENGINE_VERSION` — 동기화 시점·엔진 버전

## ppt-test가 여전히 필요한 경우

| 작업 | 이유 |
|------|------|
| 엔진 **소스 수정** | v1.6.6부터 **academizer의 `engine/`이 정본**(ppt-test 사본이 더 오래됨). 동기화는 수정을 되돌리므로 `--force` + 백업이 필요 |
| `build_academy_deck.py` CLI | academizer에 포함하지 않음 |
| `validate_pptx.py` 전체 CI | 선택 — academizer 스모크로 대부분 커버 |

## 외부 의존성 (ppt-test와 무관)

- Python 3.11+ · `requirements.txt`
- **아카데미 템플릿** `.pptx` — `TEMPLATE_PPTX` 또는 Spotlight/OneDrive 기본 경로 (`academy_template.py`)
- (선택) macOS PowerPoint — migrate 후 OOXML 정리

## 독립 배포 체크리스트

1. `engine/` 디렉터리 포함 (또는 배포 전 `sync_engine_from_ppt_test.py` 실행)
2. `.venv` + `pip install -r requirements.txt`
3. `TEMPLATE_PPTX` 환경 변수 설정
4. `./scripts/run_server.sh`

## monorepo에서 ppt-test만 수정했을 때

> ppt-test의 변경을 가져오려는 경우에만 해당합니다. `engine/`이 이미 있으면 `--force`가 필요하며,
> 기존 `engine/`은 `engine.bak-<시각>/`로 옮겨집니다. 가져온 뒤 백업과 diff해 academizer 쪽 수정을
> 다시 적용해야 합니다.

```bash
cd apps/ppt-academizer
python scripts/sync_engine_from_ppt_test.py --force
.venv/bin/python -m pytest tests/ -q
.venv/bin/python scripts/run_smoke_tests.py
```

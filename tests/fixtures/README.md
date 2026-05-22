# ppt-academizer test fixtures

대용량 파트너 덱은 Git에 넣지 않을 수 있습니다. 로컬에서 한 번 복사해 두면 스모크·회귀에 씁니다.

## 복사

```bash
cd apps/ppt-academizer
python scripts/restore_fixtures.py
```

`manifest.json`에 적힌 경로(기본: `~/Desktop/_삭제가능/바탕화면-잔여-PPTX/`)에서 찾아 `tests/fixtures/`로 복사합니다.

## 기대 파일

| 파일 | 용도 |
|------|------|
| `contrabass-partner.pptx` | CONTRABASS 1:1·제목·표 회귀 |
| `cmp-partner.pptx` | CMP migrate_cmp·목차 |
| `paas-v16.pptx` | PaaS 파일럿 (선택) |
| `cmp-like-partner.pptx` | 없으면 `make_cmp_like_fixture.py`가 생성 |
| `contrabass-academy-output-v1.0.pptx` | v1.0 출력 비교 (선택) |

## 감사(매핑표)

```bash
export TEMPLATE_PPTX="…"
python scripts/audit_slide_mapping.py \
  --source tests/fixtures/contrabass-partner.pptx \
  --output tests/fixtures/contrabass-academy-output-v1.0.pptx
```

CSV는 `output/audit/`에 저장됩니다.

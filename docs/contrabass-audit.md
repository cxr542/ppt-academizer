# CONTRABASS 슬라이드 감사 (1.1)

## 준비

```bash
cd apps/ppt-academizer
python scripts/restore_fixtures.py
export TEMPLATE_PPTX="/path/to/1.아카데미 강의안 템플릿(2026).pptx"
```

## 매핑표 (원본)

```bash
python scripts/audit_slide_mapping.py \
  --source tests/fixtures/contrabass-partner.pptx
```

`output/audit/*-source-*.csv` — 컬럼: `src_index`, `classify`, `in_plan`, `plan_kind`, `title`, shape counts.

## v1.0 출력과 비교 (선택)

```bash
python scripts/audit_slide_mapping.py \
  --source tests/fixtures/contrabass-partner.pptx \
  --output tests/fixtures/contrabass-academy-output-v1.0.pptx
```

## 1.1 변환 후 재감사

```bash
../../cursorstudy/experiments/ppt-test/.venv/bin/python \
  ../../cursorstudy/experiments/ppt-test/scripts/build_cmp_academy.py \
  --source tests/fixtures/contrabass-partner.pptx \
  --out output/contrabass-1.1-smoke.pptx

python scripts/audit_slide_mapping.py \
  --source tests/fixtures/contrabass-partner.pptx \
  --output output/contrabass-1.1-smoke.pptx
```

### 1.0 대비 확인 포인트

| 원본 | 1.0 증상 | 1.1 기대 |
|------|----------|----------|
| 7·8 (empty) | plan 제외 → 누락 | `in_plan=true`, `SLIDE_KEPT_EMPTY` 경고 |
| 18 | 영문만 ph10 | 한글 제목 유지 |
| 표 슬라이드 | scheme 색 깨짐 | §6.5 12pt·dk1 |
| 차트 | 미이식 | clone 또는 `CHART_NOT_COPIED` |

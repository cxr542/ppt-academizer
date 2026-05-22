# Pilot: spec vs migrate_cmp (cmp-like-partner fixture)

**Source:** `tests/fixtures/cmp-like-partner.pptx` (4 slides: cover, TOC, section, 2-column body)

**Detected profile:** `migrate_cmp` (filename contains `cmp`)

| Path | Slides | Notes |
|------|--------|--------|
| §5 `spec` + `front_matter_mode=auto` | 4 | Cover/TOC from source; body as placeholder text only |
| §7 `migrate_cmp` | 4+ | Shape transplant, 47% columns, academy colors |

**Symptom guide (full deck):**

- Wrong Docker/K8s TOC → fixed: `extract_front_matter` in `convert_legacy_deck_to_academy.py`
- Missing diagrams / wrong colors → use `profile=migrate_cmp` or auto for CMP filenames
- Google image-only slides → `profile=spec` + `GOOGLE_IMAGE_SLIDE` warning

Run comparison locally:

```bash
cd apps/ppt-academizer
../../cursorstudy/experiments/ppt-test/.venv/bin/python scripts/compare_deck_paths.py --source /path/to/your.pptx
```

Report JSON: `output/compare/compare-report-*.json`

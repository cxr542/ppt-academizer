# ppt-academizer

PPT **아카데미화** 서비스 (계획: [PLAN.md](PLAN.md) · 독립 실행: [docs/standalone.md](docs/standalone.md)).

**현재 릴리즈:** [1.6.0](releases/v1.6.0-RELEASE.md) · [CHANGELOG](CHANGELOG.md)

## 웹 UI

```bash
cd apps/ppt-academizer
python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
./scripts/run_server.sh
```

브라우저: **http://127.0.0.1:8765/** — **`.pptx` 업로드** → 덱 유형 확인 → **아카데미화** → 다운로드.

- **PPT 아카데미화란?** `.pptx`의 텍스트·슬라이드 구성을 아카데미 강의안 형식으로 맞춥니다. **단, 이미지(사진·차트 등)는 변환하지 않습니다.**
- **변환 방식 `자동`**: CMP·파트너 덱 → academy-design **§7** (도형 이식), Google 이미지 export → **§5** spec + 배경 이미지
- **2단계 마법사**: `POST /wizard/preview` → 프로필 선택 → `POST /academize`
- **품질 모드**: 표준(기본 최대 **40**장) / 대용량 — `PPT_ACADEMIZER_MAX_SLIDES_STANDARD` · 업로드 기본 **50**MB — `PPT_ACADEMIZER_MAX_UPLOAD_MB`
- **분석만**: `POST /analyze` — 슬라이드별 `layout`·`texts` 미리보기

`GET /health` — `engine_root`, `template_configured`, 버전 확인.

## 엔진 (ppt-test 독립)

변환 로직은 **`engine/`** 에 번들되어 있습니다. monorepo의 `ppt-test` 없이도 동작합니다.

```bash
# ppt-test에서 엔진을 갱신할 때만
python scripts/sync_engine_from_ppt_test.py
```

환경 변수: `PPT_ENGINE_ROOT` 또는 `PPT_TEST_ROOT`로 다른 엔진 경로 지정 가능.

## 스모크·테스트

```bash
cd apps/ppt-academizer
.venv/bin/pip install -r requirements.txt
.venv/bin/python -m pytest tests/ -q
.venv/bin/python scripts/run_smoke_tests.py
```

파트너 fixture 복원(로컬 원본 경로 — [manifest](tests/fixtures/manifest.json)):

```bash
.venv/bin/python scripts/restore_fixtures.py
```

필요 시:

```bash
export TEMPLATE_PPTX="/path/to/1.아카데미 강의안 템플릿(2026).pptx"
```

### 스모크 항목

| 이름 | 검증 내용 |
|------|-----------|
| `template_resolve` | 아카데미 2026 템플릿 경로 |
| `spec_build_mini` / `spec_build_apple` | §5 JSON → `save_academy_deck` |
| `save_slide_view` | §6.3 `lastView=sldView` |
| `legacy_to_spec` | AI형 fixture → heuristic spec → 빌드 |
| `cmp_like_spec_preview` | CMP형 fixture — `migrate_cmp` 라우팅 + spec 미리보기 |
| `cmp_like_migrate` | §7 migrate_cmp 슬라이드 수 |
| `contrabass_migrate` | CONTRABASS 1:1 (fixture 있을 때) |
| `apple_history_pptx` | Google 이미지 슬라이드 fixture |

```bash
.venv/bin/python scripts/compare_deck_paths.py --source tests/fixtures/cmp-like-partner.pptx
```

결과 pptx·JSON: `output/smoke-*` (타임스탬프 파일명).

**참고:** Google Slides **보내기 pptx**는 슬라이드 본문이 **배경 PNG**라 텍스트 도형이 없습니다.  
`pptx_ingest`가 배경 이미지를 추출해 `background_image`로 아카데미 슬라이드에 붙입니다.

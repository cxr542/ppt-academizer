# Deploy ppt-academizer API for EDU-TMS (Render)

TMS hosts the React UI (`module=academizer`). This repo hosts the **Python FastAPI** conversion API.

Netlify UI/proxy is **deprecated** for TMS. Point TMS env `PPT_ACADEMIZER_API_URL` at this API.

## Resume checklist (다음에 할 일)

로컬 스모크는 **2026-07-23**에 완료됨 (fixture #1 k8s lab · `127.0.0.1:8766`).  
엔진 `main`에 TMS CORS·본 문서 반영: [PR #2](https://github.com/cxr542/ppt-academizer/pull/2).

아직 안 한 것:

- [ ] Render Web Service (Dockerfile / `render.yaml`) from `cxr542/ppt-academizer` `main`
- [ ] Env: `PPT_ACADEMIZER_SKIP_PP_REPAIR=1`, CORS, **`TEMPLATE_PPTX`** (컨테이너 절대 경로; 템플릿 파일은 git에 넣지 않음)
- [ ] `GET https://<service>.onrender.com/health` → `ok` + `template_configured: true`
- [ ] EDU-TMS Vercel: `PPT_ACADEMIZER_API_URL=https://<service>.onrender.com`
- [ ] Prod E2E: `/admin?module=academizer` + `01_k8s_dashboard_lab_lecture.pptx`
- [ ] (선택) real_world fixture 2~5 — `docs/evaluation/real_world_fixture_evaluation.md`

TMS 쪽 SoT·백로그: `edu-team-tms/docs/ppt-academizer-tms.md`, `operations-backlog.md` §4c

## 1. Render Docker service

1. [Render](https://render.com) → New → Blueprint, or Web Service from `Dockerfile`
2. Repo: `cxr542/ppt-academizer`
3. Env:

| Key | Value |
|-----|--------|
| `PPT_ACADEMIZER_SKIP_PP_REPAIR` | `1` |
| `PPT_ACADEMIZER_CORS_ORIGINS` | `https://edu-team-tms-ten.vercel.app,http://localhost:3000` (+ preview origins as needed) |
| `TEMPLATE_PPTX` | Absolute path **inside the container** to the academy template `.pptx` |

4. Template options:
   - Mount/secret file into the image at e.g. `/app/templates/academy-2026.pptx` and set `TEMPLATE_PPTX=/app/templates/academy-2026.pptx`
   - Or bake a copy during a private build (do not commit the template to git)

5. Health: `https://<service>.onrender.com/health` → `ok: true`, `template_configured: true`

## 2. EDU-TMS (Vercel)

| Key | Value |
|-----|--------|
| `PPT_ACADEMIZER_API_URL` | `https://<service>.onrender.com` (no trailing slash) |

Optional client hint (same URL): `VITE_PPT_ACADEMIZER_API_URL`

Browser uploads go **directly** to the API (CORS). TMS `/api/academizer` only proxies **health** (Vercel body limits block large multipart).

## 3. Local

```bash
cd ppt-academizer
python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
# TEMPLATE_PPTX: Spotlight / mdfind for 「1.아카데미 강의안 템플릿.pptx」
export TEMPLATE_PPTX="$(mdfind 'kMDItemFSName == \"1.아카데미 강의안 템플릿.pptx\"' | head -1)"
PPT_ACADEMIZER_SKIP_PP_REPAIR=1 PORT=8766 ./scripts/run_server.sh
```

In TMS `.env.local`:

```bash
PPT_ACADEMIZER_API_URL=http://127.0.0.1:8766
VITE_PPT_ACADEMIZER_API_URL=http://127.0.0.1:8766
```

Fixture #1 (gitignored under `tests/fixtures/real_world/`):

```bash
# 예: ~/Documents/k8s_dashboard_lab_lecture.pptx
cp ~/Documents/k8s_dashboard_lab_lecture.pptx \
  tests/fixtures/real_world/01_k8s_dashboard_lab_lecture.pptx
```

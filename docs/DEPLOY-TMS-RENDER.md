# Deploy ppt-academizer API for EDU-TMS (Render)

TMS hosts the React UI (`module=academizer`). This repo hosts the **Python FastAPI** conversion API.

Netlify UI/proxy is **deprecated** for TMS. Point TMS env `PPT_ACADEMIZER_API_URL` at this API.

## Resume checklist (다음에 할 일)

로컬 스모크는 **2026-07-23**에 완료됨 (fixture #1 k8s lab · `127.0.0.1:8766`).  
엔진 `main`에 TMS CORS·본 문서 반영: [PR #2](https://github.com/cxr542/ppt-academizer/pull/2).

아직 안 한 것:

- [x] Render Web Service `ppt-academizer-api` → `https://ppt-academizer-api.onrender.com`
- [x] Env + **template fetch** → `/health` `template_configured: true` (private assets + deploy key)
- [x] EDU-TMS Vercel: `PPT_ACADEMIZER_API_URL=https://ppt-academizer-api.onrender.com`
- [x] Prod E2E: health proxy + fixture #1 academize 13장 (2026-07-31)
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

4. Template (Render secret files are **≤1 MB total** — the academy `.pptx` is ~6 MB, so do **not** upload the pptx as a secret file):

   - Private repo: [`cxr542/ppt-academizer-assets`](https://github.com/cxr542/ppt-academizer-assets) (`academy-template.pptx`)
   - Secret file on the Render service: filename `ppt_assets_deploy_key` = read-only **deploy key** (SSH private key) for that repo
   - Env: `TEMPLATE_PPTX=/tmp/academy-template.pptx` (default in Dockerfile)
   - Boot: `docker/entrypoint.sh` clones the assets repo over SSH and copies the pptx to `TEMPLATE_PPTX`

5. Health: `https://ppt-academizer-api.onrender.com/health` → `ok: true`, `template_configured: true`

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

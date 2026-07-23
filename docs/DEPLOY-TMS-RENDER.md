# Deploy ppt-academizer API for EDU-TMS (Render)

TMS hosts the React UI (`module=academizer`). This repo hosts the **Python FastAPI** conversion API.

Netlify UI/proxy is **deprecated** for TMS. Point TMS env `PPT_ACADEMIZER_API_URL` at this API.

## 1. Render Docker service

1. [Render](https://render.com) → New → Blueprint, or Web Service from `Dockerfile`
2. Repo: `cxr542/ppt-academizer`
3. Env:

| Key | Value |
|-----|--------|
| `PPT_ACADEMIZER_SKIP_PP_REPAIR` | `1` |
| `PPT_ACADEMIZER_CORS_ORIGINS` | `https://edu-team-tms-ten.vercel.app,http://localhost:3000` (+ preview origins as needed) |
| `TEMPLATE_PPTX` | Absolute path **inside the container** to the academy 2026 template `.pptx` |

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
# TEMPLATE_PPTX optional on Mac if Spotlight finds the academy template
PPT_ACADEMIZER_SKIP_PP_REPAIR=1 PORT=8766 ./scripts/run_server.sh
```

In TMS `.env.local`:

```bash
PPT_ACADEMIZER_API_URL=http://127.0.0.1:8766
```

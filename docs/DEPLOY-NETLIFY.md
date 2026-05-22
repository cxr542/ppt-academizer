# Netlify 배포 (팀 cxr542)

Netlify에는 **웹 UI(정적)** 만 올라갑니다. `.pptx` 변환 API는 **Python(FastAPI)** 이라 Netlify Functions로는 돌리기 어렵고, **별도 API 서버**가 필요합니다.

권장 구성:

| 역할 | 호스팅 |
|------|--------|
| UI | [Netlify](https://app.netlify.com/teams/cxr542/projects) |
| API | Render / Fly.io / 사무실 Mac(터널) 등 |

---

## 1. API 서버 먼저 (Render 예시)

1. [Render](https://render.com) → **New → Blueprint** 또는 Docker Web Service  
2. 이 저장소 연결, `render.yaml` 사용 또는 Dockerfile 경로 `./Dockerfile`  
3. 환경 변수 (필수):
   - `TEMPLATE_PPTX` — 아카데미 템플릿 `.pptx` **절대 경로** (컨테이너에 마운트/빌드에 포함하거나 Secret으로 경로 지정)
   - `PPT_ACADEMIZER_SKIP_PP_REPAIR=1` (기본)
4. 배포 후 URL 확인: `https://ppt-academizer-api.onrender.com/health` → `ok: true`, `template_configured: true`

로컬 Mac만 쓸 때: Cloudflare Tunnel / ngrok 등으로 `8765` 노출 후 그 URL을 API 주소로 사용.

---

## 2. Netlify 사이트 연결

1. https://app.netlify.com/teams/cxr542/projects → **Add new site** → **Import an existing project**  
2. GitHub `cxr542/ppt-academizer` (또는 monorepo면 **Base directory**: `apps/ppt-academizer`)  
3. Build settings (저장소에 `netlify.toml` 있으면 자동):
   - **Build command:** `bash scripts/netlify_prepare.sh`
   - **Publish directory:** `web`
4. **Environment variables** (Site settings → Environment variables):

| 변수 | 예시 | 설명 |
|------|------|------|
| `PPT_ACADEMIZER_API_URL` | `https://ppt-academizer-api.onrender.com` | 빌드 시 `_redirects`로 API 프록시 (끝 `/` 없이) |

5. **Deploy site**

브라우저는 Netlify 도메인에서 `/health`, `/academize` 등을 호출하고, Netlify가 API로 프록시합니다.

---

## 3. CORS (프록시 없이 API URL만 쓸 때)

`_redirects` 대신 `web/config.js`에 API 전체 URL을 넣는 방식이면 API에 허용 출처 설정:

```bash
PPT_ACADEMIZER_CORS_ORIGINS=https://your-site.netlify.app
```

---

## 4. CLI로 배포 (선택)

```bash
npm i -g netlify-cli
netlify login
cd apps/ppt-academizer   # 또는 저장소 루트
export PPT_ACADEMIZER_API_URL=https://your-api.example.com
bash scripts/netlify_prepare.sh
netlify init   # 팀 cxr542 사이트 연결
netlify deploy --prod
```

---

## 5. 동작 확인

- `https://<netlify-site>/` — UI 로드, 상단 한도 문구에 업로드 MB 표시 ( `/health` 성공 시 )
- `.pptx` 업로드 → 미리보기 → 아카데미화 → 다운로드

`/health` 실패 시 UI에 **API 미연결** 안내가 표시됩니다 → `PPT_ACADEMIZER_API_URL` 또는 API 서버·템플릿 경로를 점검하세요.

---

## 제한

- Netlify 빌드/Functions로 **python-pptx 변환 전체를 호스팅할 수 없음** (용량·시간·네이티브 의존성).
- API에 **아카데미 템플릿 파일**이 없으면 변환은 503 — `TEMPLATE_PPTX` 필수.

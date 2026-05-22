# Netlify에서 끝까지 쓰기 (빠른 설정)

사이트: **https://ppt-academizer.netlify.app/** (UI는 이미 배포됨)

변환까지 하려면 **인터넷에서 접속 가능한 API URL** 이 하나 더 필요합니다.

---

## 방법 A — Mac 터널 (가장 빠름, 무료)

Mac이 켜져 있고 터널 스크립트가 돌아가는 동안만 Netlify에서 변환됩니다.

### 1. 터미널 한 줄

```bash
cd apps/ppt-academizer   # 저장소 루트
./scripts/run_public_api_for_netlify.sh
```

잠시 후 화면에 **`https://xxxx.loca.lt`** 주소가 나옵니다. 이게 `PPT_ACADEMIZER_API_URL` 입니다.

### 2. Netlify 설정

1. https://app.netlify.com/projects/ppt-academizer  
2. **Site configuration** → **Environment variables** → **Add variable**  
   - **Key:** `PPT_ACADEMIZER_API_URL`  
   - **Value:** 스크립트에 나온 `https://xxxx.loca.lt` (끝에 `/` 없이)  
3. **Deploys** → **Trigger deploy** → **Deploy site**  
4. https://ppt-academizer.netlify.app/ 새로고침  

노란 **「API 미연결」** 배너가 사라지고, 상단에 `v1.6.5` 같은 버전이 보이면 성공입니다.

### 3. 사용

- Netlify 사이트에서 `.pptx` 업로드 → 변환 → 다운로드  
- **터미널을 끄면** API가 꺼져서 Netlify 변환이 다시 안 됩니다.

---

## 방법 B — Render (Mac 없이 24시간, 설정 더 많음)

1. [Render](https://render.com) → New **Web Service** → GitHub `cxr542/ppt-academizer`  
2. **Docker** · `Dockerfile`  
3. 환경 변수:
   - `TEMPLATE_PPTX` = 컨테이너 안 템플릿 경로 (템플릿 파일을 이미지에 넣거나 Disk에 업로드 필요)  
   - `PPT_ACADEMIZER_SKIP_PP_REPAIR=1`  
4. 배포 URL 예: `https://ppt-academizer-api.onrender.com` → Netlify `PPT_ACADEMIZER_API_URL`에 동일하게 입력 후 재배포  

자세히: [DEPLOY-NETLIFY.md](./DEPLOY-NETLIFY.md)

---

## 자주 하는 실수

| 증상 | 원인 |
|------|------|
| 노란 API 배너 | `PPT_ACADEMIZER_API_URL` 없음 또는 **재배포 안 함** |
| 미리보기 실패 | Mac 터널 종료됨 |
| 503 템플릿 | API 쪽 `TEMPLATE_PPTX` 없음 |

---

## 확인

**Netlify 사이트에서** (권장):

`https://ppt-academizer.netlify.app/health` → `{"ok":true,...}` JSON

`https://….loca.lt/health` 를 브라우저에 직접 치면 **localtunnel 안내 페이지**가 나올 수 있습니다.  
(Continue / IP 입력 화면 — 터널 보안용, API 오류 아님)

```bash
curl -H "bypass-tunnel-reminder: true" https://wild-suits-start.loca.lt/health
```

`"ok": true`, `"template_configured": true` 이면 OK.

# 사내 PC에서 엔진 운영 (Windows)

구성: `구성원 브라우저 → Tailscale Funnel(HTTPS) → 이 PC의 Docker 엔진(:8765)`

## 1. 준비물
- WSL2의 Docker Engine(무료) — Docker Desktop은 기업 규모에 따라 유료 라이선스가 필요할 수 있어 쓰지 않는다
- Tailscale
- 아카데미 템플릿·브랜드 파일(`cxr542/ppt-academizer-assets` 저장소의 `academy-template.pptx`, `brand/`)

### Docker 준비(WSL2, Windows 11)
이미 쓰는 WSL 배포판이 있으면 건드리지 않도록 전용 배포판을 따로 만든다.
```powershell
wsl --install -d Ubuntu-24.04 --name docker-host --no-launch
wsl -d docker-host -u root -- bash -c "apt-get update && apt-get install -y docker.io docker-compose-v2 git && systemctl enable --now docker"
```
Ubuntu 24.04 배포판은 `/etc/wsl.conf`에 `systemd=true`가 기본으로 들어 있다.

## 2. 설치
아래는 `docker-host` 안에서 실행한다(`wsl -d docker-host -u root`).
```bash
git clone https://github.com/cxr542/ppt-academizer.git
cd ppt-academizer/deploy/self-host
cp .env.example .env        # GEMINI_API_KEY 등을 직접 입력
mkdir assets                # academy-template.pptx 와 brand/ 폴더를 여기에 복사
docker compose up -d --build
```

## 3. 확인
```powershell
curl http://127.0.0.1:8765/health
```
`template_configured`, `brand_configured`, `llm_configured`가 모두 `true`인지, `service_version`이 최신인지 본다.

## 4. 외부 공개(HTTPS)
Tailscale 설치 후 관리 콘솔(Access → Settings)에서 HTTPS와 Funnel을 켠 다음:
```powershell
tailscale funnel --bg 8765
tailscale funnel status
```
출력된 `https://<pc이름>.<tailnet>.ts.net` 주소가 엔진 주소다. `/health`를 그 주소로 열어 확인한다.
Funnel은 인증 없이 인터넷에 열리므로, 주소는 TMS 설정과 구성원에게만 알린다.

## 5. 24/7 운영 설정
- 전원 옵션: 절전·최대 절전 끄기
- Windows 업데이트 활동 시간을 근무 시간으로, 재부팅은 새벽에
- 재부팅 후 자동 시작: Tailscale "항상 실행", WSL(`docker-host`)이 부팅 때 시작되도록 작업 스케줄러 등록(재부팅 시험 후 이 문서에 절차 추가)
- 컨테이너는 `restart: unless-stopped`라 Docker가 뜨면 같이 올라온다

## 6. 업데이트(엔진 새 버전 반영)
`docker-host` 안에서 실행한다.
```bash
cd ppt-academizer
git pull
cd deploy/self-host
docker compose up -d --build
curl http://127.0.0.1:8765/health
```

## 7. 문제 확인
```bash
docker compose logs --tail 100
docker stats --no-stream
docker compose restart
```

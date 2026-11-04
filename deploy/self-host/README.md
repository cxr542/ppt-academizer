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
- 컨테이너는 `restart: unless-stopped`라 Docker가 뜨면 같이 올라온다

### 재부팅 후 로그인 없이 자동 복구 (2026-11-04 재부팅 시험 통과)
WSL 배포판은 **계정별**이고 로그인 세션이 없으면 꺼진다. 로그인 없이 부팅만으로 올리려면 전용 일반 계정의 작업 스케줄러가 WSL을 시작하게 한다. 서버 터미널(관리자 명령 프롬프트)에서 순서대로 실행한다.

1. Tailscale이 로그인 전에도 연결되게 한다.
```powershell
tailscale set --unattended
```

2. 전용 계정을 만든다. 비밀번호는 프롬프트에 직접 입력하고 안전한 곳에 보관한다(문서·채팅에 적지 않는다).
```powershell
net user okestro * /add
net localgroup "Performance Log Users" okestro /add
powershell -NoProfile -Command "$u=[ADSI]'WinNT://./okestro,user'; $u.UserFlags.Value = $u.UserFlags.Value -bor 0x10000; $u.SetInfo()"
```
`Performance Log Users`는 "로그인 없이 도는 작업(batch 로그온)" 권한을 주려는 최소 권한 그룹이다(관리자 권한 불필요). 마지막 줄은 비밀번호가 만료돼 작업이 멈추는 것을 막는다.

3. 지금 쓰는 `docker-host`를 `okestro` 계정의 배포판으로 옮긴다(`docker-host`가 이미 다른 사용자 계정에 있을 때).
```powershell
mkdir C:\wsl
wsl --export docker-host C:\wsl\docker-host.tar
echo wsl.exe --import docker-host C:\wsl\okestro C:\wsl\docker-host.tar --version 2 ^> C:\wsl\import-okestro.log 2^>^&1 > C:\wsl\import-okestro.cmd
schtasks /create /tn "okestro-import" /sc once /st 00:00 /ru okestro /rp * /tr "C:\wsl\import-okestro.cmd" /f
schtasks /run /tn "okestro-import"
dir C:\wsl\okestro
```
`wsl --export`는 실행 중인 `docker-host`를 멈추므로 엔진이 잠시 내려간다. 1~2분 뒤 `C:\wsl\okestro\ext4.vhdx`(약 3GB)가 보이면 가져오기 성공이다(로그는 한글이 깨지므로 파일 존재로 판단).

4. 부팅 시 시작하는 작업을 만들고, 전환한다. **옛 엔진과 새 엔진을 동시에 켜지 않는다**(8765 포트가 겹친다).
```powershell
schtasks /create /tn "okestro-wsl-boot" /sc onstart /ru okestro /rp * /tr "wsl.exe -d docker-host --exec sleep infinity" /f
schtasks /end /tn "<옛 유지 작업 이름>" & wsl --terminate docker-host & schtasks /run /tn "okestro-wsl-boot"
schtasks /change /tn "<옛 유지 작업 이름>" /disable
```
옛 작업과 옛 배포판은 롤백용으로 남기고 삭제하지 않는다.

5. 재부팅(`shutdown /r /t 60`)한 뒤 **서버 화면에 로그인하지 않은 채로** `/health`가 200인지 본다. 2026-11-04 시험에서는 부팅 8초 뒤 작업이 시작되고 공개 `/health`가 200이었다.

### 알려진 제약
- WSL은 SYSTEM 계정에서 실행되지 않는다(`WSL_E_LOCAL_SYSTEM_NOT_SUPPORTED`).
- 비밀번호가 빈 계정은 로그인 없이 도는 작업에 쓸 수 없다(보안 정책 `LimitBlankPasswordUse=1`). 그래서 비밀번호가 있는 전용 계정이 필요하다.
- `okestro` 비밀번호를 바꾸면 `okestro-wsl-boot` 작업을 다시 만들어야 한다.
- 서버 재부팅 직후 StarDesk 같은 화면 원격 도구는 로그인 전까지 오프라인이다. 장애 진단은 SSH(Tailscale)로 한다.

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

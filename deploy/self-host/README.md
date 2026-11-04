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

## 8. 장애 알림과 대응

### 감시 설정 (서버 밖 모니터, 메일 알림)
서버 안에서 감시하면 서버가 죽을 때 알림도 같이 죽으므로 외부 업타임 모니터로 공개 주소를 본다. 2026-11-04에 UptimeRobot 무료 플랜으로 설정했다(Better Stack 등 다른 서비스도 같은 방식).

키워드 모니터는 모니터 하나에 키워드를 하나만 넣을 수 있어서 **키워드별로 모니터를 나눈다.**

| 항목 | 모니터 1 | 모니터 2 |
|------|----------|----------|
| 종류 | Keyword monitoring | Keyword monitoring |
| URL | `https://<pc이름>.<tailnet>.ts.net/health` | 같은 주소 |
| 키워드 | `"template_configured":true` | `"llm_configured":true` |
| 조건 | Start incident when keyword does not exist | 같음 |
| 간격 | 5분 | 5분 |
| 알림 | 메일 | 메일 |

- 키워드는 따옴표까지 그대로, 공백 없이 입력한다(응답이 `{"ok":true,...,"llm_configured":true,...}` 형태).
- 서버가 완전히 내려가면 두 모니터가 모두 실패해 메일이 두 통 온다. 502·타임아웃도 키워드가 없는 것으로 보고 같은 알림이 온다.
- 알림 수신자는 모니터마다 지정한다. 계정 메일만으로는 담당자가 받지 못하므로 담당자 메일을 알림 채널로 추가하고 두 모니터에 체크한다.
- "Test Notification"으로 메일이 오는지 확인한다. 서버를 건드리지 않고 시험하려면 키워드를 일부러 틀린 값으로 바꿨다가 되돌린다.
- 가입 때 자동 생성되는 루트 주소 HTTP 모니터는 `/health`와 키워드를 보지 않으므로 필요 없으면 지운다.
- 알림은 한 번만 실패해도 바로 온다("No delay"). 연속 실패 후 알림 설정은 플랜에 따라 다를 수 있어 확인하지 않았다.

키워드 검사가 필요한 이유: `.env`나 assets가 빠진 채 엔진이 떠도 `/health`는 200이라, 상태 코드만 보면 놓친다. 모니터에 넣는 주소는 문서·채팅에 실제 값을 적지 않는다.

### 알림을 받았을 때
1. 공개 `/health`를 직접 연다. 502면 WSL·Docker 정지, 타임아웃이면 서버·Tailscale 문제, 200인데 키워드 실패면 설정 누락이다.
2. 다른 기기에서 `tailscale status`로 서버(`okestroacademy`)가 온라인인지 본다.
3. 온라인이면 SSH(Tailscale)로 접속해 작업 `okestro-wsl-boot` 상태를 본다.
```powershell
schtasks /query /tn "okestro-wsl-boot" /v /fo list
```
4. 작업이 멈춰 있으면 다시 실행한다. 엔진이 안 올라오면 `okestro` 배포판 안에서 `docker compose restart`(7절).
```powershell
schtasks /run /tn "okestro-wsl-boot"
```
5. 서버가 오프라인이면 전원·네트워크를 확인한다. 화면 원격 도구(StarDesk 등)는 재부팅 직후 로그인 전까지 오프라인이므로 SSH를 먼저 쓴다.
6. 옛 `User` 계정의 `docker-host` 배포판은 켜지 않는다(8765 포트 충돌).

### 이 감시로 못 잡는 것
- 특정 파일만 변환에 실패하는 품질 문제
- Gemini 키 한도 초과(키 발급처에서 사용량·예산 알림을 따로 설정)
- WSL 디스크 부족(`/health`는 200이지만 변환이 실패할 수 있음) — 주기적으로 `docker system df`로 확인

# ppt-academizer 릴리즈

버전별 **릴리즈 노트**와 **다음 버전 계획**을 쌍으로 관리합니다.

| 파일 | 상태 | 설명 |
|------|------|------|
| [v1.6.0-RELEASE.md](v1.6.0-RELEASE.md) | **현재** | PPTX 입력 전용 (변환기 정체성) |
| [v1.5.0-RELEASE.md](v1.5.0-RELEASE.md) | 이전 | Markdown·텍스트·붙여넣기 (철회) |
| [v1.4.0-RELEASE.md](v1.4.0-RELEASE.md) | 이전 | 품질 모드·`engine/` 독립 번들 |
| [v1.2.0-RELEASE.md](v1.2.0-RELEASE.md) | 이전 | §6.6·§6.7·참고 UX |
| [v1.1.0-RELEASE.md](v1.1.0-RELEASE.md) | 이전 | deck_kind·1:1 슬라이드 |
| [v1.0.0-RELEASE.md](v1.0.0-RELEASE.md) | 기준선 | 최초 스냅샷 |

## 프로세스

1. 개발 전: `vX.Y.Z-PLAN.md` 작성·갱신 (필요 시)  
2. 구현·`pytest` + `run_smoke_tests.py` 통과  
3. `SERVICE_VERSION` bump · `scripts/sync_engine_from_ppt_test.py` (엔진 변경 시)  
4. `vX.Y.Z-RELEASE.md` 작성 → [CHANGELOG.md](../CHANGELOG.md)  
5. Git: `git tag -a ppt-academizer-vX.Y.Z -m "…"`

## Semver

- **1.0.x** — 호환 유지 버그 수정  
- **1.1.x** — migrate 품질  
- **1.2.x** — CMP 허브·본문 타이포  
- **1.4.x** — 품질 모드·독립 `engine/` 번들  

엔진 버전: `engine/ENGINE_VERSION` 또는 `engine/scripts/migrate_version.py`  
서비스: `core/version.py`

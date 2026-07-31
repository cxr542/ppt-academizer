# Real-world fixture evaluation

이 문서는 ppt-academizer MVP가 실제 OKESTRO 아카데미 강의안 변환 도구로 운영 가능한지 판단하기 위한 평가 체계입니다. 이번 평가는 변환 엔진을 수정하지 않고, 실제 보유 PPTX 5개를 고정 fixture로 삼아 현재 변환 품질을 측정하고 기록하는 데 목적이 있습니다.

## Fixture 위치

실제 fixture는 `tests/fixtures/real_world/` 아래에 둡니다. 테스트 안정성을 위해 ASCII 파일명을 사용하고, 원본 파일명은 아래 매핑표에 보존합니다.

민감한 회사 자료가 포함될 수 있으므로 외부 업로드나 외부 API 전송을 금지합니다. Git에 포함하기 어렵거나 파일 크기 정책에 걸리면 동일한 폴더 구조만 유지하고, 실제 PPTX는 로컬 보관소에서 복사하는 방식으로 운용합니다.

## Fixture 목록

| 우선순위 | 테스트용 파일명 | 원본 파일명 | 유형 | 테스트 목적 |
|---:|---|---|---|---|
| 1 | `01_k8s_dashboard_lab_lecture.pptx` | `k8s_dashboard_lab_lecture.pptx` | 실습 강의안형 PPT | 학습 목표, 실습 흐름, YAML, 체크리스트, 강사용 멘트가 실제 아카데미 강의안 형태로 유지되는지 확인한다. **로컬 API 스모크 통과 (2026-07-23, 13장 academize).** |
| 2 | `02_cmp_core_technology.pptx` | `클라우드 구현기술(CMP)_v1.0_수정요청.pptx` | 기술 개념 설명형 PPT | 개념 정의, 비유 설명, 기대효과 표의 제목/본문/표 구조가 안정적으로 유지되는지 확인한다. |
| 3 | `03_vmware_winback_strategy_report.pptx` | `오케스트로 VMware 윈백 시장 주도 전략 보고.pptx` | 전략 보고서형 PPT | 큰 제목, 숫자 지표, 로드맵, 전략 메시지가 강의안 템플릿으로 변환 가능한지 확인한다. |
| 4 | `04_academy_registration_page_plan.pptx` | `[기획안] 오케스트로 아카데미 교육신청 페이지_260128 v1.pptx` | 기획/요구사항형 PPT | 긴 텍스트, 표, CTA 설명, 관리자 요구사항, 폼 검증 문구가 읽을 수 있게 유지되는지 확인한다. |
| 5 | `05_contrabass_base_technology.pptx` | `CONTRABASS 기반기술@260504_수정 요청.pptx` | 복합 기술 발표형 PPT | Quorum, Galera, RAFT, OpenInfra 등 다중 기술 섹션과 긴 자료, 표, 이미지 출처가 깨지지 않는지 확인한다. |

## 수동 테스트 절차

1. 로컬 환경을 준비한다.

   ```bash
   python3 -m venv .venv
   .venv/bin/pip install -r requirements.txt
   ```

2. 필요하면 아카데미 템플릿 경로를 지정한다.

   ```bash
   export TEMPLATE_PPTX="/path/to/1.아카데미 강의안 템플릿(2026).pptx"
   ```

3. 로컬 서버를 실행한다.

   ```bash
   ./scripts/run_server.sh
   ```

4. 브라우저에서 `http://127.0.0.1:8765/`에 접속한다.
5. 우선순위 순서대로 fixture PPTX를 업로드한다.
6. preview 또는 analyze 결과에서 감지된 덱 유형, 예상 슬라이드 수, warning을 확인한다.
7. academize를 실행한다.
8. 변환 결과 PPTX를 다운로드한다.
9. PowerPoint 또는 LibreOffice에서 결과 PPTX를 연다.
10. 아래 평가표에 점수와 warning, 이상 현상을 기록한다.

## 자동 확인 가능 항목

| 항목 | 확인 방법 |
|---|---|
| 파일 존재 여부 | `tests/fixtures/real_world/*.pptx` 순회 |
| 원본 PPTX 열림 여부 | `python-pptx`로 `Presentation` 로드 |
| 원본 슬라이드 수 | `len(Presentation(...).slides)` |
| 변환 결과 PPTX 생성 여부 | `core.pipeline.academize_pptx` 출력 경로 확인 |
| 결과 PPTX 열림 여부 | 결과 파일을 `Presentation`으로 재로드 |
| 원본/결과 슬라이드 수 비교 | 원본 수와 결과 수를 결과 JSON/Markdown에 기록 |
| 파일 크기 0 byte 여부 | 원본/결과 파일 크기 기록 |
| 변환 warning 개수 | `academize_pptx` 반환 warning 리스트 길이 기록 |
| 변환 중 예외 발생 여부 | fixture별 오류 메시지 기록 |

## 수동 확인 필요 항목

| 항목 | 확인 관점 |
|---|---|
| 제목 배치 품질 | 제목이 슬라이드 상단/중심 구조에 맞게 배치되는가 |
| 본문 가독성 | 문장, bullet, 강사용 멘트가 읽을 수 있는 크기와 흐름인가 |
| 텍스트 잘림/겹침 | 긴 문구, 표 안 텍스트, 코드 블록이 잘리거나 겹치지 않는가 |
| 표/코드/이미지 시각 품질 | 표, YAML, 스크린샷, 출처 이미지가 의미를 잃지 않는가 |
| 아카데미 템플릿 적용 느낌 | OKESTRO 아카데미 강의안으로 보이는가 |
| 사람이 수정해서 쓸 수 있음 | 강사가 합리적인 수정량으로 실제 강의안화할 수 있는가 |

## 20점 평가표

각 항목은 0~2점입니다. 자동 확인 항목은 smoke 결과를 참고하되, 최종 점수는 사람이 결과 PPTX를 열어 확인한 뒤 기록합니다.

| Fixture | PPTX 열림 | 슬라이드 수 유지 | 제목 배치 | 본문 배치 | 잘림/겹침 없음 | 표/코드/이미지 | 템플릿 적용 | warning 적절성 | 다운로드 정상 | 수정해 사용 가능 | 총점 | 판정 | 메모 |
|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---|---|
| `01_k8s_dashboard_lab_lecture.pptx` |  |  |  |  |  |  |  |  |  |  |  |  |  |
| `02_cmp_core_technology.pptx` |  |  |  |  |  |  |  |  |  |  |  |  |  |
| `03_vmware_winback_strategy_report.pptx` |  |  |  |  |  |  |  |  |  |  |  |  |  |
| `04_academy_registration_page_plan.pptx` |  |  |  |  |  |  |  |  |  |  |  |  |  |
| `05_contrabass_base_technology.pptx` |  |  |  |  |  |  |  |  |  |  |  |  |  |

## 판정 기준

| 총점 | 판정 |
|---:|---|
| 18~20 | 바로 사용 가능 |
| 15~17 | 약간 수정하면 사용 가능 |
| 12~14 | 개선 필요 |
| 11 이하 | 운영 투입 어려움 |

## MVP 운영 가능 판단 기준

MVP를 실제 OKESTRO 아카데미 강의안 변환 도구로 투입하려면 최소한 다음 조건을 만족해야 합니다.

- 우선순위 1번 `k8s_dashboard_lab_lecture`가 15점 이상이어야 한다.
- 5개 fixture 모두 변환 결과 PPTX가 오류 없이 열려야 한다.
- 5개 중 3개 이상이 15점 이상이어야 한다.
- 11점 이하 fixture가 있다면 운영 투입 전 개선 대상 유형과 실패 원인을 별도 이슈로 분리해야 한다.
- warning은 사용자가 조치 가능한 문장이어야 하며, silent failure가 없어야 한다.

## 자동 smoke 실행

```bash
.venv/bin/python scripts/evaluate_real_world_fixtures.py
```

기본 출력:

- `outputs/evaluation/real_world_results.json`
- `outputs/evaluation/real_world_results.md`

변환을 수행하지 않고 원본 PPTX 열림과 메타데이터만 확인하려면:

```bash
.venv/bin/python scripts/evaluate_real_world_fixtures.py --metadata-only
```

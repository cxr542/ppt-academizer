# 다중 입력 지원 계획 (Markdown · Text · PDF)

> **취소 (v1.6.0):** 서비스는 **PPTX 변환기**로 입력을 일원화했습니다. 아래는 기록용입니다.

| 항목 | 내용 |
|------|------|
| **문서 버전** | 0.1 (검토용) |
| **작성일** | 2026-05-20 |
| **상태** | ~~Phase 1~~ **철회** (v1.6.0 PPTX only) |
| **전제** | [1.4.1](CHANGELOG.md) 안정화 한도 유지 (업로드 50MB · 표준 40장) 후 단계적 확장 |

---

## 1. 배경·목표

### 1.1 사용자 요구

- **Markdown (`.md`)**, **일반 텍스트 (`.txt`)**, **PDF** 를 넣어도 아카데미 `.pptx` 가 나오면 좋겠다.
- 현재는 **`.pptx` 만** 업로드 가능 (`api/main.py`, `web/index.html`).

### 1.2 제품 목표 (이번 확장)

| 목표 | 설명 |
|------|------|
| G1 | AI·메모·강의원고(md/txt) → 아카데미 강의안 pptx (품질은 표준 모드 한도 내) |
| G2 | PDF → **텍스트 위주** 슬라이드 분할 → spec 빌드 (1차) |
| G3 | 기존 pptx 경로(**spec + migrate_cmp**) 는 회귀 없이 유지 |
| G4 | **안정화 우선** — 화려한 PDF 레이아웃 복원·LLM 자동 구조화는 후순위 |

### 1.3 하지 않을 것 (1차~2차)

| 제외 | 이유 |
|------|------|
| PDF → `migrate_cmp` (§7) | §7은 **소스 pptx 도형·좌표** 이식 전제 |
| Google Slides API 직접 연동 | OAuth·별도 프로젝트 범위 |
| MD/PDF 실시간 WYSIWYG 편집기 | academizer는 **변환기** not 저작 도구 |
| 무제한 LLM 구조 분석 (기본) | 비용·재현성·품질 변동 — 옵션 플래그로만 후속 |

---

## 2. 현재 아키텍처 (As-Is)

```text
[.pptx 업로드]
      │
      ▼
detect_deck_profile() ──► migrate_cmp (§7) ──► academy.pptx
      │                        │
      └──► spec ◄──────────────┘ (google_image·단순 텍스트)
              │
              ▼
      convert_presentation() → SlideDeckSpec[]
              │
              ▼
      build_from_json_specs() + save_academy_deck()
```

- **중간 표준 형식**이 이미 있음: `SlideDeckSpec` JSON (`layout` + `texts` [, `background_image`]).
- 스모크: `engine/docs/examples/academy-deck-mini.json` → `build_from_json_specs` 검증 완료.
- **PLAN.md §5.1** 에도 md/txt는 “2단계”, JSON ingest 건너뛰기가 명시됨.

**핵심 결론:** 새 입력 형식은 **「pptx 이전 단계」에서 `SlideDeckSpec[]` 로만 맞추면** 이후 B4~B6는 재사용 가능.

---

## 3. To-Be 아키텍처

```text
                    ┌─────────────────────────────────────┐
                    │  InputAdapter (형식별)               │
                    │  pptx │ md │ txt │ pdf │ json(후속) │
                    └──────────────┬──────────────────────┘
                                   │
                    ┌──────────────▼──────────────────────┐
                    │  NormalizedDeck                      │
                    │  · title, subtitle                   │
                    │  · slides: [{title, body, images?}]  │
                    │  · source_format, warnings           │
                    └──────────────┬──────────────────────┘
                                   │
                    ┌──────────────▼──────────────────────┐
                    │  spec_builder (academy-design §5)    │
                    │  layout 분류: 표지·목차·간지·내지…    │
                    └──────────────┬──────────────────────┘
                                   │
              ┌────────────────────┼────────────────────┐
              ▼                    ▼                    ▼
        (pptx+partner)      spec only            spec only
        migrate_cmp         md/txt/pdf           json upload
              │                    │                    │
              └────────────────────┴────────────────────┘
                                   ▼
                         build_from_json_specs
                         save_academy_deck → .pptx
```

### 3.1 라우팅 규칙 (고정)

| 입력 | 파이프라인 | 비고 |
|------|------------|------|
| `.pptx` | **auto** (기존) | spec \| migrate_cmp \| google_image |
| `.md`, `.txt` | **spec only** | `profile=spec` 강제 |
| `.pdf` | **spec only** | 텍스트 추출 1차; 페이지 래스터 2차 |
| `SlideDeckSpec` JSON | **spec only** | ingest 생략 (Phase 2) |

---

## 4. 형식별 설계

### 4.1 Markdown (`.md`) — **Phase 1 (MVP)**

**의존성:** 표준 라이브러리 위주 (`re`, optional `markdown`/`mistune` — 제목·리스트 파싱).

**슬라이드 분할 규칙 (문서화 필수):**

| 규칙 | 동작 |
|------|------|
| `---` 단독 줄 | 슬라이드 구분자 (Marp 호환) |
| `# 제목` | 표지 후보 (문서 첫 H1) |
| `## 제목` | 간지 또는 내지 제목 |
| `###` 이하 + 본문 | `내지_거버닝 O` / `1_내지_거버닝 X` |
| `-` / `*` 리스트 | 본문 불릿 (`\n` 결합) |
| 코드 블록 | 본문 monospace — 1차는 일반 텍스트로 (스타일 미적용) |
| 이미지 `![](path)` | **1차 스킵** + warning `UNSUPPORTED_EMBED` (로컬 경로·URL) |

**표지·목차:**

- YAML front matter (`---` … `---`) optional: `title`, `subtitle`
- 없으면 첫 H1 → 표지, `##` 목록이 3개 이상이면 목차 슬라이드 자동 생성 (heuristic, `front_matter.py` 패턴 재사용)

**출력 품질 기대:** “강의 원고를 슬라이드로 쪼갠 초안” 수준. 수동 PowerPoint 보정 전제.

### 4.2 Plain text (`.txt`) — **Phase 1 (MVP)**

- 인코딩: UTF-8 (BOM 허용), 실패 시 `chardet` optional.
- 분할: 빈 줄 2개 또는 `===` / `---` 80열 구분선.
- 줄 시작 `1.` `2.` 패턴 → 목차 후보.
- 첫 비어있지 않은 줄 → 표지 제목 (50자 truncate).

MD와 **동일한 `NormalizedDeck`** 으로 합류.

### 4.3 PDF — **Phase 2 (텍스트) · Phase 3 (이미지)**

**Phase 2 — 텍스트 추출 (권장 1차 출시)**

| 항목 | 선택 |
|------|------|
| 라이브러리 | `pymupdf` (fitz) 1순위 — 텍스트+메타 안정 / 대안 `pdfplumber` |
| 페이지 → 슬라이드 | **1 page = 1 slide** (단순·예측 가능) |
| 레이아웃 | 페이지 상단 큰 폰트 블록 → 제목, 나머지 → 본문 (heuristic) |
| 표·다단 | 1차 무시 또는 plain text 붙임 + warning |
| 스캔 PDF | 텍스트 없으면 `PDF_NO_TEXT` — “이미지 PDF는 2단계” 안내 |

**Phase 3 — 페이지 래스터 (선택)**

- Google Slides export pptx 와 동일 패턴: 페이지를 PNG → `background_image`
- 품질·용량·50MB 한도와 충돌 → **표준 모드에서 페이지 수 제한** (예: 40페이지)
- 편집 불가 슬라이드 → warning 명시

**하지 않을 것 (PDF 1차):** 벡터 도형 복원, migrate_cmp.

### 4.4 기존 `.pptx`

- 변경 없음. `ingest` / `migrate` 는 `engine/scripts` 그대로.

---

## 5. API · UI 변경안

### 5.1 API

| 엔드포인트 | 변경 |
|------------|------|
| `POST /academize` | `UploadFile` + `input_format: auto \| pptx \| md \| txt \| pdf` (optional) |
| `POST /wizard/preview` | 동일 확장 — md/txt는 **spec 카드만** 표시 |
| `POST /analyze` | md/txt/pdf용 **구조 미리보기** (슬라이드 초안 JSON) |
| `POST /academize/spec` (신규, optional) | JSON body `SlideDeckSpec[]` 직접 (Phase 2) |

**공통:**

- MIME + 확장자 이중 검사.
- 슬라이드 수는 변환 **전** `slide_limits` 로 검사 (생성 예상 장수).
- 응답 `X-Academize-Source-Format: md|txt|pdf|pptx`.

### 5.2 웹 UI

```text
[ 탭: 파일 | 글 붙여넣기 ]

파일 탭: accept=".pptx,.md,.txt,.pdf"
글 붙여넣기: textarea + 형식 (md/txt) 라디오

1단계: 구조 미리보기 (슬라이드 N장, 표지/목차 감지)
2단계: 품질 모드 (기존) → 아카데미화
```

- pptx만 **프로필 카드 3종** (auto/spec/migrate); md/txt/pdf는 **「텍스트 강의안 → spec」** 단일 카드.

---

## 6. 코드 구조 (제안)

```
apps/ppt-academizer/
  core/
    ingest/
      __init__.py
      detect.py          # 확장자·MIME → InputFormat
      normalized.py      # NormalizedDeck dataclass
      md.py
      text.py
      pdf.py             # Phase 2+
    spec_from_normalized.py  # §5 layout 분류 (pptx ingest와 분리)
    pipeline.py          # academize_*() 분기 확장
  tests/
    fixtures/
      sample-lecture.md
      sample-lecture.txt
      sample-lecture.pdf   # optional, git-lfs or small
    test_ingest_md.py
    test_ingest_pdf.py
```

- `engine/` 번들은 **건드리지 않음** (pptx 전용 로직 유지).
- academizer `core/ingest` 만 추가 → 독립성 유지.

---

## 7. 품질·보안·한도

| 항목 | 정책 |
|------|------|
| 업로드 | 기존 50MB 유지; PDF는 페이지 수로 추가 제한 검토 |
| 표준 모드 | 생성 슬라이드 ≤ 40장 — 초과 시 거부 또는 분할 안내 |
| 대용량 모드 | 유지하되 md/txt/pdf에도 disclaimer |
| 악성 PDF | `pymupdf` 페이지 상한 (예: 200p read abort), 타임아웃 |
| 개인정보 | 서버 tmp 삭제 (현행 tempdir 패턴 유지) |

---

## 8. 단계별 로드맵

### Phase 1 — `v1.5.0` (MVP, 1~2주)

| # | 작업 | 완료 기준 |
|---|------|-----------|
| 1 | `NormalizedDeck` + `md`/`txt` 파서 | fixture md → ≥3 spec, validate_spec 통과 |
| 2 | `spec_from_normalized` | 표지·목차·내지 heuristic |
| 3 | `pipeline.academize_document()` | md/txt → pptx E2E |
| 4 | API + UI 파일 확장 | `.md`/`.txt` 업로드 성공 |
| 5 | 테스트·스모크 1건 | `smoke-md-mini` |
| 6 | 사용자 문서 | MD 슬라이드 작성 규칙 1페이지 |

### Phase 2 — `v1.5.1` (PDF 텍스트)

| # | 작업 | 완료 기준 |
|---|------|-----------|
| 1 | `pdf.py` 텍스트 추출 | 5페이지 이하 샘플 PDF → pptx |
| 2 | 스캔 PDF 안내 | 명확한 400 + 한글 메시지 |
| 3 | requirements | `pymupdf` 추가 |

### Phase 3 — `v1.6.0` (PDF 이미지·고급, 선택)

- 페이지 PNG 배경 슬라이드
- MD 이미지 embed (zip 업로드 또는 base64)
- `POST /academize/spec` JSON 직접

### Phase 4 — 후속 (안정화 이후)

- LLM 보조 구조화 (`confidence=low` 시 layout 제안 — PLAN §9)
- DOCX 입력
- 강의자료 생성 프로젝트와 JSON 파이프 연동

---

## 9. 리스크·완화

| 리스크 | 완화 |
|--------|------|
| MD 작성 규칙이 제각각 | 공식 **「아카데미 원고 템플릿」** md 예시 제공 |
| PDF 품질 기대 과다 | UI에 “텍스트 추출 초안” 라벨, pptx 수동 편집 안내 |
| migrate 기대와 혼동 | md/pdf 업로드 시 migrate 카드 **비표시** |
| 의존성 증가 | PDF는 Phase 2까지 optional extra `requirements-pdf.txt` |
| 회귀 | pptx fixture 스모크 11/11 CI 유지 |

---

## 10. 검증 계획

| 단계 | 테스트 |
|------|--------|
| 단위 | `tests/test_ingest_md.py`, `test_ingest_text.py` |
| 통합 | `academize_document(fixture.md)` → slide count, `lastView=sldView` |
| 회귀 | 기존 `pytest` 14 + `run_smoke_tests.py` 11 |
| 수동 | 웹에서 md 업로드 → 다운로드 pptx → PowerPoint placeholder 편집 |

---

## 11. 미결정 사항 (구현 전 확정 필요)

| ID | 질문 | 제안 기본값 |
|----|------|-------------|
| Q1 | MD 슬라이드 구분: `---` only vs `##`마다 새 슬라이드 | **`---` + `##` 둘 다** (`---` 우선) |
| Q2 | PDF 1페이지=1슬라이드 고정 vs 단락 병합 | **1페이지=1슬라이드** (안정) |
| Q3 | 붙여넣기 textarea 1차 포함 여부 | **포함** (파일과 동일 파이프) |
| Q4 | `pymupdf` vs `pdfplumber` | **pymupdf** |
| Q5 | 이미지 PDF 1차 차단 vs 래스터 시도 | **1차 차단 + 안내** (Phase 3) |

---

## 12. 요약

- **기술적으로 가능**하며, 기존 **`SlideDeckSpec` → build** 경로를 재사용하는 것이 맞다.
- **md/txt는 Phase 1**, **PDF는 Phase 2(텍스트) → Phase 3(이미지)** 로 나누어 **1.4.1 안정화 정책**과 맞춘다.
- **migrate_cmp는 pptx 전용**으로 유지해 품질·복잡도를 통제한다.

**다음 단계:** 위 미결정(Q1~Q5) 확인 후 Phase 1 구현 착수.

---
name: hwp-ops
description: HWP/HWPX 파일 읽기·플레이스홀더 치환·표 양식 채우기·포맷 변환. 정형 한글 양식 자동화. "hwp 읽어", "hwpx 채워", "한글 양식 자동화", "{{키}} 치환", "hwp → hwpx 변환" 등의 요청에 반응.
argument-hint: <file.hwp|file.hwpx> [data.json|data.md]
allowed-tools: Bash(python*), Bash(java*), Read, Write, Glob, Grep
---

# hwp-ops

한글(HWP/HWPX) 문서 자동화 스킬. 양식 채우기·치환·텍스트 추출·포맷 변환을 지원한다.

## 1. 개요

한글 포맷은 두 가지다.

- **HWP5** (`.hwp`): OLE Compound File + zlib 스트림 기반 바이너리. 직접 편집은 위험.
- **HWPX** (`.hwpx`): ZIP + XML (KS X 6101 OWPML). 편집 친화적.

이 스킬은 **HWPX 편집을 우선**하고, HWP5는 **읽기 + HWPX 변환**을 지원한다. 편집이 필요하면 먼저 HWPX로 변환하라.

## 2. 판별 먼저

모든 작업은 포맷 판별로 시작한다.

```bash
python scripts/hwp_detect.py <path>
```

출력 JSON의 `format` 필드로 분기:

- `hwpx` → HWPX 편집 워크플로우 (섹션 3)
- `hwp5` → 읽기(섹션 4) 또는 변환 후 편집(섹션 5)
- `unknown`/`zip`/`ole` → 사용자에게 확인 요청

## 3. HWPX 편집 워크플로우

### Phase 1 — 구조 파악

```bash
python scripts/hwpx_scan.py <file.hwpx> --placeholders --tables
```

출력에서 확인:

- `{{키}}` 플레이스홀더 목록과 **분할 여부** (`분할됨` 마크 있으면 run 병합 필요)
- 각 표의 `rowCnt × colCnt` + 셀 텍스트 요약

이 결과를 유저에게 보여주고 **어떤 키→값으로 치환할지, 어떤 표의 어떤 라벨에 값을 넣을지** 명시적으로 확인받아라.

### Phase 2 — mapping.json 설계

```json
{
  "placeholders": {
    "회사명": "V-Machina",
    "대표자": "홍길동",
    "신청일자": "2026-04-18"
  },
  "tables": [
    {"label": "신청일자", "value": "2026-04-18", "section": 0},
    {"label": "대표자명", "value": "홍길동", "section": 0, "occurrence": 0}
  ]
}
```

확인되지 않은 값은 `"[TODO]"` 문자열로 남겨라. 스킬은 이 마커를 보존한다.

### Phase 3 — 치환 실행

```bash
python scripts/hwpx_fill.py <in.hwpx> <out.hwpx> --json mapping.json
```

결과 출력에서:

- 치환 성공 횟수
- 매치 실패한 키 (있으면 Phase 1로 돌아가 스캔)
- 잔여 `[TODO]` 마커 목록 → 유저에게 보고

## 4. HWP5 읽기

```bash
python scripts/hwp5_extract.py <file.hwp>
```

폴백 체인:

1. `PrvText` 스트림 (미리보기, 빠름)
2. `BodyText/SectionN` 스트림 전체 추출 (정확, 느림)

옵션:

- `--prvtext-only` — 미리보기만
- `--format json` — JSON으로 출력

## 5. HWP → HWPX 변환

```bash
python scripts/hwp_to_hwpx.py <in.hwp> <out.hwpx>
```

백엔드 자동 선택:

1. `pyhwpx` (Windows + 한컴오피스 필요)
2. `hwp2hwpx.jar` (환경변수 `HWP2HWPX_JAR` + JRE 11+)

환경 없으면 명시적으로 실패하고 설치 방법을 안내한다.

## 6. 핵심 규칙 체크리스트 (v0.2 엔진)

치환 작업 시 반드시 지켜라. `hwpx_fill.py` v0.2는 모두 자동 처리.

- [ ] **UTF-8 no BOM, LF only** — XML 저장 인코딩
- [ ] **mimetype STORED + ZIP 첫 엔트리** — 어기면 한컴이 "손상된 파일입니다"
- [ ] **네임스페이스 URI 탐지** — 2011/2024 두 세대 존재. 파일별 실제 URI 사용, 하드코딩 금지
- [ ] **`<hp:linesegarray>` 제거** — 텍스트 바뀐 `<hp:p>`는 반드시 캐시 삭제 (레이아웃 재계산)
- [ ] **`<hp:run>` 구조·속성(`charPrIDRef`) 보존** — run을 병합해 첫 run에 몰아넣는 구구조 금지
- [ ] **`<hp:t>` 인라인 자식 보존** — `<hp:lineBreak/>`, `<hp:tab/>` 등 PUA sentinel로 치환·복원
- [ ] **표 라벨 접근 우선** — 논리 격자로 병합 셀 펼쳐 탐색. `>` path로 방향 이동
- [ ] **빈 셀 스타일 복제** — 주변 셀의 `paraPrIDRef`/`charPrIDRef` 가져다 쓰기
- [ ] **Scripts 스트림 절대 실행·디코드 안 함** — 보안
- [ ] **`[TODO]` 마커 보존** — 유저 확인 전까지 건드리지 말 것
- [ ] **`header.xml`/`version.xml` 불변** — 치환 로직이 건드리지 않음

## 7. 트러블슈팅

| 증상 | 원인 | 해결 |
|------|------|------|
| 한글에서 "손상된 파일입니다" | mimetype 미-STORED 또는 첫 엔트리 아님 | `hwpx_rezip.py` 재실행, `hwp_detect.py`로 `zip_ok` 확인 |
| 치환 후 줄바꿈/정렬 깨짐 | 이전 `hwpx_fill`의 `linesegarray` 미제거 버그 | **v0.2 엔진 재실행**. `linesegarray` 자동 제거됨 |
| 치환 후 서식 쏠림 (폰트/크기 섞임) | 이전 버전의 run 병합 로직 | **v0.2 엔진 재실행**. run 구조 보존 + charPrIDRef 유지 |
| 빈 셀에 넣은 값 폰트가 엉뚱함 | `paraPrIDRef`/`charPrIDRef` 미지정 | **v0.2 엔진 재실행**. 주변 셀 스타일 복제 |
| 새 요소가 한컴에 안 보임 | 네임스페이스 URI 하드코딩 (2011 vs 2024) | **v0.2 엔진 재실행**. 파일별 URI 탐지 |
| `<hp:lineBreak/>`·`<hp:tab/>` 사라짐 | 인라인 요소 보존 실패 | **v0.2 엔진 재실행**. PUA sentinel로 보존 |
| `{{키}}` 치환 안 됨 | run 분할 | `hwpx_scan.py --placeholders`의 `분할됨` 마크 확인. v0.2는 자동 splice |
| 네임스페이스가 `ns0:`로 바뀜 | lxml 미설치 → stdlib ET | `pip install lxml` 또는 v0.2의 nsmap 복원 로직 확인 |
| 라벨 다중 매칭 "중단" 경고 | 같은 라벨이 여러 셀에 | mapping.json에 `"occurrence": N` 추가 |
| 변환 실패 | pyhwpx/hwp2hwpx 환경 없음 | `hwp_detect.py`로 `env` 프로브 → 설치 |
| 한글 깨짐 (터미널) | Windows stdout 인코딩 | 스크립트 자체는 UTF-8 reconfigure. 리다이렉트 시 `chcp 65001` 권장 |

## 8. reference/

구현 상세가 막히면 아래를 Read하라.

- `reference/hwpx-anatomy.md` — HWPX ZIP/XML 구조, 네임스페이스, 표 구조
- `reference/hwp5-records.md` — HWP5 OLE 스트림, 레코드 헤더, zlib
- `reference/pitfalls.md` — 자주 틀리는 지점 모음
- `reference/verification.md` — 수동 검증 체크리스트

## 9. 스크립트 레퍼런스

| 스크립트 | 역할 |
|----------|------|
| `hwp_detect.py` | 포맷 판별 + 환경 프로브 |
| `hwpx_scan.py` | HWPX 구조·플레이스홀더·표 스캔 |
| `hwpx_rezip.py` | 디렉토리 → HWPX ZIP 재압축 (STORED 규칙 보장) |
| `hwpx_fill.py` | HWPX 치환·표 채우기 (모듈+CLI) |
| `hwp5_extract.py` | HWP5 텍스트 추출 |
| `hwp_to_hwpx.py` | HWP5 → HWPX 변환 |

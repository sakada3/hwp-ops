# 수동 검증 체크리스트

치환/변환을 마친 뒤 유저에게 인도하기 전 아래 항목들을 체크한다.

## 판별·진단

- [ ] `python scripts/hwp_detect.py sample.hwpx` → `format=hwpx`, `details.zip_ok=true`
- [ ] `python scripts/hwp_detect.py sample.hwp` → `format=hwp5`, `details.ole_ok=true`
- [ ] `--skip-env`로 빠른 판별 동작
- [ ] 환경 프로브가 없는 라이브러리를 거짓양성 없이 `false`로 보고

## HWPX 스캔

- [ ] `hwpx_scan.py` 기본 실행 시 ZIP 엔트리 목록과 **mimetype STORED: OK** 출력
- [ ] `--placeholders`가 실제 문서에 보이는 `{{키}}`를 빠짐없이 잡음
- [ ] 한컴에서 중간 편집 후 저장된 **분할된 `{{키}}`**도 `분할됨` 마크로 식별
- [ ] `--tables`가 `rowCnt × colCnt`와 각 셀 텍스트 요약 출력
- [ ] `--section N`으로 단일 섹션만 필터링
- [ ] `--raw-xml Contents/section0.xml`로 원본 XML 확인 가능

## HWPX 치환

- [ ] 단일 `{{키}}`(온전 케이스) 치환 후 결과 파일이 한컴에서 정상 열림
- [ ] 분할된 `{{키}}` 치환 성공 (run 병합 알고리즘 검증)
- [ ] 표 라벨 예: "신청일자" 옆 셀에 `2026-04-18` 주입 성공
- [ ] 병합 셀 있는 표에서도 **라벨 기반 접근이 실패하지 않음**
- [ ] `[TODO]` 마커가 있는 값은 **치환 안 하고 보고서에 나열**
- [ ] 매치 실패한 키가 있으면 **exit code 1 + 안내 메시지**
- [ ] 치환 대상 외 XML 엔트리는 **바이트 동일** (사전/사후 해시 비교)
  - `python -c "import hashlib; print(hashlib.sha256(open('x','rb').read()).hexdigest())"`
- [ ] `BinData/*`, `Scripts/*` 엔트리 크기/내용 불변
- [ ] 결과 파일을 `hwpx_scan.py`로 재스캔 → mimetype STORED + 첫 엔트리 OK

## HWPX 재압축

- [ ] `hwpx_rezip.py <dir> <out>` 실행 후 "OK" 출력
- [ ] `python -c "import zipfile; z=zipfile.ZipFile('out.hwpx'); print(z.infolist()[0])"` 로
      첫 엔트리가 `mimetype`, `compress_type=0` 확인
- [ ] 출력 파일을 한컴오피스에서 열어 경고 없이 표시

## HWP5 읽기

- [ ] `hwp5_extract.py sample.hwp` → 최소 1000자 이상 추출
- [ ] `--prvtext-only`로 미리보기만 빠르게
- [ ] `--format json`으로 기계 친화 출력
- [ ] 압축/비압축 양쪽 파일에서 동작 (FileHeader 플래그 자동 인식)
- [ ] 다중 섹션 문서에서 모든 섹션 추출
- [ ] 깨진 한글·누락 없음 (문단 단위로)

## HWP → HWPX 변환

- [ ] 환경 있을 때 (`pyhwpx` 또는 `hwp2hwpx.jar`) 변환 성공
- [ ] 환경 **없을 때 명확한 한국어 에러 메시지** 출력 후 exit 1
- [ ] `--backend pyhwpx` 명시 시 hwp2hwpx로 폴백하지 않음
- [ ] `--backend hwp2hwpx` 명시 시 pyhwpx 시도 안 함
- [ ] 변환 결과 HWPX를 스캐너·치환에 그대로 투입 가능

## 보안

- [ ] Scripts 스트림이 있는 HWP에서 **스캐너가 경고 출력**하고 **실행은 하지 않음**
- [ ] `BinData/` 바이너리를 파싱/실행하지 않음
- [ ] XML 파서에서 외부 엔티티(DTD) 처리 금지 (XXE 방어)

## 인코딩·개행

- [ ] 치환 후 XML 파일들이 UTF-8 (BOM 없음)
- [ ] 개행이 LF only (`\r\n` 없음) — `grep -U $'\r' Contents/*.xml` 결과 0
- [ ] 한컴 특수문자 (예: 전각 공백, 조판부호) 보존

## 리포팅

- [ ] 치환 스크립트 실행 후 **치환 건수·실패 키·TODO 목록** 명시
- [ ] 변환 스크립트 실행 후 사용된 backend 명시
- [ ] 스캐너 출력이 사람이 읽기 좋은 한국어 리포트

## CLI 동작

- [ ] `python hwp_detect.py` (인자 없음) → argparse usage 출력 + exit 2
- [ ] `--help` 플래그 각 스크립트에서 동작
- [ ] 경로에 공백/한글 있어도 정상 동작

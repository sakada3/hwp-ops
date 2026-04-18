# 자주 틀리는 지점 모음

## v0.2 레이아웃 보존 TOP 10 체크리스트

치환 엔진이 레이아웃을 깨지 않기 위해 반드시 지켜야 할 10가지. `hwpx_fill.py` v0.2는 모두 자동.

1. 네임스페이스 URI는 **파일에서 탐지** (2011/2024 양쪽). 하드코딩 금지.
2. 텍스트 수정한 `<hp:p>`에서 `<hp:linesegarray>` **제거** (렌더링 캐시 무효화).
3. `<hp:run>` 구조·속성(`charPrIDRef`) **훼손 금지**. run을 병합해 첫 run에 몰아넣지 말 것.
4. `<hp:t>` 내부 인라인 자식(`<hp:lineBreak/>`, `<hp:tab/>`, `<hp:markpenBegin/>` 등) **보존**.
5. 앞뒤 공백 있는 value → `xml:space="preserve"` 부착.
6. 빈 셀 채우기 시 주변 셀의 `paraPrIDRef`·`charPrIDRef` **복제** (기본값 0 하드코딩은 최후 수단).
7. 표 라벨은 **정규화 후 매칭** (공백 축약 + 끝 콜론 제거). 병합 셀은 논리 격자로 펼쳐 탐색.
8. run이 부분 치환될 때 `charPrIDRef`를 **다시 명시적으로 set** (파서가 속성 날리는 케이스 방어).
9. `header.xml`·`version.xml`은 **읽지도 쓰지도 않음**.
10. 저장 시 **원본 nsmap 유지**, UTF-8 no BOM, LF only, mimetype STORED + 첫 엔트리.

## HWPX 재압축 규칙

- **mimetype STORED + 첫 엔트리**: 어기면 한컴이 "손상된 파일입니다"로 거부.
  `zipfile.ZipInfo("mimetype")` + `compress_type = ZIP_STORED` + 맨 먼저 write.
- **UTF-8 no BOM**: XML 저장 시 `\xef\xbb\xbf` 금지. lxml은 기본적으로 BOM 안 씀.
  stdlib ET는 encoding="utf-8" 지정 시 BOM 없이 저장됨. **UTF-16 금지**.
- **LF only**: `\r\n`을 `\n`으로 바꿔 저장. Windows에서 파일을 텍스트 모드로 쓰면
  CRLF가 되므로 **바이너리 모드 or 사후 변환** 필요.

## XML 네임스페이스 (v0.2 강화)

- **URI 세대 2개 존재**: 2011 (`hancom.co.kr/hwpml/2011/...`)와 2024 (`owpml.org/owpml/2024/...`).
  파일마다 다르고, 한컴 구버전은 2011, 신버전은 2024. 섞여 있을 수도 있음.
- **네임스페이스 하드코딩 금지**. 새 요소 생성 시 하드코딩된 URI로 만들면 원본과 불일치 →
  한컴이 요소 인식 못 함. `hwpx_fill.py`는 파일별 실제 URI를 탐지해 `SectionDoc.ns_hp`에 저장.
- stdlib `xml.etree.ElementTree`는 원본 prefix 대신 `ns0:`, `ns1:`로 재작성.
  한컴이 이를 거부한 사례가 보고됨. **`pip install lxml` 권장**.
- `hwpx_fill.py` v0.2는 저장 후 원본 nsmap 기준으로 `ns0:` → `hp:` 재치환으로 방어.

## 플레이스홀더 치환 (v0.2 알고리즘)

### 현재 구현(v0.2)이 방어하는 것

- `{{키}}` 분할(run 경계 걸침) → 자동 감지 후 splice. **run을 병합하지 않고 구조 유지**.
  첫 교차 run에 value 삽입, 나머지 교차 run의 t.text에서 target 조각만 제거.
- `<hp:run>`의 `charPrIDRef` 등 속성 **보존**. 치환 후 명시적으로 다시 set.
- `<hp:linesegarray>` 자동 제거 (텍스트/구조 변경된 `<hp:p>`에만). **한컴 포럼 1677 준수**:
  linesegarray는 렌더링 캐시 — 재계산을 위해 제거해야 줄바꿈·정렬이 올바름.
- `<hp:t>` 인라인 자식(`<hp:lineBreak/>`, `<hp:tab/>`, `<hp:markpenBegin/>`, `<hp:fieldBegin/>` 등) 보존.
  PUA sentinel 문자(`\uE000`)로 위치를 표시 → 치환 후 원래 요소로 복원.
- 치환 value에 `\n` → `<hp:lineBreak/>` 자동 분할 삽입.
- 치환 value의 제어문자 제거, 탭 공백 변환.
- 앞뒤 공백 있는 value(` 홍길동 `) → 해당 `<hp:t>`에 `xml:space="preserve"` 자동 부착.
- 빈 셀 채우기 시 주변 셀(왼쪽 라벨 셀 → 같은 표의 최다 사용 스타일) 속성을 템플릿으로 복제.
  `paraPrIDRef`·`charPrIDRef`·`subList.id(제외)` 등.

### 유저가 주의할 것

- 치환 대상 0건이면 `hwpx_scan.py --placeholders`로 **분할 상태 재확인**.
- 출력 경고에 "서식 쏠림 가능"이 뜨면 → 템플릿에서 `{{키}}`를 같은 서식으로 다시 입력 권장.
- `[TODO]` 마커가 value에 있으면 치환 생략(보존) — 유저 확인 전까지.

## 표 채우기 (v0.2 강화)

- `rowCnt`, `colCnt`는 논리 좌표. 실제 `<hp:tc>` 수는 병합 때문에 더 적을 수 있음.
- **라벨 접근 우선**: 라벨 셀 찾고 → 같은 `<hp:tr>`의 다음 `<hp:tc>`.
- 인덱스 접근(`[row][col]`)은 병합 셀 있으면 틀어짐. v0.2는 내부에서 논리 격자로 변환해 안전.
- **라벨 정규화**: 공백 축약(`\s+` → ` `) + 끝의 `:`·`：` 제거 + strip. "신청일:" == "신청일".
- **방향 path 지원**: `"성명 > right > down"` 같은 경로. 기본은 `right`. 병합된 공석 좌표는
  같은 방향으로 자동 스킵.
- **occurrence 지정**: 라벨이 여러 번 나오면 mapping.json에 `"occurrence": 1` 추가. 지정 없이
  다중 매칭이면 경고 출력 후 중단(기존 "첫 매치만" 동작과 다름 — 명시적 안전 모드).
- 빈 셀은 `<hp:subList>` 자체가 없을 수 있음. v0.2는 주변 셀 스타일을 템플릿으로 복제해
  `paraPrIDRef`·`charPrIDRef` 채운 상태로 구조 생성. **기본값 0 하드코딩은 최후 수단**.

## 보안

- `Scripts/*` 스트림: 매크로 JScript. **절대 실행·eval·파싱하지 말 것**.
- `BinData/*`: 임의 바이너리. 치환 작업에서는 건드리지 말고 그대로 재압축.
- 외부 네트워크/파일 접근: XML External Entity 공격 위험. lxml은 기본 `resolve_entities`
  있는데, 치환 파이프라인에선 `XMLParser(resolve_entities=False)` 고려.

## `header.xml`과 연동

- `Contents/header.xml`에 폰트 ID, 글자 속성 ID, 표 테두리 ID 등이 정의됨.
- `section*.xml`의 `charPrIDRef`, `paraPrIDRef` 등은 header의 ID를 참조.
- **본문만 수정**하면 안전. header 수정하면 여러 참조가 어긋남. 지양.
- `hwpx_fill.py` v0.2는 `header.xml`·`version.xml`을 읽지도 쓰지도 않음 (불변 보장).

## HWP5 직접 편집

- HWP5는 레코드 구조라 삽입·삭제 시 후속 offset과 size 연쇄 변경 필요.
- Level이 트리 깊이를 표현하는데, 삽입 순서를 틀리면 한컴이 로드 실패.
- **현실적으로 HWP5 편집은 지양**. 반드시 HWPX 변환 후 편집.

## Windows + Python 관련

- 터미널 깨짐: `sys.stdout.reconfigure(encoding="utf-8")` 각 스크립트 상단.
  리다이렉트 시 `chcp 65001` 권장.
- OLE 읽기: `olefile.OleFileIO(str(path))` — Path 객체 직접 전달보다 str 안전.
- zlib raw deflate: `zlib.decompress(data, -15)` 또는 `decompressobj(wbits=-15)`.
  **기본값(15)로는 실패**.
- pyhwpx는 한컴오피스가 **실제로 설치되어 있어야** 동작. 설치 없으면 Dispatch 실패.

## `[TODO]` 마커 처리

- mapping.json의 value에 `"[TODO]"`가 있으면 **치환하지 말고 보존**.
- 유저가 확인한 뒤 값을 채우도록 보고서에 명시.
- 최종 산출물에 `[TODO]`가 남아있으면 반드시 보고.

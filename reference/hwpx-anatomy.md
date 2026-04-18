# HWPX 해부도

HWPX = ZIP 컨테이너 + 여러 XML 파일. 표준은 **KS X 6101 (OWPML)**.

## ZIP 구조

필수 구조는 ODF와 유사하다. 첫 엔트리는 반드시 `mimetype` (STORED, 즉 비압축).

```
<file.hwpx>
├── mimetype                      # application/hwp+zip (STORED, 첫 엔트리)
├── META-INF/
│   ├── container.xml             # rootfile 경로 지정
│   └── manifest.xml              # 파일 목록·ID·타입
├── Contents/
│   ├── content.hpf               # OPF-like 패키지 문서
│   ├── header.xml                # 문서 전역 설정 (폰트, borderFill, charPr 등)
│   ├── section0.xml              # 본문 섹션 0
│   ├── section1.xml              # 본문 섹션 1 (있을 때)
│   └── ...
├── Preview/
│   └── PrvImage.png              # 미리보기 이미지 (선택)
├── BinData/                      # 포함 바이너리 (이미지 등)
│   ├── image1.png
│   └── ...
├── settings.xml                  # 뷰 설정
└── version.xml                   # 버전 정보
```

## mimetype 규칙 (매우 중요)

- 파일명: `mimetype` (소문자, 확장자 없음)
- 내용: `application/hwp+zip` (개행 없어도 되고, 있어도 됨)
- ZIP 엔트리 **첫 번째**여야 함
- 압축 방식은 **STORED (0)**, DEFLATE 금지
- 어기면 한컴이 "손상된 파일입니다"로 거부

## 주요 네임스페이스 (버전 2세대 — 2011 vs 2024)

네임스페이스 URI는 한컴 버전에 따라 두 가지가 존재한다. 파일마다 다르고, 두 세대가
섞여 있을 수도 있다.

### 2011 네임스페이스 (대다수 실제 파일)

```
hp : http://www.hancom.co.kr/hwpml/2011/paragraph
hh : http://www.hancom.co.kr/hwpml/2011/head
hc : http://www.hancom.co.kr/hwpml/2011/core
hs : http://www.hancom.co.kr/hwpml/2011/section
```

### 2024 네임스페이스 (신규 한컴 저장본)

```
hp : http://www.owpml.org/owpml/2024/paragraph
hh : http://www.owpml.org/owpml/2024/head
hc : http://www.owpml.org/owpml/2024/core
hs : http://www.owpml.org/owpml/2024/section
```

**핵심 규칙**: 새 요소를 생성할 때는 **원본 파일이 실제로 사용하는 URI**를 그대로 써야 한다.
하드코딩된 2011 URI로 요소를 만들어 2024 파일에 삽입하면 한컴이 요소를 인식 못 한다.
`hwpx_fill.py` v0.2는 루트의 `nsmap`/`xmlns` 선언에서 URI 키워드(`paragraph`,`head` 등)를
탐지해 `SectionDoc.ns_hp` 등에 저장한 뒤 해당 URI로 요소를 만든다.

각 XML은 자신의 루트에서 prefix를 선언한다. stdlib `xml.etree`는 prefix를 `ns0:`, `ns1:`
식으로 재작성하기 쉬우므로 `lxml` 사용을 강력히 권장한다.

## 본문 paragraph 구조

```xml
<hp:sec>
  <hp:p id="..." paraPrIDRef="0" styleIDRef="0">
    <hp:run charPrIDRef="0">
      <hp:t>안녕하세요.</hp:t>
    </hp:run>
  </hp:p>
  <hp:p>
    <hp:run charPrIDRef="0">
      <hp:t>회사명: </hp:t>
    </hp:run>
    <hp:run charPrIDRef="1">
      <hp:t>V-Machina</hp:t>
    </hp:run>
  </hp:p>
</hp:sec>
```

- `<hp:p>` = 문단
- `<hp:run>` = 문자 속성이 동일한 텍스트 구간
- `<hp:t>` = 실제 문자열

## 핵심 속성 · 요소 역할

- **`paraPrIDRef`** (`<hp:p>` 속성): header.xml의 paraPr(문단 속성)를 참조하는 ID.
  정렬·들여쓰기·줄간격 등을 결정. 빈 셀 생성 시 주변 값을 복제해야 레이아웃 일관.
- **`charPrIDRef`** (`<hp:run>` 속성): header.xml의 charPr(글자 속성)를 참조하는 ID.
  폰트·크기·색상·굵기 등을 결정. **run을 병합/삭제하면 이 ID가 잘못된 텍스트에 적용되어
  "서식 쏠림"이 발생한다.**
- **`<hp:linesegarray>`** (`<hp:p>`의 자식): 렌더링 캐시. 줄별 segment 좌표·높이가 박혀 있다.
  텍스트가 바뀌면 캐시가 틀려지므로 **반드시 제거**해야 한컴이 재계산한다.
  (한컴 포럼 1677 / python-hwpx `_clear_paragraph_layout_cache` 패턴)

## `<hp:t>` 인라인 자식 요소

`<hp:t>` 안에는 순수 텍스트뿐 아니라 인라인 요소가 혼재할 수 있다. 주요 종류:

- `<hp:lineBreak/>` — soft line break (Shift+Enter).
- `<hp:tab/>` 또는 `<hp:ctrl id="tab"/>` — 탭 문자.
- `<hp:markpenBegin/>` / `<hp:markpenEnd/>` — 형광펜 구간.
- `<hp:fieldBegin/>` / `<hp:fieldEnd/>` — 하이퍼링크·필드 구간.

```xml
<hp:t>안녕<hp:lineBreak/>세상</hp:t>
```

`e.text`만 읽으면 "안녕"만 얻고 뒷부분은 잃는다. 올바른 읽기/쓰기는 `e.text` + 각 자식의
`child.tail`을 순서대로 순회해야 한다. 치환 시에는 인라인 요소를 임시 sentinel 문자로
치환하고, 치환 후 다시 복원해야 원본 인라인 구조가 유지된다.

## 치환 함정: run 분할

사용자가 `{{회사명}}` 하나를 한 번에 타이핑해도, 한컴이 저장할 때 커서 이동·편집
내역에 따라 다음처럼 쪼개질 수 있다:

```xml
<hp:run charPrIDRef="0"><hp:t>{{회사</hp:t></hp:run>
<hp:run charPrIDRef="7"><hp:t>명}}</hp:t></hp:run>
```

단순 문자열 치환으로는 `{{회사명}}`을 찾지 못한다. 해법 (v0.2):

1. `<hp:p>` 단위로 모든 `<hp:t>`의 논리 텍스트를 concat.
2. target이 단일 `<hp:t>` 안에 있으면 거기서만 replace (인라인 요소 보존).
3. 여러 run에 걸치면 **run을 병합하지 않고** 각 run의 t.text를 조각 단위로 편집:
   - target 범위를 덮는 첫 run에 value 삽입.
   - 나머지 교차 run의 t.text에서 target 조각만 제거.
   - `<hp:run>` 자체와 속성(charPrIDRef)은 **보존**.
4. 교차 run들의 charPrIDRef가 다르면 경고 — 서식 쏠림 가능성.
5. 치환 완료 후 해당 `<hp:p>`의 `<hp:linesegarray>` 제거.

## 표 구조

```xml
<hp:tbl rowCnt="3" colCnt="2">
  <hp:tr>
    <hp:tc colSpan="1" rowSpan="1">
      <hp:subList>
        <hp:p><hp:run><hp:t>신청일자</hp:t></hp:run></hp:p>
      </hp:subList>
    </hp:tc>
    <hp:tc colSpan="1" rowSpan="1">
      <hp:subList>
        <hp:p><hp:run><hp:t>2026-04-18</hp:t></hp:run></hp:p>
      </hp:subList>
    </hp:tc>
  </hp:tr>
  ...
</hp:tbl>
```

- `rowCnt`/`colCnt`는 논리적 행·열 수 (병합 감안 없이 grid 기준)
- `<hp:tc colSpan rowSpan>`: 병합 셀
- 병합되면 `<hp:tc>` 자체가 없는 좌표가 생김 → 인덱스 접근 불안정
- **라벨 기반 접근 권장**: "신청일자" 셀 찾고 → 같은 행의 다음 `<hp:tc>` 주입

## 셀에 텍스트 주입 (v0.2)

기존 텍스트가 있을 때:
- 셀 안 첫 `<hp:t>`에 새 값 쓰고 나머지 `<hp:t>`는 비움. `charPrIDRef` 재확인 set.
- `<hp:linesegarray>`가 있으면 제거.

기존 텍스트가 없을 때 (빈 셀):
- `<hp:tc>` 밑에 `<hp:subList>` → `<hp:p>` → `<hp:run>` → `<hp:t>` 계층 생성.
- 속성 템플릿 우선순위:
  1. 같은 `<hp:tr>`의 왼쪽 이웃 `<hp:tc>`의 첫 p/run 속성.
  2. 같은 `<hp:tbl>`의 다른 셀 중 가장 많이 쓰인 속성.
  3. 최후: `paraPrIDRef="0"`, `charPrIDRef="0"`.
- `subList.id`는 복제하지 않음 (ID 충돌 방지).

## 참고 링크

- 한컴 기술 가이드: https://tech.hancom.com/hwpxformat/
- KS X 6101 (OWPML 표준) : https://standard.go.kr/
- 한컴 공식 HWP OWPML 다운로드: https://www.hancom.com/support/downloadCenter/hwpOwpml

# HWP5 바이너리 포맷 노트

HWP5 = **OLE Compound File** (Microsoft 구 Office 컨테이너) + 각 스트림 내부는
**raw zlib deflate**. 레코드 기반 구조.

## 주요 스트림

| 스트림 | 용도 | 압축 |
|--------|------|------|
| `FileHeader` | 버전/속성 플래그 (256B 고정) | 비압축 |
| `DocInfo` | 폰트/스타일/ID 맵핑 등 문서 전역 | zlib |
| `BodyText/Section0`, `Section1`, ... | 본문 섹션별 레코드 | zlib |
| `ViewText/SectionN` | 배포용 읽기 전용 뷰 (암호화 가능성) | zlib |
| `HwpSummaryInformation` | OLE 속성: 제목/저자 등 | MS 표준 |
| `PrvText` | 미리보기 텍스트 (UTF-16LE, 앞 몇 KB) | 비압축 |
| `PrvImage` | 미리보기 이미지 | 바이너리 |
| `Scripts/*` | 매크로 (JScript 등) | 건드리지 말 것 |
| `BinData/*` | 삽입된 바이너리 (이미지 등) | 대체로 zlib |

## FileHeader 플래그

- 총 256B
- offset 0..31: 서명 `HWP Document File\x00...`
- offset 32..35: 버전 (UINT32 LE, MMmmPPrr 형식)
- offset 36..39: **properties** (UINT32 LE)
  - bit 0: 압축 여부 (1 = 본문 zlib 압축)
  - bit 1: 암호 설정
  - bit 2: 배포용 문서
  - 등등

```python
data = ole.openstream("FileHeader").read()
props = struct.unpack_from("<I", data, 36)[0]
compressed = bool(props & 0x1)
```

## 레코드 헤더 (4B 비트필드)

```
[ Size:12b | Level:10b | TagID:10b ]
  bits 31..20   bits 19..10  bits 9..0
```

- **TagID** (10b, 0..1023): 레코드 종류
- **Level** (10b): 구조 계층 (문단 안의 런 등)
- **Size** (12b): 페이로드 바이트 수. `0xFFF`면 다음 4B LE로 실제 크기.

```python
header = struct.unpack("<I", buf[i:i+4])[0]
tag_id = header & 0x3FF
level  = (header >> 10) & 0x3FF
size   = (header >> 20) & 0xFFF
i += 4
if size == 0xFFF:
    size = struct.unpack("<I", buf[i:i+4])[0]
    i += 4
payload = buf[i:i+size]
i += size
```

## 주요 TagID

`HWPTAG_BEGIN = 0x10` 을 기준으로 증가.

| 이름 | 값 | 의미 |
|------|----|----|
| `HWPTAG_DOCUMENT_PROPERTIES` | 16 (0x10) | DocInfo 시작 |
| `HWPTAG_ID_MAPPINGS` | 17 | 폰트/스타일 ID 맵핑 |
| `HWPTAG_PARA_HEADER` | 66 (0x42) | 문단 헤더 (BodyText) |
| `HWPTAG_PARA_TEXT` | 67 (0x43) | **문단 텍스트** |
| `HWPTAG_PARA_CHAR_SHAPE` | 68 (0x44) | 문자 속성 인덱스 |
| `HWPTAG_PARA_LINE_SEG` | 69 | 라인 세그먼트 |
| `HWPTAG_CTRL_HEADER` | 71 | 컨트롤 (표, 그림 등) |
| `HWPTAG_TABLE` | 91 | 표 |

## PARA_TEXT 페이로드

- UTF-16LE 스트림
- 문자 범위 외에 **인라인 제어문자**(0x00..0x1F)가 섞임
- 제어문자 중 일부(1,2,3,4,...,0x1F 대부분)는 **8 UTF-16 unit = 16B 확장 블록**
  - 예: 테이블/그림 앵커, 필드 시작/끝 등
- 안전한 단순 추출:
  - `\n` (0x0A), `\t` (0x09) 유지
  - 나머지 0x00..0x1F는 14B skip (이미 2B 읽었으므로)
  - surrogate pair 엄밀 처리는 생략

## 압축 해제

```python
import zlib
# wbits=-15 → raw deflate (zlib 헤더 없음)
decoded = zlib.decompress(raw_stream, -15)
```

**wbits=15 (기본값)로 하면 안 됨.** `zlib.error: invalid block type` 등 실패.

## 구현 한계

- 제어문자 16B 확장은 태그별로 크기가 미세하게 다를 수 있음 (공식 스펙 PDF 참고).
- 정확한 텍스트 추출이 필요하면 **pyhwp** 같은 전문 라이브러리 사용 권장.
- 삽입 객체(수식, 차트, 이미지 캡션 등)의 텍스트는 이 단순 구현으로는 누락됨.
- 암호 설정된 문서는 뷰텍스트가 별도 암호화. 복호화 로직 필요 (본 스킬 범위 외).

## 참고

- **HWP 5.0 공식 스펙** (한컴, revision 1.3, 2018 PDF):
  https://www.hancom.com/support/downloadCenter/hwpOwpml
- pyhwp (참고 구현): https://github.com/mete0r/pyhwp
- olefile (OLE 파싱): https://pypi.org/project/olefile/

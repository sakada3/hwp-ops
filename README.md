# hwp-ops

[![License: Apache 2.0](https://img.shields.io/badge/License-Apache_2.0-blue.svg)](https://opensource.org/licenses/Apache-2.0)
[![AI-authored](https://img.shields.io/badge/code-100%25_AI--authored-ff69b4.svg)](#이-프로젝트가-조금-특이한-이유)
[![Model: Claude Opus 4.7](https://img.shields.io/badge/model-Claude_Opus_4.7-8a2be2.svg)](https://www.anthropic.com)

한글과컴퓨터 **HWP / HWPX** 문서를 읽고, 플레이스홀더를 치환하고, 표 양식을 자동으로 채우는 **Claude Code skill**. 정형 한글 양식을 자동화할 때 쓴다.

## 이 프로젝트가 조금 특이한 이유

이 저장소의 **코드는 전부 Anthropic Claude Opus 4.7이 작성**했다 (Claude Code 환경, 서브에이전트 기반 조사 → 계획 → 개발 → 검수 4단계 루프). 사람은 방향을 지시하고 결과를 검토·푸시만 했다. "AI가 얼마나 제대로 된 Claude Code skill을 end-to-end로 만들 수 있는가"를 확인하는 실험이자 실사용 도구다.

## 주요 기능

- `hwp_detect` — HWP5/HWPX 포맷 자동 판별 + 환경 프로브 (lxml / python-hwpx / 한컴오피스 COM / Java hwp2hwpx).
- `hwpx_scan` — HWPX 내부 구조·플레이스홀더(`{{키}}`)·표 스캔. **분할된 플레이스홀더와 서식 쏠림 위험(charPrIDRef 이질)**을 사전 경고.
- `hwpx_fill` — **레이아웃 보존** 치환 엔진. 주요 불변식:
  - `<hp:run>` 구조 보존, `charPrIDRef` 재지정으로 서식 유지.
  - 치환된 `<hp:p>`에서 **`<hp:linesegarray>` 렌더링 캐시 제거** (한컴이 자동 재계산) → "셀 높이 고정" 같은 고질 버그 방어.
  - `<hp:lineBreak/>`·`<hp:tab/>`·`<hp:markpen*>` 등 `<hp:t>` 내부 인라인 요소 PUA sentinel 보존.
  - 빈 셀 채울 때 이웃 셀 `charPrIDRef` template 복제 (4단계 fallback).
  - `value != value.strip()`이면 `xml:space="preserve"` 자동 부착, `\n` → `<hp:lineBreak/>` 변환.
  - 네임스페이스 2011/2024 두 세대 **동적 탐지**.
- `hwpx_rezip` — `mimetype` STORED 첫 엔트리 강제, UTF-8 no BOM / LF, ZIP CRC 자가 검증.
- `hwp5_extract` — HWP5 텍스트 추출. PrvText 미리보기 → BodyText 섹션 레코드 순회 폴백 체인.
- `hwp_to_hwpx` — HWP5 → HWPX 변환 라우터 (pyhwpx / hwp2hwpx). 환경 없으면 명시적 실패.

## 설치

Claude Code의 `~/.claude/skills/` 아래에 복사하거나 심볼릭 링크.

```bash
# 저장소 clone
git clone https://github.com/sakada3/hwp-ops.git

# Windows (관리자 CMD)
mklink /D %USERPROFILE%\.claude\skills\hwp-ops path\to\hwp-ops

# macOS / Linux
ln -s "$(pwd)/hwp-ops" ~/.claude/skills/hwp-ops
```

의존성:

```bash
pip install olefile           # 필수 (HWP5 읽기)
pip install lxml              # 강력 권장 (HWPX XML 편집 안정성)
```

## 빠른 사용 예

```bash
# 구조 스캔
python scripts/hwpx_scan.py sample.hwpx --placeholders --tables

# 플레이스홀더 치환
python scripts/hwpx_fill.py template.hwpx out.hwpx \
  --kv "회사명=V-Machina" --kv "대표자=홍길동"

# 표 라벨 자동 채우기
python scripts/hwpx_fill.py template.hwpx out.hwpx \
  --table-label "신청일자=2026-04-18"

# JSON 매핑
python scripts/hwpx_fill.py template.hwpx out.hwpx --json mapping.json
```

Claude Code 안에서는 그냥 "이 hwpx 파일 채워줘"라고 말하면 스킬이 트리거된다.

## 디렉토리 구조

```
hwp-ops/
├── SKILL.md                         # Claude가 읽는 메타데이터 + 워크플로우
├── scripts/
│   ├── hwp_detect.py
│   ├── hwpx_scan.py
│   ├── hwpx_rezip.py
│   ├── hwpx_fill.py                 # 레이아웃 보존 치환 엔진 (핵심)
│   ├── hwp5_extract.py
│   └── hwp_to_hwpx.py
└── reference/
    ├── hwpx-anatomy.md              # ZIP 구조·NS·표 구조 치트시트
    ├── hwp5-records.md              # OLE 레코드 헤더·TagID
    ├── pitfalls.md                  # v0.2 레이아웃 보존 규칙 TOP 10
    └── verification.md              # 수동 검증 체크리스트
```

## 주의 / 면책

- 본 스킬은 **한컴오피스의 비공식 도구**이며 한컴과 무관합니다 (unofficial).
- 실제 한컴오피스 렌더링 결과는 **버전·글꼴 가용성·OS**에 따라 달라질 수 있습니다. 최종 출력은 반드시 한글에서 직접 열어 검증하세요.
- **HWP5 바이너리의 쓰기·수정은 지원하지 않습니다** (안전상). HWP5는 읽기 + HWPX 변환만.
- **상업/기업 대규모 배포** 용도라면 한컴 공식 [HwpCtrl / Hancom HWP SDK](https://www.hancom.com/product/sdk/hwpSdk) 라이선스 검토를 권장합니다.
- HWPX 스펙 미공개 영역에 의한 엣지 케이스가 발생할 수 있습니다. 깨진 샘플은 이슈로 올려주시면 패치합니다.
- 본 소프트웨어는 Apache-2.0 라이선스의 **"AS IS" 조항**에 따라 어떠한 보증도 하지 않습니다.

## 기여

[CONTRIBUTING.md](CONTRIBUTING.md) 참조. 외부 코드 직접 복사는 금지 (특히 Non-Commercial·GPL·AGPL 계열 HWP 라이브러리).

## Acknowledgements

자세한 출처는 [NOTICE](NOTICE) 파일을 참조. 요약:

- **olefile** (Philippe Lagadec, BSD) — HWP5 OLE 스트림 파싱 런타임 의존.
- **Hancom HWP 5.0 공개 스펙**, **KS X 6101 OWPML** — 파일 포맷 구현 근거.
- **python-hwpx** (airmang, Non-Commercial) — `<hp:linesegarray>` 제거 필요성과 `fill_by_path` 경로 문법 아이디어를 공개 구현으로부터 학습. **코드는 복사하지 않았으며 런타임 의존도 없음**, 모든 알고리즘은 OWPML 스펙으로부터 독립 재구현.
- **hwpxlib / hwplib** (neolord0, Apache-2.0) — OWPML 요소 해석 교차 참조.
- **한컴 공식 개발자 포럼** — linesegarray 렌더링 캐시 동작의 공식 답변.

## License

Apache License 2.0 — 자세한 내용은 [LICENSE](LICENSE) 참조.

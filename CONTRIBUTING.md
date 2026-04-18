# Contributing to hwp-ops

이 프로젝트는 Anthropic Claude Opus 4.7 (Claude Code)이 조사·계획·개발·검수 4단계를 서브에이전트로 수행해 **100% AI가 작성**한 Claude Code skill 실험입니다. 사람 기여도 환영하지만 아래 규칙을 따라주세요.

## 기본 원칙

1. **외부 코드 직접 복사 금지.** 특히 아래 프로젝트의 코드는 어떤 상황에서도 이 저장소로 복붙하지 마세요:
   - [python-hwpx](https://github.com/airmang/python-hwpx) — Non-Commercial 라이선스
   - [airmang/hwpx-skill](https://github.com/airmang/hwpx-skill) — 라이선스 부재 (All Rights Reserved 간주)
   - GPL / AGPL / LGPL 계열 HWP 라이브러리 (pyhwp, H2Orestart 등) — 전염성 라이선스
2. 알고리즘·아이디어 참고는 OK. 단 NOTICE 파일의 "Inspiration and cross-reference" 섹션에 출처를 한 줄 추가해주세요.
3. 공식 HWP 5.0 스펙·KS X 6101 OWPML에 근거한 구현은 자유롭게 추가 가능.
4. 모든 새 파이썬 파일 맨 위에 `# SPDX-License-Identifier: Apache-2.0` 헤더를 포함하세요.
5. PR로 코드를 제출할 때는 Apache-2.0 하에 기여하는 것에 동의하는 것으로 간주됩니다 (Developer Certificate of Origin 관례).

## 코드 스타일

- Python 3.8+ 타깃. 선택 의존성은 `lxml`만 허용.
- 스킬의 워크플로우 규칙은 `SKILL.md`와 `reference/pitfalls.md`에 명문화되어 있으니 새 기능 추가 시 그 불변식을 위반하지 않는지 확인하세요.
- 에러/로그 메시지는 한국어 OK.

## 테스트

자동 테스트 인프라는 없습니다. 새 기능 추가 시 `reference/verification.md`의 체크리스트를 확장하고, 실제 HWPX 템플릿으로 한컴오피스에서 열어본 결과를 PR 설명에 첨부해주세요.

## 이슈 리포트

레이아웃 깨짐 버그는 **깨진 샘플 HWPX**(개인정보 제거)를 함께 첨부해주시면 재현이 쉽습니다.

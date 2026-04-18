#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# SPDX-License-Identifier: Apache-2.0
# HWP5 → HWPX 변환. pyhwpx(한컴오피스 연동) → hwp2hwpx.jar 순서로 시도.

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import textwrap
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass


NO_ENV_MSG = textwrap.dedent(
    """
    HWP → HWPX 변환 환경이 없습니다. 아래 중 하나를 준비해주세요:
      (a) Windows + 한컴오피스 설치 + `pip install pyhwpx`
      (b) JRE 11+ + hwp2hwpx.jar 다운로드 + 환경변수 HWP2HWPX_JAR 설정
          (https://github.com/neolord0/hwp2hwpx)
    또는 처음부터 HWPX로 입력을 준비해주세요.
    """
).strip()


def try_pyhwpx(in_path: Path, out_path: Path) -> bool:
    try:
        from pyhwpx import Hwp  # type: ignore
    except ImportError:
        return False

    try:
        hwp = Hwp(visible=False)
        hwp.open(str(in_path.resolve()))
        hwp.save_as(str(out_path.resolve()), format="HWPX")
        try:
            hwp.quit()
        except Exception:
            pass
        return out_path.exists()
    except Exception as e:
        print(f"pyhwpx 시도 실패: {e}", file=sys.stderr)
        return False


def try_hwp2hwpx(in_path: Path, out_path: Path) -> bool:
    jar = os.environ.get("HWP2HWPX_JAR")
    if not jar or not os.path.exists(jar):
        return False
    if not shutil.which("java"):
        print("java 실행 파일을 PATH에서 찾을 수 없습니다.", file=sys.stderr)
        return False
    try:
        subprocess.run(
            ["java", "-jar", jar, str(in_path.resolve()), str(out_path.resolve())],
            check=True,
        )
        return out_path.exists()
    except subprocess.CalledProcessError as e:
        print(f"hwp2hwpx 실행 실패: exit={e.returncode}", file=sys.stderr)
        return False


def convert(in_path: Path, out_path: Path, backend: str) -> str:
    if backend in ("auto", "pyhwpx"):
        if try_pyhwpx(in_path, out_path):
            return "pyhwpx"
        if backend == "pyhwpx":
            print("pyhwpx 백엔드 사용 실패", file=sys.stderr)
            sys.exit(1)

    if backend in ("auto", "hwp2hwpx"):
        if try_hwp2hwpx(in_path, out_path):
            return "hwp2hwpx"
        if backend == "hwp2hwpx":
            print(
                "hwp2hwpx 백엔드 사용 실패 "
                "(HWP2HWPX_JAR 환경변수 + JRE 확인)",
                file=sys.stderr,
            )
            sys.exit(1)

    print(NO_ENV_MSG, file=sys.stderr)
    sys.exit(1)


def main() -> int:
    ap = argparse.ArgumentParser(description="HWP5 → HWPX 변환")
    ap.add_argument("input", help="입력 .hwp")
    ap.add_argument("output", help="출력 .hwpx")
    ap.add_argument(
        "--backend",
        choices=["auto", "pyhwpx", "hwp2hwpx"],
        default="auto",
    )
    args = ap.parse_args()

    in_path = Path(args.input)
    out_path = Path(args.output)
    if not in_path.exists():
        print(f"입력 없음: {in_path}", file=sys.stderr)
        return 2

    used = convert(in_path, out_path, args.backend)
    print(f"변환 완료: {out_path} (backend={used})")
    return 0


if __name__ == "__main__":
    sys.exit(main())

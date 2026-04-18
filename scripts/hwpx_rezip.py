#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# SPDX-License-Identifier: Apache-2.0
# 디렉토리 → HWPX ZIP 재압축. mimetype STORED + 첫 엔트리 규칙 강제.

from __future__ import annotations

import argparse
import os
import sys
import zipfile
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass


def rezip(src_dir: Path, out_path: Path) -> None:
    mimetype_path = src_dir / "mimetype"
    if not mimetype_path.exists():
        raise FileNotFoundError(
            f"{mimetype_path} 없음 — HWPX 언패킹 디렉토리가 아닙니다"
        )

    # 기존 출력 제거 (append 방지)
    if out_path.exists():
        out_path.unlink()

    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zf:
        # 1) mimetype 먼저, STORED — 한컴 인식 조건
        zi = zipfile.ZipInfo("mimetype")
        zi.compress_type = zipfile.ZIP_STORED
        zf.writestr(zi, mimetype_path.read_bytes())

        # 2) 나머지 DEFLATED. 경로 구분자는 ZIP 규약대로 '/' 강제.
        for root, _, files in os.walk(src_dir):
            for name in sorted(files):
                full = Path(root) / name
                if full == mimetype_path:
                    continue
                rel = str(full.relative_to(src_dir)).replace(os.sep, "/")
                zf.write(full, rel, zipfile.ZIP_DEFLATED)

    _verify(out_path)


def _verify(out_path: Path) -> None:
    with zipfile.ZipFile(out_path) as zf:
        infos = zf.infolist()
        if not infos:
            out_path.unlink(missing_ok=True)
            raise RuntimeError("빈 ZIP")
        first = infos[0]
        if first.filename != "mimetype" or first.compress_type != 0:
            out_path.unlink(missing_ok=True)
            raise RuntimeError(
                f"mimetype 첫 엔트리/STORED 보장 실패: "
                f"filename={first.filename} compress={first.compress_type}"
            )
        bad = zf.testzip()
        if bad:
            out_path.unlink(missing_ok=True)
            raise RuntimeError(f"ZIP 손상: {bad}")


def main() -> int:
    ap = argparse.ArgumentParser(description="HWPX 재압축 유틸")
    ap.add_argument("src_dir", help="HWPX 언패킹 디렉토리")
    ap.add_argument("out", help="출력 .hwpx 경로")
    args = ap.parse_args()

    src = Path(args.src_dir)
    out = Path(args.out)

    if not src.is_dir():
        print(f"디렉토리 아님: {src}", file=sys.stderr)
        return 2

    try:
        rezip(src, out)
    except Exception as e:
        print(f"재압축 실패: {e}", file=sys.stderr)
        return 1

    print("OK")
    return 0


if __name__ == "__main__":
    sys.exit(main())

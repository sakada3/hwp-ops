#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# SPDX-License-Identifier: Apache-2.0
# HWP/HWPX 포맷 판별 + 변환·편집 환경 프로브. 출력은 JSON(stdout).

from __future__ import annotations

import argparse
import json
import os
import platform
import subprocess
import sys
import zipfile
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass


ZIP_SIG = b"PK\x03\x04"
OLE_SIG = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"


def detect_format(path: Path) -> tuple[str, dict]:
    with open(path, "rb") as f:
        head = f.read(8)

    if head[:4] == ZIP_SIG:
        return _inspect_zip(path)

    if head == OLE_SIG:
        return _inspect_ole(path)

    return "unknown", {"head_hex": head.hex()}


def _inspect_zip(path: Path) -> tuple[str, dict]:
    try:
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
            if "mimetype" in names:
                mt = z.read("mimetype").decode("ascii", errors="replace").strip()
                if "hwp" in mt.lower():
                    first = z.infolist()[0]
                    # STORED(0) + 첫 엔트리 mimetype이어야 한컴이 정상 인식
                    zip_ok = (
                        first.filename == "mimetype" and first.compress_type == 0
                    )
                    return "hwpx", {
                        "zip_ok": zip_ok,
                        "mimetype": mt,
                        "entry_count": len(names),
                        "first_entry": first.filename,
                        "first_compress_type": first.compress_type,
                    }
            return "zip", {"entry_count": len(names), "has_mimetype": "mimetype" in names}
    except zipfile.BadZipFile as e:
        return "zip", {"error": f"손상된 ZIP: {e}"}


def _inspect_ole(path: Path) -> tuple[str, dict]:
    try:
        import olefile  # type: ignore
    except ImportError:
        return "ole", {"error": "olefile 미설치 (pip install olefile)"}

    try:
        if not olefile.isOleFile(str(path)):
            return "ole", {"error": "OLE 시그니처 있으나 olefile이 거부"}
        ole = olefile.OleFileIO(str(path))
        streams = ole.listdir()
        stream_names = ["/".join(s) for s in streams]
        has_fileheader = "FileHeader" in stream_names
        has_bodytext = any(n.startswith("BodyText/") for n in stream_names)
        result = {
            "ole_ok": has_fileheader,
            "has_bodytext": has_bodytext,
            "streams": stream_names[:20],
            "stream_count": len(stream_names),
        }
        ole.close()
        if has_fileheader:
            return "hwp5", result
        return "ole", result
    except Exception as e:
        return "ole", {"error": str(e)}


def probe_env() -> dict:
    env = {
        "platform": platform.system(),
        "python_version": platform.python_version(),
        "has_olefile": _try_import("olefile"),
        "has_lxml": _try_import("lxml.etree"),
        "has_python_hwpx": _try_import("hwpx"),
        "has_pyhwpx": _try_import("pyhwpx"),
        "has_hancom": False,
        "has_java_hwp2hwpx": False,
    }

    if platform.system() == "Windows":
        env["has_hancom"] = _probe_hancom()

    env["has_java_hwp2hwpx"] = _probe_hwp2hwpx_jar()
    return env


def _try_import(mod: str) -> bool:
    try:
        __import__(mod)
        return True
    except Exception:
        return False


def _probe_hancom() -> bool:
    try:
        import win32com.client  # type: ignore
    except Exception:
        return False
    try:
        obj = win32com.client.Dispatch("HWPFrame.HWPObject")
        try:
            obj.Quit()
        except Exception:
            pass
        return True
    except Exception:
        return False


def _probe_hwp2hwpx_jar() -> bool:
    jar = os.environ.get("HWP2HWPX_JAR")
    if not jar or not os.path.exists(jar):
        return False
    try:
        r = subprocess.run(
            ["java", "-version"],
            capture_output=True,
            timeout=5,
        )
        return r.returncode == 0
    except Exception:
        return False


def main() -> int:
    ap = argparse.ArgumentParser(description="HWP/HWPX 포맷 판별 및 환경 프로브")
    ap.add_argument("path", help="검사할 파일 경로 (.hwp / .hwpx)")
    ap.add_argument(
        "--skip-env",
        action="store_true",
        help="환경 프로브 건너뛰기 (빠름)",
    )
    args = ap.parse_args()

    path = Path(args.path)
    if not path.exists():
        print(
            json.dumps(
                {"error": f"파일 없음: {path}"}, ensure_ascii=False, indent=2
            )
        )
        return 2

    fmt, details = detect_format(path)
    out: dict = {
        "path": str(path.resolve()),
        "format": fmt,
        "size_bytes": path.stat().st_size,
        "details": details,
    }
    if not args.skip_env:
        out["env"] = probe_env()

    print(json.dumps(out, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())

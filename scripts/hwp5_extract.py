#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# SPDX-License-Identifier: Apache-2.0
# HWP5(OLE) 텍스트 추출. PrvText → BodyText/SectionN 폴백.

from __future__ import annotations

import argparse
import json
import struct
import sys
import zlib
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass


HWPTAG_BEGIN = 0x10
HWPTAG_PARA_TEXT = HWPTAG_BEGIN + 51  # = 67


def _require_olefile():
    try:
        import olefile  # type: ignore

        return olefile
    except ImportError:
        print(
            "olefile 미설치. 설치: pip install olefile",
            file=sys.stderr,
        )
        sys.exit(2)


def extract_prvtext(ole) -> str:
    if not ole.exists("PrvText"):
        return ""
    data = ole.openstream("PrvText").read()
    # PrvText는 비압축 UTF-16LE
    try:
        return data.decode("utf-16-le", errors="replace").rstrip("\x00")
    except Exception:
        return ""


def is_compressed(ole) -> bool:
    """FileHeader의 압축 플래그 확인. offset 36 바이트 근방 bit 0."""
    if not ole.exists("FileHeader"):
        return True  # 보수적으로 압축 가정
    data = ole.openstream("FileHeader").read()
    # HWP 5.0 스펙: offset 36 (0x24) 부터 properties(UINT32)
    if len(data) < 40:
        return True
    props = struct.unpack_from("<I", data, 36)[0]
    return bool(props & 0x1)


def iter_section_streams(ole) -> list[str]:
    names = []
    for entry in ole.listdir():
        joined = "/".join(entry)
        if joined.startswith("BodyText/Section") or joined.startswith(
            "ViewText/Section"
        ):
            names.append(joined)
    # Section0, Section1, ... 순서 보장
    return sorted(names, key=_section_num)


def _section_num(name: str) -> int:
    tail = name.rsplit("Section", 1)[-1]
    try:
        return int(tail)
    except ValueError:
        return 0


def decompress_stream(raw: bytes, compressed: bool) -> bytes:
    if not compressed:
        return raw
    # HWP5 본문 스트림은 raw deflate (zlib 헤더 없음) → wbits=-15
    return zlib.decompress(raw, -15)


def iter_records(buf: bytes):
    i = 0
    n = len(buf)
    while i + 4 <= n:
        header = struct.unpack_from("<I", buf, i)[0]
        tag_id = header & 0x3FF
        level = (header >> 10) & 0x3FF
        size = (header >> 20) & 0xFFF
        i += 4
        if size == 0xFFF:
            if i + 4 > n:
                break
            size = struct.unpack_from("<I", buf, i)[0]
            i += 4
        if i + size > n:
            break
        payload = buf[i : i + size]
        i += size
        yield tag_id, level, payload


def extract_para_text(payload: bytes) -> str:
    """PARA_TEXT 페이로드에서 일반 문자만 추출.

    HWP5 PARA_TEXT는 UTF-16LE 스트림인데 일부 값이 제어문자(0x00-0x1F)로
    인라인 오브젝트를 표시한다. 정밀 파서는 pyhwp 권장. 여기선 단순화:
    - \\n(0x0A), \\t(0x09)는 유지
    - 그 외 0x00-0x1F 전체 skip
    - 확장 제어 문자는 16바이트 블록이지만 단순 skip 처리로 커버
    """
    out: list[str] = []
    i = 0
    n = len(payload)
    while i + 2 <= n:
        code = struct.unpack_from("<H", payload, i)[0]
        i += 2
        if code == 0x0A or code == 0x0D:
            out.append("\n")
            continue
        if code == 0x09:
            out.append("\t")
            continue
        if code < 0x20:
            # 확장 제어문자는 뒤에 추가 데이터(총 16바이트 = 8 WORD)가 이어진다.
            # 일부 코드(1,2,3,4,5,6,7,8,9,11,13,14,15,16,17,18,19,20,21,22,23,24,25)
            # 가 16바이트 단위 확장. 안전하게 넉넉히 skip.
            skip_codes = {
                0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x0B,
                0x0C, 0x0E, 0x0F, 0x10, 0x11, 0x12, 0x13, 0x14, 0x15,
                0x16, 0x17, 0x18, 0x19, 0x1A, 0x1B, 0x1C, 0x1D, 0x1E, 0x1F,
            }
            if code in skip_codes:
                # 이미 2B 읽었으니 14B 더 skip (총 16B)
                i += 14
            continue
        # surrogate half는 버림 (pair 처리 생략)
        if 0xD800 <= code <= 0xDFFF:
            continue
        out.append(chr(code))
    return "".join(out)


def extract_section(ole, name: str, compressed: bool) -> str:
    raw = ole.openstream(name).read()
    try:
        buf = decompress_stream(raw, compressed)
    except zlib.error:
        # 압축 플래그와 실제가 다른 경우 한 번 더 시도
        buf = raw
    parts: list[str] = []
    for tag_id, _level, payload in iter_records(buf):
        if tag_id == HWPTAG_PARA_TEXT:
            parts.append(extract_para_text(payload))
    return "\n".join(parts)


def extract(path: Path, prvtext_only: bool) -> dict:
    olefile = _require_olefile()
    if not olefile.isOleFile(str(path)):
        return {"error": "OLE 파일 아님"}
    ole = olefile.OleFileIO(str(path))
    try:
        meta: dict = {"fallback_level": 0, "sections_found": 0}
        prv = extract_prvtext(ole)
        meta["prvtext_chars"] = len(prv)

        if prvtext_only:
            meta["fallback_level"] = 1
            return {"text": prv, "meta": meta}

        compressed = is_compressed(ole)
        meta["compressed"] = compressed

        section_names = iter_section_streams(ole)
        meta["sections_found"] = len(section_names)

        section_texts: list[str] = []
        for name in section_names:
            try:
                txt = extract_section(ole, name, compressed)
                section_texts.append(txt)
            except Exception as e:
                section_texts.append(f"[섹션 추출 실패: {name} ({e})]")

        if section_texts:
            full = "\n\n---\n\n".join(
                f"--- Section {i} ---\n{t}" for i, t in enumerate(section_texts)
            )
            meta["fallback_level"] = 2
            meta["total_chars"] = sum(len(t) for t in section_texts)
            return {"text": full, "meta": meta, "prvtext": prv}

        # 섹션 없으면 PrvText로 폴백
        meta["fallback_level"] = 1
        return {"text": prv, "meta": meta}
    finally:
        ole.close()


def main() -> int:
    ap = argparse.ArgumentParser(description="HWP5 텍스트 추출")
    ap.add_argument("path", help="입력 .hwp")
    ap.add_argument(
        "--format", choices=["text", "json"], default="text", help="출력 형식"
    )
    ap.add_argument("--prvtext-only", action="store_true", help="PrvText만 추출")
    args = ap.parse_args()

    path = Path(args.path)
    if not path.exists():
        print(f"파일 없음: {path}", file=sys.stderr)
        return 2

    result = extract(path, prvtext_only=args.prvtext_only)
    if "error" in result:
        print(result["error"], file=sys.stderr)
        return 1

    if args.format == "json":
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0

    print(result["text"])
    meta = result.get("meta", {})
    print()
    print("[meta]")
    for k, v in meta.items():
        print(f"{k}: {v}")
    return 0


if __name__ == "__main__":
    sys.exit(main())

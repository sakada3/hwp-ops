#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# SPDX-License-Identifier: Apache-2.0
# HWPX 구조 스캐너 v0.2: ZIP 엔트리·플레이스홀더·서식·논리 격자 표.

from __future__ import annotations

import argparse
import re
import sys
import zipfile
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass


# 2011/2024 공통 키워드 — URI 버전 상관 없이 동작하도록 localname 기반으로 매칭
NS_KEYWORDS = {
    "hp": "paragraph",
    "hs": "section",
    "hh": "head",
    "hc": "core",
}

# 기본 fallback (둘 다 지원)
DEFAULT_NS_2011 = {
    "hp": "http://www.hancom.co.kr/hwpml/2011/paragraph",
    "hs": "http://www.hancom.co.kr/hwpml/2011/section",
    "hh": "http://www.hancom.co.kr/hwpml/2011/head",
    "hc": "http://www.hancom.co.kr/hwpml/2011/core",
}

PLACEHOLDER_RE = re.compile(r"\{\{([^{}]+?)\}\}")
PLACEHOLDER_RAW_RE = re.compile(r"\{\{([^{}<>]+?)\}\}")
HP_T_BLOCK_RE = re.compile(
    r"<(?:[a-zA-Z0-9]+:)?t(?:\s[^>]*)?>([^<]*)</(?:[a-zA-Z0-9]+:)?t>",
    re.DOTALL,
)


def _get_etree():
    try:
        from lxml import etree  # type: ignore
        return etree, True
    except ImportError:
        import xml.etree.ElementTree as etree  # type: ignore
        return etree, False


def _localname(tag: str) -> str:
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def list_entries(hwpx: Path) -> list[tuple[str, int, int]]:
    rows: list[tuple[str, int, int]] = []
    with zipfile.ZipFile(hwpx) as z:
        for info in z.infolist():
            rows.append((info.filename, info.compress_type, info.file_size))
    return rows


def print_entries(hwpx: Path) -> None:
    rows = list_entries(hwpx)
    print(f"# ZIP 엔트리 ({len(rows)}개)")
    first = rows[0] if rows else None
    if first:
        stored = "OK" if (first[0] == "mimetype" and first[1] == 0) else "경고"
        print(f"  첫 엔트리: {first[0]} (compress={first[1]}) → mimetype STORED: {stored}")
    for name, ctype, size in rows:
        print(f"  - {name}  [{'STORED' if ctype == 0 else 'DEFLATED'}, {size}B]")


def iter_section_xmls(hwpx: Path):
    with zipfile.ZipFile(hwpx) as z:
        names = sorted(
            n for n in z.namelist()
            if re.match(r"Contents/section\d+\.xml$", n)
        )
        for n in names:
            yield n, z.read(n).decode("utf-8", errors="replace")


def _section_index(entry_name: str) -> int:
    m = re.search(r"section(\d+)\.xml$", entry_name)
    return int(m.group(1)) if m else -1


def _detect_ns_version(raw: str) -> str:
    """원본 네임스페이스 세대 식별 (안내용)."""
    if "owpml.org/owpml/2024" in raw:
        return "2024"
    if "hancom.co.kr/hwpml/2011" in raw:
        return "2011"
    return "미상"


def scan_placeholders(hwpx: Path, section_filter: int | None) -> None:
    print("# 플레이스홀더 스캔")
    any_found = False
    etree, is_lxml = _get_etree()
    for name, raw in iter_section_xmls(hwpx):
        idx = _section_index(name)
        if section_filter is not None and idx != section_filter:
            continue
        ns_ver = _detect_ns_version(raw)
        print(f"\n## {name} (section {idx}, 네임스페이스 {ns_ver})")

        raw_matches = list(PLACEHOLDER_RAW_RE.finditer(raw))
        t_blocks = [m.group(1) for m in HP_T_BLOCK_RE.finditer(raw)]
        joined_t = "".join(t_blocks)

        intact_keys: set[str] = set()
        for m in PLACEHOLDER_RE.finditer(joined_t):
            for b in t_blocks:
                if m.group(0) in b:
                    intact_keys.add(m.group(1))
                    break

        # charPrIDRef 분석용: XML 파싱
        charprs_per_key: dict[str, list[str]] = {}
        try:
            root = etree.fromstring(raw.encode("utf-8"))
            charprs_per_key = _analyze_charpr(root, is_lxml)
        except Exception:
            pass

        if not raw_matches:
            split_only = [m.group(1) for m in PLACEHOLDER_RE.finditer(joined_t)]
            if not split_only:
                print("  (플레이스홀더 없음)")
                continue
            any_found = True
            for key in split_only:
                mark = "분할됨 (주의)" if key not in intact_keys else "온전"
                cpr = charprs_per_key.get(key, [])
                cpr_note = _charpr_note(cpr)
                print(f"  - {{{{{key}}}}}  [{mark}]{cpr_note}")
            continue

        any_found = True
        all_keys = sorted({m.group(1) for m in raw_matches}
                          | {m.group(1) for m in PLACEHOLDER_RE.finditer(joined_t)})
        for key in all_keys:
            if key in intact_keys:
                mark = "[온전]"
            else:
                mark = "[분할됨 (주의) — v0.2 엔진이 자동 처리]"
            cpr = charprs_per_key.get(key, [])
            cpr_note = _charpr_note(cpr)
            print(f"  - {{{{{key}}}}}  {mark}{cpr_note}")

    if not any_found:
        print("\n(전체 섹션에서 플레이스홀더를 찾지 못했습니다)")


def _charpr_note(cpr: list[str]) -> str:
    if not cpr:
        return ""
    unique = sorted(set(cpr))
    if len(unique) == 1:
        return f"  charPrIDRef={unique[0]}"
    return f"  charPrIDRef={unique} [서식 쏠림 위험]"


def _analyze_charpr(root, is_lxml: bool) -> dict[str, list[str]]:
    """각 placeholder가 걸친 run들의 charPrIDRef 수집."""
    result: dict[str, list[str]] = {}
    # 모든 <hp:p> 순회 → run 평문 연결 → {{키}} 매칭
    for p in root.iter():
        if _localname(p.tag) != "p":
            continue
        runs = [r for r in p.iter() if _localname(r.tag) == "run"]
        # run별 텍스트 구간 [start, end] + charPrIDRef
        segs: list[tuple[int, int, str]] = []
        pos = 0
        for r in runs:
            text_len = 0
            for t in r.iter():
                if _localname(t.tag) == "t":
                    if t.text:
                        text_len += len(t.text)
                    for c in t:
                        if c.tail:
                            text_len += len(c.tail)
            cpr = r.get("charPrIDRef", "?")
            segs.append((pos, pos + text_len, cpr))
            pos += text_len
        concat = "".join(
            (t.text or "") + "".join((c.tail or "") for c in t)
            for r in runs
            for t in r.iter()
            if _localname(t.tag) == "t"
        )
        for m in PLACEHOLDER_RE.finditer(concat):
            key = m.group(1)
            ts, te = m.span()
            crossing = [cpr for (s, e, cpr) in segs if s < te and e > ts]
            result.setdefault(key, []).extend(crossing)
    return result


def scan_tables(hwpx: Path, section_filter: int | None) -> None:
    etree, is_lxml = _get_etree()
    print("\n# 표 스캔 (논리 격자)")
    for name, raw in iter_section_xmls(hwpx):
        idx = _section_index(name)
        if section_filter is not None and idx != section_filter:
            continue
        print(f"\n## {name} (section {idx})")
        try:
            root = etree.fromstring(raw.encode("utf-8"))
        except Exception as e:
            print(f"  파싱 실패: {e}")
            continue
        tbls = _iter_local(root, "tbl")
        if not tbls:
            print("  (표 없음)")
            continue
        for ti, tbl in enumerate(tbls):
            row_cnt_attr = tbl.get("rowCnt") or "?"
            col_cnt_attr = tbl.get("colCnt") or "?"
            print(f"\n  표 #{ti}: rowCnt={row_cnt_attr}, colCnt={col_cnt_attr}")
            grid, anchors = _build_grid(tbl)
            if not grid:
                print("    (빈 표)")
                continue
            for ri, row in enumerate(grid):
                for ci, anchor_pos in enumerate(row):
                    if anchor_pos is None:
                        print(f"    [{ri},{ci}] (병합 공석)")
                        continue
                    ar, ac = anchor_pos
                    tc = anchors[(ar, ac)]
                    cspan = tc.get("colSpan") or "1"
                    rspan = tc.get("rowSpan") or "1"
                    txt = _cell_text(tc)
                    preview = (txt[:30] + "…") if len(txt) > 30 else txt
                    if (ar, ac) != (ri, ci):
                        print(
                            f"    [{ri},{ci}] [merged from ({ar},{ac})]"
                        )
                    else:
                        span_note = ""
                        if cspan != "1" or rspan != "1":
                            span_note = f" (colSpan={cspan},rowSpan={rspan})"
                        print(f"    [{ri},{ci}] {preview!r}{span_note}")


def _iter_local(root, local: str) -> list:
    return [e for e in root.iter() if _localname(e.tag) == local]


def _build_grid(tbl):
    """논리 격자 + 앵커 위치 맵 반환. 병합된 좌표는 (앵커r, 앵커c) 튜플."""
    rows = [e for e in tbl if _localname(e.tag) == "tr"]
    if not rows:
        return [], {}
    try:
        col_cnt = int(tbl.get("colCnt") or 0)
    except Exception:
        col_cnt = 0
    if col_cnt == 0:
        for tr in rows:
            s = 0
            for tc in tr:
                if _localname(tc.tag) != "tc":
                    continue
                try:
                    s += int(tc.get("colSpan") or 1)
                except Exception:
                    s += 1
            col_cnt = max(col_cnt, s)
    row_cnt = len(rows)
    grid: list[list] = [[None] * col_cnt for _ in range(row_cnt)]
    anchors: dict = {}
    for ri, tr in enumerate(rows):
        ci = 0
        for tc in tr:
            if _localname(tc.tag) != "tc":
                continue
            while ci < col_cnt and grid[ri][ci] is not None:
                ci += 1
            if ci >= col_cnt:
                break
            try:
                cspan = int(tc.get("colSpan") or 1)
                rspan = int(tc.get("rowSpan") or 1)
            except Exception:
                cspan = rspan = 1
            anchors[(ri, ci)] = tc
            for rr in range(ri, min(ri + rspan, row_cnt)):
                for cc in range(ci, min(ci + cspan, col_cnt)):
                    grid[rr][cc] = (ri, ci)
            ci += cspan
    return grid, anchors


def _cell_text(tc) -> str:
    parts: list[str] = []
    for e in tc.iter():
        if _localname(e.tag) == "t":
            if e.text:
                parts.append(e.text)
            for c in e:
                if c.tail:
                    parts.append(c.tail)
    return " ".join(" ".join(parts).split())


def dump_raw_xml(hwpx: Path, inner_path: str) -> None:
    with zipfile.ZipFile(hwpx) as z:
        if inner_path not in z.namelist():
            print(f"엔트리 없음: {inner_path}", file=sys.stderr)
            print("사용 가능:", file=sys.stderr)
            for n in z.namelist():
                print(f"  - {n}", file=sys.stderr)
            sys.exit(2)
        sys.stdout.write(z.read(inner_path).decode("utf-8", errors="replace"))


def main() -> int:
    ap = argparse.ArgumentParser(description="HWPX 구조 스캐너 v0.2")
    ap.add_argument("path", help="입력 .hwpx")
    ap.add_argument("--placeholders", action="store_true", help="{{키}} + charPrIDRef 탐지")
    ap.add_argument("--tables", action="store_true", help="표 논리 격자 덤프")
    ap.add_argument("--section", type=int, default=None, help="특정 섹션만")
    ap.add_argument("--raw-xml", metavar="INNER_PATH", help="ZIP 내부 XML 덤프")
    args = ap.parse_args()

    path = Path(args.path)
    if not path.exists():
        print(f"파일 없음: {path}", file=sys.stderr)
        return 2

    if args.raw_xml:
        dump_raw_xml(path, args.raw_xml)
        return 0

    print_entries(path)

    if args.placeholders:
        print()
        scan_placeholders(path, args.section)

    if args.tables:
        scan_tables(path, args.section)

    if not (args.placeholders or args.tables):
        print("\n(힌트: --placeholders / --tables 플래그로 상세 스캔)")
    return 0


if __name__ == "__main__":
    sys.exit(main())

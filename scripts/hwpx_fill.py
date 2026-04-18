#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# SPDX-License-Identifier: Apache-2.0
# HWPX 치환·표 채우기 엔진 v0.2 — 레이아웃 보존 우선.
# 핵심 불변식:
#   1) 원본의 실제 네임스페이스 URI를 탐지해 사용 (2011 vs 2024).
#   2) 텍스트를 수정한 <hp:p>에서 <hp:linesegarray> 제거 (렌더링 캐시 무효화).
#   3) <hp:run> 구조/속성 (charPrIDRef) 절대 훼손 금지.
#   4) <hp:t> 안의 인라인 자식 (lineBreak, tab, markpenBegin 등) 보존.

from __future__ import annotations

import argparse
import copy
import json
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass


# ---------- 네임스페이스 기본값 (탐지 실패 시 fallback) ----------

# 2011 네임스페이스 (대다수 실제 파일)
DEFAULT_NS_2011 = {
    "hp": "http://www.hancom.co.kr/hwpml/2011/paragraph",
    "hs": "http://www.hancom.co.kr/hwpml/2011/section",
    "hh": "http://www.hancom.co.kr/hwpml/2011/head",
    "hc": "http://www.hancom.co.kr/hwpml/2011/core",
}

# 2024 네임스페이스 (최신 한컴)
DEFAULT_NS_2024 = {
    "hp": "http://www.owpml.org/owpml/2024/paragraph",
    "hs": "http://www.owpml.org/owpml/2024/section",
    "hh": "http://www.owpml.org/owpml/2024/head",
    "hc": "http://www.owpml.org/owpml/2024/core",
}

# prefix 키워드 → URI 식별 (URI 끝부분으로 판정)
NS_KEYWORDS = {
    "hp": "paragraph",
    "hs": "section",
    "hh": "head",
    "hc": "core",
}

# 인라인 요소 보존용 sentinel (PUA 영역 — 일반 문서 텍스트와 충돌 없음)
SENTINEL_CH = "\uE000"

# 한컴이 거부하는 제어문자 (탭·줄바꿈 제외)
CTRL_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]")

XML_SPACE_ATTR = "{http://www.w3.org/XML/1998/namespace}space"


def _get_etree():
    try:
        from lxml import etree  # type: ignore
        return etree, True
    except ImportError:
        import xml.etree.ElementTree as etree  # type: ignore
        return etree, False


def _localname(tag: str) -> str:
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def _ns_of(tag: str) -> str:
    if tag.startswith("{"):
        return tag[1:].split("}", 1)[0]
    return ""


def _sanitize_value(value: str) -> str:
    """제어문자 제거 + 탭 제거 (HWPX는 <hp:ctrl id='tab'/> 사용)."""
    if not value:
        return value
    v = CTRL_RE.sub("", value)
    if "\t" in v:
        v = v.replace("\t", " ")  # 탭은 공백으로. 경고는 호출자가.
    return v


class SectionDoc:
    """섹션 하나(또는 header 외 XML 하나)의 래퍼. 자체 네임스페이스 맵 + tree 보유."""

    def __init__(self, xml_path: Path, etree_mod, is_lxml: bool):
        self.path = xml_path
        self.etree = etree_mod
        self.is_lxml = is_lxml
        self.tree = self._parse(xml_path)
        self.root = self.tree.getroot()
        # 원본 prefix 매핑 보존 (저장 후 복구용) — _detect_ns가 사용하므로 먼저 수집
        self._orig_nsmap = self._collect_orig_nsmap()
        # 원본 URI 탐지 — 하드코딩된 2011/2024를 따르지 않음
        self.ns_hp = self._detect_ns("paragraph") or DEFAULT_NS_2011["hp"]
        self.ns_hh = self._detect_ns("head") or DEFAULT_NS_2011["hh"]
        self.ns_hc = self._detect_ns("core") or DEFAULT_NS_2011["hc"]
        self.ns_hs = self._detect_ns("section") or DEFAULT_NS_2011["hs"]

    def _parse(self, xml_path: Path):
        if self.is_lxml:
            parser = self.etree.XMLParser(remove_blank_text=False, resolve_entities=False)
            return self.etree.parse(str(xml_path), parser)
        return self.etree.parse(str(xml_path))

    def _collect_orig_nsmap(self) -> dict:
        """원본 XML 헤더에서 prefix→URI 매핑 수집 (stdlib ET fallback 용도)."""
        try:
            raw = self.path.read_text(encoding="utf-8", errors="replace")
        except Exception:
            return {}
        mapping: dict = {}
        for m in re.finditer(r'xmlns:([A-Za-z0-9_-]+)="([^"]+)"', raw):
            mapping[m.group(1)] = m.group(2)
        default = re.search(r'xmlns="([^"]+)"', raw)
        if default:
            mapping[""] = default.group(1)
        return mapping

    def _detect_ns(self, keyword: str) -> str | None:
        """루트 nsmap/xmlns 선언에서 keyword('paragraph' 등)를 포함하는 URI 반환."""
        if self.is_lxml:
            nsmap = getattr(self.root, "nsmap", {}) or {}
            for uri in nsmap.values():
                if uri and keyword in uri:
                    return uri
        # stdlib ET 또는 lxml 공통: 원본 파일에서 xmlns 선언을 읽음
        for uri in self._orig_nsmap.values():
            if uri and keyword in uri:
                return uri
        # 태그 자체의 네임스페이스에서도 힌트
        if self.is_lxml:
            for el in self.root.iter():
                uri = _ns_of(el.tag)
                if uri and keyword in uri:
                    return uri
        else:
            for el in self.root.iter():
                uri = _ns_of(el.tag)
                if uri and keyword in uri:
                    return uri
        return None

    def qname(self, prefix_key: str, local: str) -> str:
        """'hp', 'p' → '{URI}p' 형태의 Clark notation."""
        uri = getattr(self, "ns_" + prefix_key)
        return "{" + uri + "}" + local

    # ---------- 쿼리 헬퍼 ----------

    def iter_local(self, root, local: str):
        for e in root.iter():
            if _localname(e.tag) == local:
                yield e

    def first_child(self, elem, local: str):
        for c in elem:
            if _localname(c.tag) == local:
                return c
        return None

    def children_local(self, elem, local: str):
        return [c for c in elem if _localname(c.tag) == local]


class HwpxFiller:
    """HWPX 언패킹 → XML 편집 → 재압축. v0.2 레이아웃 보존 엔진."""

    def __init__(self, path: str | Path):
        self.src = Path(path)
        if not self.src.exists():
            raise FileNotFoundError(self.src)
        self._tmp = Path(tempfile.mkdtemp(prefix="hwpxfill_"))
        with zipfile.ZipFile(self.src) as z:
            z.extractall(self._tmp)
        self.etree, self.is_lxml = _get_etree()
        self._dirty: set[Path] = set()
        self._sections: dict[Path, SectionDoc] = {}
        self._warnings: list[str] = []
        if not self.is_lxml:
            # stdlib ET로도 인라인 요소 보존은 가능하지만 prefix 재작성 위험.
            print(
                "경고: lxml 미설치. stdlib ET 사용 — 네임스페이스 prefix가 "
                "변경될 수 있습니다 (pip install lxml 강력 권장).",
                file=sys.stderr,
            )

    def _section_files(self, section: int | None) -> list[Path]:
        # header.xml, version.xml은 치환 대상에서 제외 (불변식 #9)
        contents = self._tmp / "Contents"
        if not contents.is_dir():
            return []
        all_secs = sorted(contents.glob("section*.xml"))
        if section is None:
            return all_secs
        target = contents / f"section{section}.xml"
        return [target] if target.exists() else []

    def _get_section(self, xml_path: Path) -> SectionDoc:
        if xml_path not in self._sections:
            self._sections[xml_path] = SectionDoc(xml_path, self.etree, self.is_lxml)
        return self._sections[xml_path]

    def _write(self, sec: SectionDoc) -> None:
        xml_path = sec.path
        if self.is_lxml:
            sec.tree.write(
                str(xml_path),
                encoding="UTF-8",
                xml_declaration=True,
                standalone=True,
            )
            data = xml_path.read_bytes().replace(b"\r\n", b"\n")
            xml_path.write_bytes(data)
        else:
            sec.tree.write(
                str(xml_path),
                encoding="utf-8",
                xml_declaration=True,
            )
            data = xml_path.read_bytes().replace(b"\r\n", b"\n")
            data = self._restore_ns_prefixes(data, sec)
            xml_path.write_bytes(data)
        self._dirty.add(xml_path)

    def _restore_ns_prefixes(self, data: bytes, sec: SectionDoc) -> bytes:
        """stdlib ET가 재작성한 ns0:/ns1: → 원본 prefix로 복구.
        xmlns 선언(`xmlns:ns0="..."`)과 태그 참조(`<ns0:tag>`, `</ns0:tag>`) 모두 치환."""
        text = data.decode("utf-8", errors="replace")
        ns_decls = re.findall(r'xmlns:(ns\d+)="([^"]+)"', text)
        # 원본 nsmap을 기준으로 URI → 원본 prefix 매핑
        uri_to_target = {uri: p for p, uri in sec._orig_nsmap.items() if p}
        for ns_prefix, uri in ns_decls:
            target = uri_to_target.get(uri)
            if target and target != ns_prefix:
                # 태그 참조: `<ns0:xxx`, `</ns0:xxx`, 속성 prefix `ns0:attr=`
                text = re.sub(rf'\b{ns_prefix}:', f'{target}:', text)
                # xmlns 선언 자체: `xmlns:ns0="URI"` → `xmlns:hs="URI"`
                text = re.sub(
                    rf'xmlns:{ns_prefix}="',
                    f'xmlns:{target}="',
                    text,
                )
        return text.encode("utf-8")

    # ---------- 인라인 직렬화 헬퍼 (핵심 — 불변식 #4) ----------

    def _serialize_t(self, t_el) -> list[tuple[str, object]]:
        """<hp:t> 내부를 (text, str) 또는 (elem, deepcopy) 파츠 리스트로.
        중요: lxml의 deepcopy는 tail을 함께 복제하므로, parts 재조합 시
        이중 삽입을 막기 위해 deepcopy의 tail을 비워야 한다."""
        parts: list[tuple[str, object]] = []
        if t_el.text:
            parts.append(("text", t_el.text))
        for child in list(t_el):
            dup = copy.deepcopy(child)
            dup.tail = None  # _deserialize_t가 tail을 따로 세팅하므로 중복 방지
            parts.append(("elem", dup))
            if child.tail:
                parts.append(("text", child.tail))
        return parts

    def _plain_and_sentinels(self, parts: list[tuple[str, object]]):
        """parts → (평문 with sentinel, sentinel 순서대로의 element 리스트)."""
        out: list[str] = []
        sentinels: list[object] = []
        for kind, val in parts:
            if kind == "text":
                out.append(val)  # type: ignore[arg-type]
            else:
                out.append(SENTINEL_CH)
                sentinels.append(val)
        return "".join(out), sentinels

    def _rebuild_parts(self, plain: str, sentinels: list[object]) -> list[tuple[str, object]]:
        parts: list[tuple[str, object]] = []
        buf: list[str] = []
        idx = 0
        for ch in plain:
            if ch == SENTINEL_CH:
                if buf:
                    parts.append(("text", "".join(buf)))
                    buf = []
                if idx < len(sentinels):
                    parts.append(("elem", sentinels[idx]))
                    idx += 1
            else:
                buf.append(ch)
        if buf:
            parts.append(("text", "".join(buf)))
        return parts

    def _deserialize_t(self, parts: list[tuple[str, object]], t_el):
        """parts를 <hp:t>에 재기록. 기존 자식·텍스트 모두 지우고 재구성."""
        for c in list(t_el):
            t_el.remove(c)
        t_el.text = None
        last_elem = None
        for kind, val in parts:
            if kind == "text":
                text_val = val  # type: ignore[assignment]
                if last_elem is None:
                    t_el.text = (t_el.text or "") + text_val  # type: ignore[operator]
                else:
                    last_elem.tail = (last_elem.tail or "") + text_val  # type: ignore[operator]
            else:
                t_el.append(val)  # type: ignore[arg-type]
                last_elem = val

    def _apply_value_to_parts(
        self, parts: list[tuple[str, object]], sec: SectionDoc, value: str
    ) -> list[tuple[str, object]]:
        """value 안의 \\n → <hp:lineBreak/>로 치환된 파츠 시퀀스로 전환."""
        v = _sanitize_value(value)
        if "\n" not in v:
            return parts + [("text", v)] if v else parts
        # \n split → lineBreak 요소 삽입
        result: list[tuple[str, object]] = list(parts)
        chunks = v.split("\n")
        for i, chunk in enumerate(chunks):
            if chunk:
                result.append(("text", chunk))
            if i < len(chunks) - 1:
                lb = self._make_element(sec, "hp", "lineBreak")
                result.append(("elem", lb))
        return result

    def _make_element(self, sec: SectionDoc, prefix_key: str, local: str, attrib: dict | None = None):
        """SectionDoc의 네임스페이스 URI로 요소 생성."""
        tag = sec.qname(prefix_key, local)
        if self.is_lxml:
            el = self.etree.Element(tag, attrib=attrib or {})
        else:
            el = self.etree.Element(tag, attrib=attrib or {})
        return el

    # ---------- 플레이스홀더 치환 (핵심 — 불변식 #3) ----------

    def replace_placeholder(
        self, key: str, value: str, section: int | None = None
    ) -> int:
        target = "{{" + key + "}}"
        total = 0
        for xml_path in self._section_files(section):
            sec = self._get_section(xml_path)
            count = self._replace_in_section(sec, target, value)
            if count:
                self._write(sec)
                total += count
        return total

    def _replace_in_section(self, sec: SectionDoc, target: str, value: str) -> int:
        count = 0
        for p in sec.iter_local(sec.root, "p"):
            # p 내부의 모든 <hp:t> 수집
            t_elements = list(sec.iter_local(p, "t"))
            if not t_elements:
                continue
            # 1차: 단일 <hp:t>에 target 완전 포함 → 인라인 보존 치환
            replaced_single = False
            total_in_p = 0
            for t_el in t_elements:
                parts = self._serialize_t(t_el)
                plain, sentinels = self._plain_and_sentinels(parts)
                if target in plain:
                    occurrences = plain.count(target)
                    total_in_p += occurrences
                    new_plain = self._splice_value(plain, target, value)
                    new_parts = self._rebuild_parts(new_plain, sentinels)
                    # value에 \n 있으면 sec-ns의 lineBreak로 분할 삽입
                    new_parts = self._convert_newlines_in_parts(new_parts, sec)
                    self._deserialize_t(new_parts, t_el)
                    self._attach_xml_space(t_el, value)
                    replaced_single = True
            if replaced_single:
                # 같은 run의 charPrIDRef 재지정 방어 (불변식 #8)
                for t_el in t_elements:
                    run = self._parent_run(p, t_el)
                    if run is not None and "charPrIDRef" in run.attrib:
                        run.set("charPrIDRef", run.attrib["charPrIDRef"])
                self._clear_layout_cache(p, sec)
                count += total_in_p
                continue

            # 2차: run 경계에 걸친 case → multi-run splice
            spanning = self._splice_across_runs(p, sec, target, value, t_elements)
            if spanning > 0:
                self._clear_layout_cache(p, sec)
                count += spanning
        return count

    def _splice_value(self, plain: str, target: str, value: str) -> str:
        # value에 '\n'이 있으면 1단계에서는 표식 문자로 두고 파츠 빌드 때 변환
        sanitized = _sanitize_value(value)
        return plain.replace(target, sanitized)

    def _convert_newlines_in_parts(
        self, parts: list[tuple[str, object]], sec: SectionDoc
    ) -> list[tuple[str, object]]:
        """parts 안의 text 조각에서 \\n을 <hp:lineBreak/>로 분할."""
        out: list[tuple[str, object]] = []
        for kind, val in parts:
            if kind == "text" and isinstance(val, str) and "\n" in val:
                chunks = val.split("\n")
                for i, chunk in enumerate(chunks):
                    if chunk:
                        out.append(("text", chunk))
                    if i < len(chunks) - 1:
                        out.append(("elem", self._make_element(sec, "hp", "lineBreak")))
            else:
                out.append((kind, val))
        return out

    def _attach_xml_space(self, t_el, value: str) -> None:
        """앞뒤 공백 보존 필요시 xml:space='preserve' 부착 (불변식 #5)."""
        if value != value.strip():
            t_el.set(XML_SPACE_ATTR, "preserve")

    def _parent_run(self, p, t_el):
        """<hp:t>의 부모 <hp:run>을 찾아 반환 (p 하위 트리에서)."""
        for run in p.iter():
            if _localname(run.tag) != "run":
                continue
            for c in run.iter():
                if c is t_el:
                    return run
        return None

    def _clear_layout_cache(self, p, sec: SectionDoc) -> None:
        """텍스트 바뀐 <hp:p>에서 <hp:linesegarray> 제거 (불변식 #2).
        한컴 포럼 1677: linesegarray는 렌더링 캐시로, 텍스트 변경 시 재계산 필요."""
        for child in list(p):
            if _localname(child.tag).lower() == "linesegarray":
                p.remove(child)

    def _splice_across_runs(
        self, p, sec: SectionDoc, target: str, value: str, t_elements
    ) -> int:
        """run 경계에 target이 걸친 경우. run 구조 유지, <hp:t>.text만 조정.
        여러 occurrence를 처리한다."""
        # 각 t_el의 평문·sentinels·기준 offset 수집
        rec: list[dict] = []
        cursor = 0
        for t_el in t_elements:
            parts = self._serialize_t(t_el)
            plain, sentinels = self._plain_and_sentinels(parts)
            rec.append(
                {
                    "t": t_el,
                    "parts": parts,
                    "sentinels": sentinels,
                    "plain": plain,
                    "start": cursor,
                    "end": cursor + len(plain),
                }
            )
            cursor += len(plain)
        concat = "".join(r["plain"] for r in rec)
        if target not in concat:
            return 0
        replaced = 0
        sanitized_value = _sanitize_value(value)
        # 반복 치환 (겹침 방지: target 길이로 전진)
        search_from = 0
        edits: list[tuple[int, int]] = []  # target의 (start, end) 목록
        while True:
            idx = concat.find(target, search_from)
            if idx < 0:
                break
            edits.append((idx, idx + len(target)))
            search_from = idx + len(target)
        if not edits:
            return 0
        # run charPrIDRef 일관성 체크 — 서식 쏠림 경고
        charprs = set()
        for ts, te in edits:
            for r in rec:
                if r["start"] < te and r["end"] > ts:
                    run = self._parent_run(p, r["t"])
                    if run is not None and "charPrIDRef" in run.attrib:
                        charprs.add(run.attrib["charPrIDRef"])
        if len(charprs) > 1:
            self._warnings.append(
                f"서식 쏠림 가능 — 대상 '{target}'이(가) 서로 다른 charPrIDRef "
                f"({sorted(charprs)})를 가진 run에 걸쳐 있습니다. "
                "템플릿에서 {{키}}를 동일 서식으로 다시 입력하는 것을 권장합니다."
            )
        # 각 record별로 target 범위 제거 / 첫 교차 record엔 value 삽입
        # 뒤에서부터 거꾸로 적용해야 offset 꼬이지 않음
        new_plains: dict[int, str] = {i: r["plain"] for i, r in enumerate(rec)}
        for ts, te in reversed(edits):
            # 이 target에 교차하는 rec들의 인덱스 모음
            crossing = [i for i, r in enumerate(rec) if r["start"] < te and r["end"] > ts]
            if not crossing:
                continue
            # value는 첫 교차 record에만 주입
            first_i = crossing[0]
            for i in crossing:
                r = rec[i]
                lstart = max(0, ts - r["start"])
                lend = min(len(r["plain"]), te - r["start"])
                if i == first_i:
                    new_plains[i] = (
                        new_plains[i][:lstart]
                        + sanitized_value
                        + new_plains[i][lend:]
                    )
                else:
                    new_plains[i] = new_plains[i][:lstart] + new_plains[i][lend:]
            replaced += 1
        # 재구성
        for i, r in enumerate(rec):
            if new_plains[i] == r["plain"]:
                continue
            new_parts = self._rebuild_parts(new_plains[i], r["sentinels"])
            new_parts = self._convert_newlines_in_parts(new_parts, sec)
            self._deserialize_t(new_parts, r["t"])
            self._attach_xml_space(r["t"], value)
            run = self._parent_run(p, r["t"])
            if run is not None and "charPrIDRef" in run.attrib:
                run.set("charPrIDRef", run.attrib["charPrIDRef"])
        return replaced

    # ---------- 표 라벨 기반 채우기 (불변식 #7) ----------

    def fill_table_by_label(
        self,
        label: str,
        value: str,
        section: int | None = None,
        occurrence: int | None = None,
    ) -> bool:
        """라벨로 찾아 방향 path 따라 이동해 값 주입."""
        return self.fill_by_path(label, value, section=section, occurrence=occurrence)

    # NOTE: API naming (fill_by_label/fill_by_path) inspired by python-hwpx.
    # 구현은 독립적으로 작성 — OWPML 스펙과 자체 파서로 처리. 코드는 복사하지 않음.
    def fill_by_path(
        self,
        label_path: str,
        value: str,
        section: int | None = None,
        occurrence: int | None = None,
    ) -> bool:
        """'성명 > right > down' 형식의 경로 지원. 기본은 label 후 right."""
        parts = [p.strip() for p in label_path.split(">")]
        label = parts[0]
        directions = parts[1:] if len(parts) > 1 else ["right"]
        label_norm = self._normalize_label(label)

        matched: list[tuple[SectionDoc, object, int, int, list[list[object]]]] = []
        for xml_path in self._section_files(section):
            sec = self._get_section(xml_path)
            for tbl in sec.iter_local(sec.root, "tbl"):
                grid = self._build_grid(tbl, sec)
                if not grid:
                    continue
                for r, row in enumerate(grid):
                    for c, cell in enumerate(row):
                        if cell is None:
                            continue
                        if self._normalize_label(self._cell_text(cell)) == label_norm:
                            matched.append((sec, tbl, r, c, grid))

        if not matched:
            return False
        if occurrence is None and len(matched) > 1:
            self._warnings.append(
                f"라벨 '{label}' 다중 매칭 ({len(matched)}건) — "
                "occurrence를 지정하지 않아 중단. mapping.json에 occurrence 추가 필요."
            )
            return False
        idx = occurrence or 0
        if idx >= len(matched):
            return False
        sec, _tbl, r, c, grid = matched[idx]
        # 방향 이동
        for d in directions:
            r, c = self._move(grid, r, c, d)
            if r < 0:
                return False
        target_cell = grid[r][c]
        if target_cell is None:
            return False
        self._set_cell_text(target_cell, value, sec)
        self._write(sec)
        return True

    def _normalize_label(self, text: str) -> str:
        t = re.sub(r"\s+", " ", text).strip()
        # 끝의 콜론 제거 (한글·영문)
        return re.sub(r"[:\uFF1A]+$", "", t).strip()

    def _build_grid(self, tbl, sec: SectionDoc) -> list[list[object]]:
        """병합 셀을 논리 격자로 펼침. 앵커 셀은 좌상단에만 남고 나머지는 None."""
        rows = list(sec.children_local(tbl, "tr"))
        if not rows:
            return []
        try:
            col_cnt = int(tbl.get("colCnt") or 0)
        except Exception:
            col_cnt = 0
        if col_cnt == 0:
            for tr in rows:
                span_sum = 0
                for tc in sec.children_local(tr, "tc"):
                    try:
                        span_sum += int(tc.get("colSpan") or 1)
                    except Exception:
                        span_sum += 1
                if span_sum > col_cnt:
                    col_cnt = span_sum
        row_cnt = len(rows)
        grid: list[list[object]] = [[None] * col_cnt for _ in range(row_cnt)]
        for ri, tr in enumerate(rows):
            ci = 0
            for tc in sec.children_local(tr, "tc"):
                while ci < col_cnt and grid[ri][ci] is not None:
                    ci += 1
                if ci >= col_cnt:
                    break
                try:
                    cspan = int(tc.get("colSpan") or 1)
                    rspan = int(tc.get("rowSpan") or 1)
                except Exception:
                    cspan = rspan = 1
                grid[ri][ci] = tc
                for rr in range(ri, min(ri + rspan, row_cnt)):
                    for cc in range(ci, min(ci + cspan, col_cnt)):
                        if rr == ri and cc == ci:
                            continue
                        grid[rr][cc] = None  # 병합된 공석
                ci += cspan
        return grid

    def _move(self, grid, r: int, c: int, direction: str) -> tuple[int, int]:
        dr, dc = {"right": (0, 1), "down": (1, 0), "left": (0, -1), "up": (-1, 0)}.get(
            direction, (0, 1)
        )
        nr, nc = r + dr, c + dc
        # 병합된 빈 좌표면 한 번 더 같은 방향
        while 0 <= nr < len(grid) and 0 <= nc < len(grid[0]) and grid[nr][nc] is None:
            nr += dr
            nc += dc
        if 0 <= nr < len(grid) and 0 <= nc < len(grid[0]):
            return nr, nc
        return -1, -1

    def _cell_text(self, tc) -> str:
        parts: list[str] = []
        for e in tc.iter():
            if _localname(e.tag) == "t":
                if e.text:
                    parts.append(e.text)
                for c in e:
                    if c.tail:
                        parts.append(c.tail)
        return " ".join(" ".join(parts).split())

    # ---------- 셀 setter (불변식 #6) ----------

    def fill_table_by_index(
        self,
        table_idx: int,
        row: int,
        col: int,
        value: str,
        section: int = 0,
    ) -> bool:
        paths = self._section_files(section)
        if not paths:
            return False
        sec = self._get_section(paths[0])
        tbls = list(sec.iter_local(sec.root, "tbl"))
        if table_idx >= len(tbls):
            return False
        tbl = tbls[table_idx]
        grid = self._build_grid(tbl, sec)
        if row >= len(grid) or col >= (len(grid[0]) if grid else 0):
            return False
        cell = grid[row][col]
        if cell is None:
            return False
        self._set_cell_text(cell, value, sec)
        self._write(sec)
        return True

    def _set_cell_text(self, tc, value: str, sec: SectionDoc) -> None:
        """셀의 <hp:t>에 value 주입. 인라인 요소 보존, linesegarray 제거,
        빈 셀이면 주변 셀 스타일 템플릿으로 초기화 (p/run 속성 복제)."""
        # 기존 t 탐색
        sublist = sec.first_child(tc, "subList")
        p = None
        if sublist is not None:
            p = sec.first_child(sublist, "p")
        if p is not None:
            t_elements = list(sec.iter_local(p, "t"))
            if t_elements:
                first_t = t_elements[0]
                first_run = self._parent_run(p, first_t)
                # 첫 t에 value, 인라인 보존. 나머지 t는 text만 비움 (구조 유지).
                parts = self._serialize_t(first_t)
                # 기존 텍스트를 완전히 덮어씀 — 인라인 요소(tab 등)는 첫 t에서 제거
                # 하지만 흔한 시나리오에서는 셀은 단일 텍스트라 문제 없음.
                new_parts: list[tuple[str, object]] = []
                new_parts = self._apply_value_to_parts(new_parts, sec, value)
                self._deserialize_t(new_parts, first_t)
                self._attach_xml_space(first_t, value)
                if first_run is not None and "charPrIDRef" in first_run.attrib:
                    first_run.set("charPrIDRef", first_run.attrib["charPrIDRef"])
                for extra in t_elements[1:]:
                    # 다른 run의 t는 비우고 인라인도 제거 (텍스트 중복 방지)
                    for c in list(extra):
                        extra.remove(c)
                    extra.text = ""
                self._clear_layout_cache(p, sec)
                return
        # 빈 셀 — 주변 스타일 템플릿 찾아 구조 생성
        tpl_p_attr, tpl_run_attr, tpl_sub_attr = self._find_style_template(tc, sec)
        E = self.etree
        if sublist is None:
            sublist = E.SubElement(tc, sec.qname("hp", "subList"), tpl_sub_attr)
        p = E.SubElement(sublist, sec.qname("hp", "p"), tpl_p_attr)
        run = E.SubElement(p, sec.qname("hp", "run"), tpl_run_attr)
        t = E.SubElement(run, sec.qname("hp", "t"))
        new_parts: list[tuple[str, object]] = self._apply_value_to_parts([], sec, value)
        self._deserialize_t(new_parts, t)
        self._attach_xml_space(t, value)
        if "charPrIDRef" in run.attrib:
            run.set("charPrIDRef", run.attrib["charPrIDRef"])
        self._clear_layout_cache(p, sec)

    def _find_style_template(self, tc, sec: SectionDoc) -> tuple[dict, dict, dict]:
        """빈 셀 채우기용 p/run/subList 속성 템플릿. 우선순위:
        1) 같은 tc의 기존 첫 p·run 속성 (이미 있음 — 이 분기는 도달하지 않지만 안전용)
        2) 같은 tr의 왼쪽 이웃 tc의 첫 run/p 속성
        3) 같은 tbl의 다른 tc 중 가장 많이 쓰인 속성
        4) 기본값 paraPrIDRef='0', charPrIDRef='0'."""
        # 2) 같은 tr의 왼쪽 이웃
        tr = self._find_parent(sec.root, tc, "tr")
        tbl = self._find_parent(sec.root, tc, "tbl") if tr is not None else None
        if tr is not None:
            cells = sec.children_local(tr, "tc")
            try:
                pos = cells.index(tc)
            except ValueError:
                pos = -1
            for neighbor in reversed(cells[:pos]) if pos >= 0 else []:
                p_attr, run_attr, sub_attr = self._extract_cell_style(neighbor, sec)
                if p_attr or run_attr:
                    return p_attr, run_attr, sub_attr
        # 3) 같은 tbl의 다른 셀
        if tbl is not None:
            counter_p: dict = {}
            counter_r: dict = {}
            sub_sample: dict = {}
            for other in sec.iter_local(tbl, "tc"):
                if other is tc:
                    continue
                p_attr, run_attr, sub_attr = self._extract_cell_style(other, sec)
                key_p = tuple(sorted(p_attr.items())) if p_attr else None
                key_r = tuple(sorted(run_attr.items())) if run_attr else None
                if key_p is not None:
                    counter_p[key_p] = counter_p.get(key_p, 0) + 1
                if key_r is not None:
                    counter_r[key_r] = counter_r.get(key_r, 0) + 1
                if sub_attr and not sub_sample:
                    sub_sample = sub_attr
            best_p = max(counter_p.items(), key=lambda kv: kv[1])[0] if counter_p else None
            best_r = max(counter_r.items(), key=lambda kv: kv[1])[0] if counter_r else None
            if best_p is not None or best_r is not None:
                return (
                    dict(best_p) if best_p else {"paraPrIDRef": "0"},
                    dict(best_r) if best_r else {"charPrIDRef": "0"},
                    dict(sub_sample) if sub_sample else {},
                )
        # 4) 최후의 기본값
        return {"paraPrIDRef": "0"}, {"charPrIDRef": "0"}, {}

    def _extract_cell_style(self, tc, sec: SectionDoc) -> tuple[dict, dict, dict]:
        sub = sec.first_child(tc, "subList")
        if sub is None:
            return {}, {}, {}
        sub_attr = dict(sub.attrib)
        # 동적 id 속성은 빼고 복제 (충돌 방지) — id는 sec마다 고유해야 함
        sub_attr.pop("id", None)
        p = sec.first_child(sub, "p")
        if p is None:
            return {}, {}, sub_attr
        p_attr = dict(p.attrib)
        p_attr.pop("id", None)
        run = sec.first_child(p, "run")
        run_attr = dict(run.attrib) if run is not None else {}
        return p_attr, run_attr, sub_attr

    def _find_parent(self, root, target, local: str):
        for el in root.iter():
            if _localname(el.tag) == local:
                for c in el.iter():
                    if c is target:
                        return el
        return None

    # ---------- 저장 ----------

    def save(self, out_path: str | Path) -> None:
        out = Path(out_path)
        from hwpx_rezip import rezip  # 동일 디렉토리 스크립트
        rezip(self._tmp, out)

    def cleanup(self) -> None:
        shutil.rmtree(self._tmp, ignore_errors=True)

    def warnings(self) -> list[str]:
        return list(self._warnings)


# 동일 디렉토리 스크립트 import 지원
sys.path.insert(0, str(Path(__file__).parent))


def _load_mapping(path: Path) -> dict:
    data = json.loads(path.read_text(encoding="utf-8"))
    data.setdefault("placeholders", {})
    data.setdefault("tables", [])
    return data


def _parse_kv(items: list[str]) -> dict[str, str]:
    out: dict[str, str] = {}
    for it in items or []:
        if "=" not in it:
            raise ValueError(f"--kv 형식 오류 (KEY=VALUE 필요): {it}")
        k, v = it.split("=", 1)
        out[k.strip()] = v
    return out


def _parse_table_kv(items: list[str]) -> list[dict]:
    out: list[dict] = []
    for it in items or []:
        if "=" not in it:
            raise ValueError(f"--table-label 형식 오류: {it}")
        label, value = it.split("=", 1)
        out.append({"label": label.strip(), "value": value})
    return out


TODO_MARK_RE = re.compile(r"\[TODO\]")


def main() -> int:
    ap = argparse.ArgumentParser(description="HWPX 치환·표 채우기 v0.2 (레이아웃 보존)")
    ap.add_argument("input", help="입력 .hwpx")
    ap.add_argument("output", help="출력 .hwpx")
    ap.add_argument("--json", metavar="MAP", help="mapping.json 경로")
    ap.add_argument("--kv", action="append", default=[], help="KEY=VALUE 반복")
    ap.add_argument(
        "--table-label",
        action="append",
        default=[],
        help="라벨=값 반복 (표 채우기)",
    )
    ap.add_argument("--section", type=int, default=None)
    args = ap.parse_args()

    in_path = Path(args.input)
    out_path = Path(args.output)
    if not in_path.exists():
        print(f"입력 없음: {in_path}", file=sys.stderr)
        return 2

    placeholders: dict[str, str] = {}
    tables: list[dict] = []
    if args.json:
        mp = _load_mapping(Path(args.json))
        placeholders.update(mp["placeholders"])
        tables.extend(mp["tables"])
    try:
        placeholders.update(_parse_kv(args.kv))
        tables.extend(_parse_table_kv(args.table_label))
    except ValueError as e:
        print(str(e), file=sys.stderr)
        return 2

    if not placeholders and not tables:
        print(
            "치환 대상이 없습니다 (--json / --kv / --table-label 중 하나 필요)",
            file=sys.stderr,
        )
        return 2

    filler = HwpxFiller(in_path)
    try:
        total_subs = 0
        failed_keys: list[str] = []
        todo_keys: list[str] = []
        for key, value in placeholders.items():
            if isinstance(value, str) and TODO_MARK_RE.search(value):
                todo_keys.append(key)
                continue
            n = filler.replace_placeholder(key, str(value), args.section)
            if n == 0:
                failed_keys.append(key)
            total_subs += n
        table_ok: list[str] = []
        table_fail: list[str] = []
        for t in tables:
            value = t.get("value", "")
            if isinstance(value, str) and TODO_MARK_RE.search(value):
                todo_keys.append(f"table:{t.get('label', '?')}")
                continue
            occ_raw = t.get("occurrence", None)
            occ = int(occ_raw) if occ_raw is not None else None
            ok = filler.fill_by_path(
                t["label"],
                str(value),
                section=t.get("section", args.section),
                occurrence=occ,
            )
            (table_ok if ok else table_fail).append(t["label"])
        if total_subs == 0 and not table_ok and not todo_keys:
            print(
                "치환 0건 — 플레이스홀더/라벨을 찾지 못했습니다. "
                "hwpx_scan.py로 실제 구조 확인 필요.",
                file=sys.stderr,
            )
            filler.cleanup()
            return 1
        filler.save(out_path)
        print(f"치환 완료: {out_path}")
        print(f"  플레이스홀더 치환: {total_subs}건")
        if failed_keys:
            print(f"  매치 실패: {failed_keys}")
        if table_ok:
            print(f"  표 라벨 성공: {table_ok}")
        if table_fail:
            print(f"  표 라벨 실패: {table_fail}")
        if todo_keys:
            print(f"  [TODO] 보존: {todo_keys}")
        for w in filler.warnings():
            print(f"  경고: {w}", file=sys.stderr)
        return 0
    finally:
        filler.cleanup()


if __name__ == "__main__":
    sys.exit(main())

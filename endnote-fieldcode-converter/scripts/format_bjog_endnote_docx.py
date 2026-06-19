#!/usr/bin/env python3
"""Format an EndNote-fielded DOCX for BJOG visible citation style.

Preserves Word/EndNote field-code structure while changing only visible display text:
- in-text citation display: [1,2] -> superscript 1,2
- reference-list display: [1] -> 1.
- punctuation immediately after citation fields is moved before the superscript field.

Usage:
  python3 format_bjog_endnote_docx.py -i input_EndNote.docx -o output_BJOG.docx
"""
from __future__ import annotations

import argparse
import copy
import re
import zipfile
from pathlib import Path
from lxml import etree

NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_XML = "http://www.w3.org/XML/1998/namespace"
W = "{%s}" % NS_W
XML = "{%s}" % NS_XML


def get_or_make_rpr(r):
    rpr = r.find(W + "rPr")
    if rpr is None:
        rpr = etree.Element(W + "rPr")
        r.insert(0, rpr)
    return rpr


def set_superscript(r):
    rpr = get_or_make_rpr(r)
    for va in list(rpr.findall(W + "vertAlign")):
        rpr.remove(va)
    etree.SubElement(rpr, W + "vertAlign", {W + "val": "superscript"})


def clear_superscript(r):
    rpr = r.find(W + "rPr")
    if rpr is not None:
        for va in list(rpr.findall(W + "vertAlign")):
            rpr.remove(va)


def make_text_run(text: str, template_run=None):
    r = etree.Element(W + "r")
    if template_run is not None:
        rpr = template_run.find(W + "rPr")
        if rpr is not None:
            r.append(copy.deepcopy(rpr))
    t = etree.SubElement(r, W + "t")
    if text.startswith(" ") or text.endswith(" "):
        t.set(XML + "space", "preserve")
    t.text = text
    return r


def bjog_cite_text(s: str) -> str:
    s = (s or "").strip()
    if s.startswith("[") and s.endswith("]"):
        s = s[1:-1]
    return re.sub(r"\s+", "", s)


def ref_text(s: str) -> str:
    return re.sub(r"^\s*\[(\d+)\]\s+", r"\1. ", s or "")


def first_text_el(r):
    ts = r.findall(".//" + W + "t")
    return ts[0] if ts else None


def last_text_el(r):
    ts = r.findall(".//" + W + "t")
    return ts[-1] if ts else None


def format_docx(input_path: Path, output_path: Path) -> dict:
    with zipfile.ZipFile(input_path, "r") as zin:
        files = {n: zin.read(n) for n in zin.namelist()}
    root = etree.fromstring(
        files["word/document.xml"],
        etree.XMLParser(remove_blank_text=False, recover=True),
    )
    body = root.find(W + "body")

    cite_fields = 0
    cite_display_runs = 0
    reflist_fields = 0
    ref_entries_changed = 0
    instr_display_changed = 0
    in_field = False
    instr = ""
    field_kind = None
    after_sep = False

    # Format visible citation/reference display text while preserving fields.
    for elem in body.iter():
        if elem.tag != W + "r":
            continue
        fld = elem.find(W + "fldChar")
        if fld is not None:
            typ = fld.get(W + "fldCharType")
            if typ == "begin":
                in_field = True
                instr = ""
                field_kind = None
                after_sep = False
            elif typ == "separate" and in_field:
                field_kind = "cite" if "ADDIN EN.CITE" in instr else ("reflist" if "ADDIN EN.REFLIST" in instr else None)
                if field_kind == "cite":
                    cite_fields += 1
                if field_kind == "reflist":
                    reflist_fields += 1
                after_sep = True
            elif typ == "end" and in_field:
                in_field = False
                instr = ""
                field_kind = None
                after_sep = False
            continue

        instr_el = elem.find(W + "instrText")
        if instr_el is not None and in_field:
            if instr_el.text:
                old = instr_el.text
                new = re.sub(r"<DisplayText>\[([0-9,]+)\]</DisplayText>", r"<DisplayText>\1</DisplayText>", old)
                if new != old:
                    instr_display_changed += old.count("<DisplayText>[")
                    instr_el.text = new
                instr += instr_el.text or ""
            continue

        if after_sep and field_kind == "cite":
            for t in elem.findall(".//" + W + "t"):
                if t.text:
                    t.text = bjog_cite_text(t.text)
                    set_superscript(elem)
                    cite_display_runs += 1
        elif after_sep and field_kind == "reflist":
            for t in elem.findall(".//" + W + "t"):
                if t.text:
                    new = ref_text(t.text)
                    if new != t.text:
                        ref_entries_changed += 1
                        t.text = new
                    clear_superscript(elem)

    # Move periods/commas after citation fields before the superscript field.
    punct_moved = 0
    for p in list(body.iterchildren(W + "p")):
        children = list(p)
        ranges = []
        i = 0
        while i < len(children):
            ch = children[i]
            fld = ch.find(W + "fldChar") if ch.tag == W + "r" else None
            if fld is not None and fld.get(W + "fldCharType") == "begin":
                instr = ""
                is_cite = False
                j = i + 1
                while j < len(children):
                    r = children[j]
                    if r.tag == W + "r":
                        ins = r.find(W + "instrText")
                        if ins is not None and ins.text:
                            instr += ins.text
                        f = r.find(W + "fldChar")
                        if f is not None:
                            if f.get(W + "fldCharType") == "separate":
                                is_cite = "ADDIN EN.CITE" in instr
                            elif f.get(W + "fldCharType") == "end":
                                if is_cite:
                                    ranges.append((i, j))
                                break
                    j += 1
                i = j
            i += 1

        for start, end in reversed(ranges):
            children = list(p)
            if end + 1 >= len(children):
                continue
            next_run = children[end + 1]
            if next_run.tag != W + "r":
                continue
            tnext = first_text_el(next_run)
            if tnext is None or not tnext.text or tnext.text[0] not in ".,":
                continue
            punct = tnext.text[0]
            tnext.text = tnext.text[1:]
            template = None
            for k in range(start - 1, -1, -1):
                if children[k].tag == W + "r":
                    tprev = last_text_el(children[k])
                    if tprev is not None and tprev.text is not None:
                        tprev.text = tprev.text.rstrip()
                        template = children[k]
                        break
            p.insert(start, make_text_run(punct, template))
            punct_moved += 1

    files["word/document.xml"] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for n, d in files.items():
            zout.writestr(n, d)

    with zipfile.ZipFile(output_path) as z:
        bad = z.testzip()
        xml = z.read("word/document.xml")
    begin = len(re.findall(rb"fldCharType=['\"]begin['\"]", xml))
    separate = len(re.findall(rb"fldCharType=['\"]separate['\"]", xml))
    end = len(re.findall(rb"fldCharType=['\"]end['\"]", xml))
    return {
        "output": str(output_path),
        "zip_test_bad_file": bad,
        "EN.CITE": xml.count(b"ADDIN EN.CITE"),
        "EN.REFLIST": xml.count(b"ADDIN EN.REFLIST"),
        "field_markers": (begin, separate, end),
        "balanced": begin == separate == end,
        "cite_fields_seen": cite_fields,
        "cite_display_runs_superscripted": cite_display_runs,
        "reflist_fields_seen": reflist_fields,
        "reference_entries_changed": ref_entries_changed,
        "instr_DisplayText_changed": instr_display_changed,
        "punctuation_moved_before_superscripts": punct_moved,
        "size": output_path.stat().st_size,
    }


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("-i", "--input", required=True)
    ap.add_argument("-o", "--output", required=True)
    args = ap.parse_args()
    stats = format_docx(Path(args.input), Path(args.output))
    for k, v in stats.items():
        print(f"{k}: {v}")
    return 0 if stats["zip_test_bad_file"] is None and stats["balanced"] else 1


if __name__ == "__main__":
    raise SystemExit(main())

#!/usr/bin/env python3
"""
EndNote Field Code Converter

Convert a DOCX manuscript with plain-text numbered citations into a DOCX
containing EndNote-compatible ADDIN EN.CITE field codes and an ADDIN EN.REFLIST
bibliography field.
"""
from __future__ import annotations

import argparse
import copy
import json
import os
import re
import sys
import time
import urllib.parse
import urllib.request
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

try:
    from lxml import etree
except ImportError as exc:  # pragma: no cover
    raise SystemExit("Missing dependency: lxml. Install with: python3 -m pip install lxml") from exc

NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_XML = "http://www.w3.org/XML/1998/namespace"
W = "{%s}" % NS_W

STOP_HEADINGS = {
    "acknowledgements",
    "acknowledgments",
    "author contributions",
    "funding",
    "conflicts of interest",
    "conflict of interest",
    "competing interests",
    "declarations",
    "supplementary material",
    "supplemental material",
    "tables",
    "figures",
    "figure legends",
    "disclosures",
}

REF_HEADING_RE = re.compile(r"^(references|reference|bibliography)\s*$", re.I)
REF_START_RE = re.compile(r"^\s*\[\s*(\d+)\s*\]\s*(.*)$", re.S)
CITATION_RE = re.compile(r"\[(\s*\d+\s*(?:(?:,|;|\-|\u2013|\u2014)\s*\d+\s*)*)\]")
YEAR_RE = re.compile(r"\b(18|19|20)\d{2}\b")
DOI_RE = re.compile(r"\b(?:doi:\s*)?(10\.\d{4,9}/[-._;()/:A-Za-z0-9]+)", re.I)


@dataclass
class ParsedRef:
    num: int
    raw_text: str
    ref_type: str = "journal_article"
    authors: List[Dict[str, str]] = field(default_factory=list)
    title: str = ""
    journal: str = ""
    book_title: str = ""
    publisher: str = ""
    year: str = ""
    volume: str = ""
    issue: str = ""
    pages: str = ""
    doi: str = ""
    parsed: bool = False


def paragraph_text(p: etree._Element) -> str:
    return "".join(t.text or "" for t in p.iter(W + "t"))


def normalize_space(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def clean_field_text(text: str) -> str:
    text = normalize_space(text)
    text = re.sub(r"\s+([,.;:])", r"\1", text)
    text = re.sub(r"\s*\.\s*\.\s*", ". ", text)
    return text.strip()


def xml_escape(s: str) -> str:
    if not s:
        return ""
    return (
        s.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def read_docx_xml(docx_path: str) -> Tuple[Dict[str, bytes], etree._Element]:
    with zipfile.ZipFile(docx_path, "r") as zin:
        files = {name: zin.read(name) for name in zin.namelist()}
    if "word/document.xml" not in files:
        raise ValueError("word/document.xml not found. Input must be a valid DOCX file.")
    parser = etree.XMLParser(remove_blank_text=False, recover=True)
    root = etree.fromstring(files["word/document.xml"], parser)
    return files, root


def get_body(root: etree._Element) -> etree._Element:
    body = root.find(W + "body")
    if body is None:
        raise ValueError("No w:body element found in document.xml")
    return body


def body_paragraphs(body: etree._Element) -> List[etree._Element]:
    return list(body.iterchildren(W + "p"))


def find_reference_bounds(paras: Sequence[etree._Element]) -> Tuple[int, int]:
    ref_idx: Optional[int] = None
    for i, p in enumerate(paras):
        txt = normalize_space(paragraph_text(p))
        if REF_HEADING_RE.match(txt):
            ref_idx = i
            break
    if ref_idx is None:
        raise ValueError("References heading not found. Expected a paragraph named References, Reference, or Bibliography.")

    end_idx = len(paras)
    for i in range(ref_idx + 1, len(paras)):
        txt = normalize_space(paragraph_text(paras[i])).lower()
        if txt in STOP_HEADINGS:
            end_idx = i
            break
    return ref_idx, end_idx


def parse_reference_paragraphs(paras: Sequence[etree._Element], ref_idx: int, end_idx: int) -> List[Tuple[int, str]]:
    refs: List[Tuple[int, str]] = []
    current_num: Optional[int] = None
    current_text: List[str] = []

    def flush() -> None:
        nonlocal current_num, current_text
        if current_num is not None:
            refs.append((current_num, clean_field_text(" ".join(current_text))))
        current_num = None
        current_text = []

    for p in paras[ref_idx + 1 : end_idx]:
        txt = clean_field_text(paragraph_text(p))
        if not txt:
            continue
        m = REF_START_RE.match(txt)
        if m:
            flush()
            current_num = int(m.group(1))
            current_text = [m.group(2).strip()]
        elif current_num is not None:
            current_text.append(txt)
    flush()
    return refs


def parse_author_token(token: str) -> Optional[Dict[str, str]]:
    token = normalize_space(token.strip(" .;"))
    if not token:
        return None
    token = re.sub(r"\bet\s+al\.?$", "", token, flags=re.I).strip(" ,.;")
    if not token:
        return None

    # Common numeric styles: "Smith J", "Smith JA", "Smith J A".
    parts = token.split()
    if len(parts) == 1:
        return {"last": parts[0], "first": ""}

    # If token is "Last, First", keep that structure.
    if "," in token:
        last, first = token.split(",", 1)
        return {"last": normalize_space(last), "first": normalize_space(first)}

    # Treat trailing initials as first names; otherwise use the last token as initials/name.
    trailing = []
    while parts and re.match(r"^[A-Z][A-Z.\-]*$", parts[-1]):
        trailing.insert(0, parts.pop())
    if parts and trailing:
        return {"last": " ".join(parts), "first": " ".join(trailing).replace(".", "")}

    return {"last": " ".join(parts[:-1]), "first": parts[-1]}


def parse_authors(text: str) -> List[Dict[str, str]]:
    text = re.sub(r"\bet\s+al\.?", "", text or "", flags=re.I)
    text = text.replace(" and ", ", ")
    raw_tokens = [normalize_space(t) for t in re.split(r"\s*,\s*|\s*;\s*", text) if normalize_space(t)]
    authors = []
    for tok in raw_tokens:
        author = parse_author_token(tok)
        if author:
            authors.append(author)
    return authors


def split_reference_sentences(text: str) -> List[str]:
    # Split on period followed by whitespace and a capital/number. Avoid splitting inside DOI when possible.
    text = clean_field_text(text)
    if not text:
        return []
    protected = text.replace("doi:", "doi@@")
    parts = re.split(r"\.\s+(?=[A-Z0-9])", protected)
    return [p.replace("doi@@", "doi:").strip(" .") for p in parts if p.strip(" .")]


def parse_ref_text(num: int, raw_text: str) -> ParsedRef:
    ref = ParsedRef(num=num, raw_text=clean_field_text(raw_text))
    text = ref.raw_text
    doi_match = DOI_RE.search(text)
    if doi_match:
        ref.doi = doi_match.group(1).rstrip(". ;,")
        text_wo_doi = (text[: doi_match.start()] + text[doi_match.end() :]).strip()
    else:
        text_wo_doi = text

    sentences = split_reference_sentences(text_wo_doi)
    if len(sentences) < 2:
        ref.title = text
        ref.ref_type = "generic"
        return ref

    ref.authors = parse_authors(sentences[0])
    ref.title = sentences[1] if len(sentences) > 1 else text

    # Book chapter pattern: Title. In: Book Title. Publisher; Year:Pages.
    in_match = re.search(r"\bIn:\s*(.+?)\.\s*([^.;]+);\s*((?:18|19|20)\d{2})\s*:?\s*([A-Za-z0-9eE\-\u2013\u2014]+)?", text_wo_doi, re.I)
    if in_match:
        ref.ref_type = "book_section"
        ref.book_title = clean_field_text(in_match.group(1))
        ref.journal = ref.book_title
        ref.publisher = clean_field_text(in_match.group(2))
        ref.year = in_match.group(3)
        ref.pages = (in_match.group(4) or "").replace("\u2013", "-").replace("\u2014", "-")
        ref.parsed = bool(ref.title)
        return ref

    # Journal article: Authors. Title. Journal. Year;Volume(Issue):Pages.
    year_match = YEAR_RE.search(text_wo_doi)
    if year_match:
        ref.year = year_match.group(0)
        before_year = text_wo_doi[: year_match.start()].strip(" .")
        before_parts = split_reference_sentences(before_year)
        if len(before_parts) >= 3:
            ref.authors = parse_authors(before_parts[0])
            ref.title = before_parts[1]
            ref.journal = before_parts[-1]
        elif len(sentences) >= 3:
            ref.journal = sentences[2]

        after_year = text_wo_doi[year_match.end() :]
        vol_match = re.search(
            r";\s*([A-Za-z0-9]+)\s*(?:\(([^)]+)\))?\s*:?\s*([A-Za-z0-9eE\-\u2013\u2014]+)?",
            after_year,
        )
        if vol_match:
            ref.volume = vol_match.group(1) or ""
            ref.issue = vol_match.group(2) or ""
            ref.pages = (vol_match.group(3) or "").replace("\u2013", "-").replace("\u2014", "-")
        else:
            pages_match = re.search(r":\s*([A-Za-z0-9eE\-\u2013\u2014]+)", after_year)
            if pages_match:
                ref.pages = pages_match.group(1).replace("\u2013", "-").replace("\u2014", "-")

    ref.parsed = bool(ref.authors or ref.year or ref.journal or ref.title)
    if not ref.title:
        ref.title = text
        ref.ref_type = "generic"
    return ref


def crossref_lookup(title: str, year: str = "", mailto: str = "") -> Optional[Dict[str, object]]:
    if not title:
        return None
    params = {"query.title": title[:160], "rows": "5"}
    if year:
        params["filter"] = f"from-pub-date:{year},until-pub-date:{year}"
    if mailto:
        params["mailto"] = mailto
    url = "https://api.crossref.org/works?" + urllib.parse.urlencode(params)
    headers = {"User-Agent": "EndNoteFieldCodeConverter/1.0" + (f" (mailto:{mailto})" if mailto else "")}
    try:
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req, timeout=12) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except Exception:
        return None

    items = data.get("message", {}).get("items", [])
    if not items:
        return None

    query_words = set(re.findall(r"[a-z0-9]+", title.lower()))
    best = None
    best_score = 0.0
    for item in items:
        item_title = (item.get("title") or [""])[0]
        item_words = set(re.findall(r"[a-z0-9]+", item_title.lower()))
        if not query_words or not item_words:
            continue
        score = len(query_words & item_words) / max(1, len(query_words | item_words))
        if year:
            candidate_year = extract_crossref_year(item)
            if candidate_year == year:
                score += 0.1
        if score > best_score:
            best_score = score
            best = item

    if best is None or best_score < 0.25:
        return None
    return best


def extract_crossref_year(item: Dict[str, object]) -> str:
    for key in ("published-print", "published-online", "published", "created"):
        date = item.get(key, {}) if isinstance(item, dict) else {}
        parts = date.get("date-parts", []) if isinstance(date, dict) else []
        if parts and parts[0] and parts[0][0]:
            return str(parts[0][0])
    return ""


def enrich_with_crossref(refs: List[ParsedRef], mailto: str = "") -> None:
    for ref in refs:
        if not ref.title:
            continue
        print(f"  [{ref.num}] CrossRef: {ref.title[:70]}...", end=" ", flush=True)
        item = crossref_lookup(ref.title, ref.year, mailto)
        if not item:
            print("not found")
            time.sleep(0.25)
            continue
        doi = item.get("DOI", "") or ""
        if doi and not ref.doi:
            ref.doi = str(doi)
        titles = item.get("title") or []
        if titles:
            ref.title = str(titles[0])
        containers = item.get("container-title") or []
        if containers and not ref.journal:
            ref.journal = str(containers[0])
        if not ref.year:
            ref.year = extract_crossref_year(item)
        for attr, key in (("volume", "volume"), ("issue", "issue"), ("pages", "page"), ("publisher", "publisher")):
            val = item.get(key, "") or ""
            if val and not getattr(ref, attr):
                setattr(ref, attr, str(val))
        authors = []
        for a in item.get("author", []) or []:
            if isinstance(a, dict):
                authors.append({"last": a.get("family", "") or "", "first": a.get("given", "") or ""})
        if authors:
            ref.authors = authors
        print("doi=" + (ref.doi or "none"))
        time.sleep(0.25)


def expand_citation_numbers(raw: str) -> List[int]:
    raw = raw.replace(";", ",").replace("\u2013", "-").replace("\u2014", "-")
    nums: List[int] = []
    for part in raw.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            bits = [b.strip() for b in part.split("-", 1)]
            if len(bits) == 2 and bits[0].isdigit() and bits[1].isdigit():
                start, end = int(bits[0]), int(bits[1])
                if start <= end and end - start <= 200:
                    nums.extend(range(start, end + 1))
                else:
                    nums.extend([start, end])
            continue
        if part.isdigit():
            nums.append(int(part))
    seen = set()
    out = []
    for n in nums:
        if n not in seen:
            out.append(n)
            seen.add(n)
    return out


def ref_type_code(ref: ParsedRef) -> Tuple[str, str]:
    if ref.ref_type == "book_section":
        return "Book Section", "5"
    if ref.ref_type == "generic":
        return "Unpublished", "39"
    return "Journal Article", "17"


def build_record_xml(ref: ParsedRef) -> str:
    ref_name, ref_code = ref_type_code(ref)
    journal = ref.journal or ref.book_title
    authors_xml = "".join(
        f"<author>{xml_escape(a.get('last', ''))}, {xml_escape(a.get('first', ''))}</author>"
        for a in ref.authors
    )
    parts = [
        "<record>",
        f"<rec-number>{ref.num}</rec-number>",
        f"<foreign-keys><key app=\"EN\" db-id=\"converter\">{ref.num}</key></foreign-keys>",
        f"<ref-type name=\"{ref_name}\">{ref_code}</ref-type>",
    ]
    if authors_xml:
        parts.append(f"<contributors><authors>{authors_xml}</authors></contributors>")
    parts.append("<titles>")
    if ref.title:
        parts.append(f"<title>{xml_escape(ref.title)}</title>")
    if journal:
        parts.append(f"<secondary-title>{xml_escape(journal)}</secondary-title>")
    if ref.book_title and ref.ref_type == "book_section":
        parts.append(f"<book-title>{xml_escape(ref.book_title)}</book-title>")
    parts.append("</titles>")
    if journal:
        parts.append(f"<periodical><full-title>{xml_escape(journal)}</full-title></periodical>")
    if ref.pages:
        parts.append(f"<pages>{xml_escape(ref.pages)}</pages>")
    if ref.volume:
        parts.append(f"<volume>{xml_escape(ref.volume)}</volume>")
    if ref.issue:
        parts.append(f"<number>{xml_escape(ref.issue)}</number>")
    if ref.year:
        parts.append(f"<dates><year>{xml_escape(ref.year)}</year></dates>")
    if ref.publisher:
        parts.append(f"<publisher>{xml_escape(ref.publisher)}</publisher>")
    if ref.doi:
        parts.append(f"<electronic-resource-num>{xml_escape(ref.doi)}</electronic-resource-num>")
    if ref.raw_text:
        parts.append(f"<notes>{xml_escape(ref.raw_text)}</notes>")
    parts.append("</record>")
    return "".join(parts)


def build_en_cite_xml(nums: Sequence[int], ref_by_num: Dict[int, ParsedRef]) -> Tuple[str, str]:
    cites = []
    for num in nums:
        ref = ref_by_num.get(num)
        if ref is None:
            ref = ParsedRef(num=num, raw_text=f"[{num}]", title=f"[{num}]", ref_type="generic")
        author_str = "; ".join(
            f"{a.get('last', '')}, {a.get('first', '')}".strip().strip(",") for a in ref.authors
        )
        cite_parts = ["<Cite>"]
        if author_str:
            cite_parts.append(f"<Author>{xml_escape(author_str)}</Author>")
        if ref.year:
            cite_parts.append(f"<Year>{xml_escape(ref.year)}</Year>")
        cite_parts.append(f"<RecNum>{num}</RecNum>")
        cite_parts.append(f"<DisplayText>[{num}]</DisplayText>")
        cite_parts.append(build_record_xml(ref))
        cite_parts.append("</Cite>")
        cites.append("".join(cite_parts))
    display = "[" + ",".join(str(n) for n in nums) + "]"
    return "<EndNote>" + "".join(cites) + "</EndNote>", display


def clone_run_properties(src_run: Optional[etree._Element]) -> Optional[etree._Element]:
    if src_run is None:
        return None
    rpr = src_run.find(W + "rPr")
    return copy.deepcopy(rpr) if rpr is not None else None


def make_text_run(text: str, rpr: Optional[etree._Element] = None) -> etree._Element:
    r = etree.Element(W + "r")
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    t = etree.SubElement(r, W + "t")
    if text.startswith(" ") or text.endswith(" "):
        t.set("{%s}space" % NS_XML, "preserve")
    t.text = text
    return r


def make_field_runs(instr_text: str, display: str, rpr: Optional[etree._Element] = None) -> List[etree._Element]:
    runs: List[etree._Element] = []
    r = etree.Element(W + "r")
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    etree.SubElement(r, W + "fldChar", attrib={W + "fldCharType": "begin"})
    runs.append(r)

    r = etree.Element(W + "r")
    instr = etree.SubElement(r, W + "instrText")
    instr.set("{%s}space" % NS_XML, "preserve")
    instr.text = instr_text
    runs.append(r)

    r = etree.Element(W + "r")
    etree.SubElement(r, W + "fldChar", attrib={W + "fldCharType": "separate"})
    runs.append(r)

    runs.append(make_text_run(display, rpr))

    r = etree.Element(W + "r")
    etree.SubElement(r, W + "fldChar", attrib={W + "fldCharType": "end"})
    runs.append(r)
    return runs


def paragraph_text_runs(p: etree._Element) -> List[etree._Element]:
    return [r for r in p.iterchildren(W + "r") if r.find(W + "t") is not None]


def replace_citations_in_paragraph(
    p: etree._Element,
    ref_by_num: Dict[int, ParsedRef],
    convert_unknown: bool = False,
) -> Tuple[int, List[str]]:
    text = paragraph_text(p)
    matches = list(CITATION_RE.finditer(text))
    if not matches:
        return 0, []

    valid_matches = []
    warnings = []
    for m in matches:
        nums = expand_citation_numbers(m.group(1))
        if not nums:
            continue
        unknown = [n for n in nums if n not in ref_by_num]
        if unknown and not convert_unknown:
            warnings.append(f"Skipped citation {m.group(0)} because reference number(s) not found: {unknown}")
            continue
        valid_matches.append((m, nums))

    if not valid_matches:
        return 0, warnings

    text_runs = paragraph_text_runs(p)
    template_rpr = clone_run_properties(text_runs[0] if text_runs else None)
    new_runs: List[etree._Element] = []
    prev = 0
    for m, nums in valid_matches:
        if m.start() > prev:
            new_runs.append(make_text_run(text[prev : m.start()], template_rpr))
        en_xml, display = build_en_cite_xml(nums, ref_by_num)
        new_runs.extend(make_field_runs(" ADDIN EN.CITE " + en_xml, display, template_rpr))
        prev = m.end()
    if prev < len(text):
        new_runs.append(make_text_run(text[prev:], template_rpr))

    # Remove existing direct child runs only; keep pPr and other paragraph children.
    for child in list(p):
        if child.tag == W + "r":
            p.remove(child)
    for nr in new_runs:
        p.append(nr)
    return len(valid_matches), warnings


def format_author_list(ref: ParsedRef) -> str:
    if not ref.authors:
        return ""
    names = []
    for a in ref.authors:
        last = a.get("last", "").strip()
        first = a.get("first", "").strip()
        if last and first:
            names.append(f"{last} {first}")
        else:
            names.append(last or first)
    return ", ".join([n for n in names if n])


def format_bibliography_entry(ref: ParsedRef) -> str:
    auth = format_author_list(ref)
    pieces = []
    if auth:
        pieces.append(auth + ".")
    if ref.title:
        pieces.append(ref.title.rstrip(".") + ".")
    journal = ref.journal or ref.book_title
    if journal:
        pieces.append(journal.rstrip(".") + ".")
    tail = ""
    if ref.year:
        tail += ref.year
    if ref.volume:
        tail += ";" + ref.volume
        if ref.issue:
            tail += f"({ref.issue})"
    if ref.pages:
        tail += ":" + ref.pages
    if tail:
        pieces.append(tail.rstrip(".") + ".")
    if ref.doi:
        pieces.append("doi:" + ref.doi.rstrip(".") + ".")
    if not pieces:
        pieces = [ref.raw_text]
    return f"[{ref.num}] " + " ".join(pieces)


def make_bibliography_paragraph(entry: str, begin: bool = False, end: bool = False) -> etree._Element:
    p = etree.Element(W + "p")
    ppr = etree.SubElement(p, W + "pPr")
    etree.SubElement(ppr, W + "pStyle", attrib={W + "val": "Bibliography"})
    if begin:
        r = etree.SubElement(p, W + "r")
        etree.SubElement(r, W + "fldChar", attrib={W + "fldCharType": "begin"})
        r = etree.SubElement(p, W + "r")
        instr = etree.SubElement(r, W + "instrText")
        instr.set("{%s}space" % NS_XML, "preserve")
        instr.text = " ADDIN EN.REFLIST "
        r = etree.SubElement(p, W + "r")
        etree.SubElement(r, W + "fldChar", attrib={W + "fldCharType": "separate"})
    p.append(make_text_run(entry))
    if end:
        r = etree.SubElement(p, W + "r")
        etree.SubElement(r, W + "fldChar", attrib={W + "fldCharType": "end"})
    return p


def build_reflist_paragraphs(refs: Sequence[ParsedRef]) -> List[etree._Element]:
    sorted_refs = sorted(refs, key=lambda r: r.num)
    if not sorted_refs:
        return [make_bibliography_paragraph("", begin=True, end=True)]
    paras = []
    for i, ref in enumerate(sorted_refs):
        paras.append(
            make_bibliography_paragraph(
                format_bibliography_entry(ref),
                begin=(i == 0),
                end=(i == len(sorted_refs) - 1),
            )
        )
    return paras


def count_field_markers(xml_bytes: bytes) -> Tuple[int, int, int]:
    begin = len(re.findall(rb"fldCharType=['\"]begin['\"]", xml_bytes))
    separate = len(re.findall(rb"fldCharType=['\"]separate['\"]", xml_bytes))
    end = len(re.findall(rb"fldCharType=['\"]end['\"]", xml_bytes))
    return begin, separate, end


def convert_docx(
    input_path: str,
    output_path: str,
    skip_crossref: bool = False,
    dry_run: bool = False,
    convert_unknown_citations: bool = False,
    keep_reference_text: bool = False,
    mailto: str = "",
) -> bool:
    print(f"Input: {input_path}")
    files, root = read_docx_xml(input_path)
    body = get_body(root)
    paras = body_paragraphs(body)
    ref_idx, end_idx = find_reference_bounds(paras)

    raw_refs = parse_reference_paragraphs(paras, ref_idx, end_idx)
    print(f"References found: {len(raw_refs)}")

    refs = [parse_ref_text(num, txt) for num, txt in raw_refs]
    parsed_count = sum(1 for r in refs if r.parsed)
    print(f"Parsed references: {parsed_count}/{len(refs)}")
    for r in refs:
        if not r.parsed:
            print(f"  Unparsed reference [{r.num}]: {r.raw_text[:140]}")

    if not skip_crossref:
        print("CrossRef lookup: enabled")
        enrich_with_crossref(refs, mailto=mailto)
    else:
        print("CrossRef lookup: skipped")

    ref_by_num = {r.num: r for r in refs}

    converted = 0
    all_warnings: List[str] = []
    for p in paras[:ref_idx]:
        n, warnings = replace_citations_in_paragraph(p, ref_by_num, convert_unknown=convert_unknown_citations)
        converted += n
        all_warnings.extend(warnings)
    print(f"Converted citations: {converted}")
    for warning in all_warnings[:50]:
        print("  " + warning)
    if len(all_warnings) > 50:
        print(f"  ... {len(all_warnings) - 50} additional skipped citation warnings")

    if not keep_reference_text:
        # Recompute paragraphs after citation replacement, then remove old reference entries.
        paras = body_paragraphs(body)
        ref_heading = paras[ref_idx]
        current_ref_entries = paras[ref_idx + 1 : end_idx]
        for p in current_ref_entries:
            try:
                body.remove(p)
            except ValueError:
                pass
        insert_pos = list(body).index(ref_heading) + 1
        for p in build_reflist_paragraphs(refs):
            body.insert(insert_pos, p)
            insert_pos += 1

    xml_bytes = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    b, s, e = count_field_markers(xml_bytes)
    balanced = b == s == e
    print(f"Field markers: begin={b} separate={s} end={e}")
    print("Field markers balanced: " + ("yes" if balanced else "no"))

    if dry_run:
        print("Dry run only: no output written")
        return balanced and len(refs) > 0

    files["word/document.xml"] = xml_bytes
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)
    print(f"Output: {output_path}")
    return balanced and len(refs) > 0


def default_output_path(input_path: str) -> str:
    p = Path(input_path)
    return str(p.with_name(p.stem + "_EndNote" + p.suffix))


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Convert DOCX numeric citations into EndNote field codes.")
    parser.add_argument("-i", "--input", required=True, help="Input DOCX path")
    parser.add_argument("-o", "--output", default=None, help="Output DOCX path; default: *_EndNote.docx")
    parser.add_argument("--skip-crossref", action="store_true", help="Skip CrossRef DOI/metadata lookup")
    parser.add_argument("--offline", action="store_true", help="Alias for --skip-crossref")
    parser.add_argument("--dry-run", action="store_true", help="Parse and report without writing output")
    parser.add_argument(
        "--convert-unknown-citations",
        action="store_true",
        help="Convert numeric brackets even when the number is missing from the parsed references",
    )
    parser.add_argument(
        "--keep-reference-text",
        action="store_true",
        help="Keep original plain-text reference entries after inserting EndNote bibliography fields",
    )
    parser.add_argument("--mailto", default="", help="Optional email for CrossRef polite pool")
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = build_arg_parser().parse_args(argv)
    input_path = args.input
    if not os.path.exists(input_path):
        print(f"ERROR: input file not found: {input_path}", file=sys.stderr)
        return 2
    output_path = args.output or default_output_path(input_path)
    try:
        ok = convert_docx(
            input_path=input_path,
            output_path=output_path,
            skip_crossref=bool(args.skip_crossref or args.offline),
            dry_run=bool(args.dry_run),
            convert_unknown_citations=bool(args.convert_unknown_citations),
            keep_reference_text=bool(args.keep_reference_text),
            mailto=args.mailto,
        )
    except Exception as exc:
        print("ERROR: " + str(exc), file=sys.stderr)
        return 1
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())

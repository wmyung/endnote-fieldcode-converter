#!/usr/bin/env python3
from __future__ import annotations
import copy, importlib.util, re, sys, zipfile
from pathlib import Path
from lxml import etree

BASE = Path(__file__).with_name('endnote_converter.py')
spec = importlib.util.spec_from_file_location('endnote_converter', BASE)
ec = importlib.util.module_from_spec(spec)
sys.modules['endnote_converter'] = ec
spec.loader.exec_module(ec)  # type: ignore

W = ec.W
NS_XML = ec.NS_XML

# Support Lancet-style reference list entries like: "1  Author..."
ec.REF_START_RE = re.compile(r"^\s*(?:\[\s*(\d+)\s*\]|(\d+)(?:\.|\s+))\s*(.*)$", re.S)

# Prefer longer citation numbers first. If single digits are tried first,
# runs like superscript "14" are incorrectly converted as two fields: "1" and "4".
CITE_NUM = r"(?:[12]\d\d|[1-9]\d|[1-9])"
CITE_SEQ = rf"{CITE_NUM}(?:\s*(?:,|;|\-|\u2013|\u2014)\s*{CITE_NUM})*"
PLAIN_CITE_RE = re.compile(rf"(?<=[A-Za-z\),.;:])({CITE_SEQ})(?=(?:\s|[.;,:)\u2013\u2014-]|$))")
SUPER_CITE_RE = re.compile(rf"({CITE_SEQ})")

# Hard-coded non-citation tokens that commonly look like bare numeric references.
# Keep this narrower than "any uppercase letter before a number": plain-text rescue
# should still recover missed citations like diagnosis.16, disorder25, and use,15.
ICD10_PREFIXES = set("ABCDEFGHJKLMNPQRSTVWXYZ")  # WHO ICD-10 first-character codes; U is reserved.
AIR_POLLUTANT_PREFIXES = {
    "PM", "NO", "NOX", "SO", "SOX", "CO", "CO2", "O3", "NH3", "CH4", "BC", "EC", "OC", "VOC", "VOCs",
}
ELEMENT_SYMBOLS = {
    "H", "He", "Li", "Be", "B", "C", "N", "O", "F", "Ne", "Na", "Mg", "Al", "Si", "P", "S", "Cl", "Ar",
    "K", "Ca", "Sc", "Ti", "V", "Cr", "Mn", "Fe", "Co", "Ni", "Cu", "Zn", "Ga", "Ge", "As", "Se", "Br", "Kr",
    "Rb", "Sr", "Y", "Zr", "Nb", "Mo", "Tc", "Ru", "Rh", "Pd", "Ag", "Cd", "In", "Sn", "Sb", "Te", "I", "Xe",
    "Cs", "Ba", "La", "Ce", "Pr", "Nd", "Pm", "Sm", "Eu", "Gd", "Tb", "Dy", "Ho", "Er", "Tm", "Yb", "Lu",
    "Hf", "Ta", "W", "Re", "Os", "Ir", "Pt", "Au", "Hg", "Tl", "Pb", "Bi", "Po", "At", "Rn", "Fr", "Ra",
    "Ac", "Th", "Pa", "U", "Np", "Pu", "Am", "Cm", "Bk", "Cf", "Es", "Fm", "Md", "No", "Lr", "Rf", "Db", "Sg",
    "Bh", "Hs", "Mt", "Ds", "Rg", "Cn", "Nh", "Fl", "Mc", "Lv", "Ts", "Og",
}


def is_superscript(run):
    return bool(run.xpath('./w:rPr/w:vertAlign[@w:val="superscript"]', namespaces={'w': ec.NS_W}))


def clone_rpr(run):
    rpr = run.find(W+'rPr')
    return copy.deepcopy(rpr) if rpr is not None else None


def ensure_superscript_rpr(run):
    rpr = clone_rpr(run)
    if rpr is None:
        rpr = etree.Element(W+'rPr')
    va = rpr.find(W+'vertAlign')
    if va is None:
        va = etree.SubElement(rpr, W+'vertAlign')
    va.set(W+'val', 'superscript')
    return rpr


def make_text_run(text, rpr=None):
    return ec.make_text_run(text, rpr)


def build_en_cite_xml_display(nums, ref_by_num, visible):
    cites=[]
    for num in nums:
        ref = ref_by_num.get(num)
        if ref is None:
            ref = ec.ParsedRef(num=num, raw_text=str(num), title=str(num), ref_type='generic')
        author_str = '; '.join(f"{a.get('last','')}, {a.get('first','')}".strip().strip(',') for a in ref.authors)
        cite=['<Cite>']
        if author_str:
            cite.append(f'<Author>{ec.xml_escape(author_str)}</Author>')
        if ref.year:
            cite.append(f'<Year>{ec.xml_escape(ref.year)}</Year>')
        cite.append(f'<RecNum>{num}</RecNum>')
        cite.append(f'<DisplayText>{ec.xml_escape(visible)}</DisplayText>')
        cite.append(ec.build_record_xml(ref))
        cite.append('</Cite>')
        cites.append(''.join(cite))
    return '<EndNote>' + ''.join(cites) + '</EndNote>'


def field_runs_for_visible(visible, nums, ref_by_num, rpr):
    en_xml = build_en_cite_xml_display(nums, ref_by_num, visible)
    return ec.make_field_runs(' ADDIN EN.CITE ' + en_xml, visible, rpr)


def contiguous_alpha_before(text, pos):
    i = pos
    while i > 0 and text[i-1].isalpha():
        i -= 1
    return text[i:pos]


def run_text(run):
    return ''.join(t.text or '' for t in run.iter(W+'t'))


def set_run_text(run, text):
    texts = list(run.iter(W+'t'))
    if not texts:
        return
    texts[0].text = text
    for t in texts[1:]:
        t.text = ''


def run_has_field_code(run):
    return run.find(W+'fldChar') is not None or run.find(W+'instrText') is not None


def strip_existing_field_code_runs(p):
    """Remove stale Word field-control runs before rebuilding EndNote fields.

    Some supposedly "clean" manuscripts still contain empty w:fldChar begin/separate/end
    runs from broken EndNote fields around bare superscripts. If left in place, newly
    inserted fields become nested in those stale markers, producing duplicate or empty
    citation fields and apparent over-citation. The Lancet converter rebuilds fields
    from visible text, so these stale field-code runs must be removed first.
    """
    removed = 0
    for child in list(p):
        if child.tag == W+'r' and run_has_field_code(child):
            p.remove(child)
            removed += 1
    return removed


def previous_visible_text(run, limit=80):
    parent = run.getparent()
    if parent is None:
        return ''
    chunks = []
    for sib in reversed(list(parent)[:parent.index(run)]):
        if sib.tag != W+'r' or run_has_field_code(sib):
            continue
        txt = run_text(sib)
        if txt:
            chunks.append(txt)
            if sum(len(c) for c in chunks) >= limit:
                break
    return ''.join(reversed(chunks))[-limit:]


def next_visible_text(run, limit=40):
    parent = run.getparent()
    if parent is None:
        return ''
    chunks = []
    for sib in list(parent)[parent.index(run)+1:]:
        if sib.tag != W+'r' or run_has_field_code(sib):
            continue
        txt = run_text(sib)
        if txt:
            chunks.append(txt)
            if sum(len(c) for c in chunks) >= limit:
                break
    return ''.join(chunks)[:limit]


def is_blocked_superscript_citation_context(run, visible):
    """Reject superscript numerals that are part of scientific/pollutant notation.

    This is intentionally contextual rather than a blanket uppercase-letter rule:
    true Lancet citations can follow words, but PM2·5/NO2/O3/C14-like tokens are
    identified from the immediate left/right text around the superscript run.
    """
    left = previous_visible_text(run)
    right = next_visible_text(run)
    left_compact = re.sub(r"\s+", "", left)
    right_lstrip = right.lstrip()
    visible_clean = visible.strip()

    # PM2·5 can be stored as normal "PM" + superscript "2" + normal "·5".
    # Also protect cases where only the trailing decimal digit is superscript.
    pollutant_base = r"(?:PM|NOx?|SOx?|CO2?|O3?|NH3|CH4|BC|EC|OC|VOC|VOCs|SF6)"
    if re.search(pollutant_base + r"(?:\d+)?(?:[\.\u00b7])?$", left_compact, re.I):
        if re.fullmatch(r"\d{1,3}", visible_clean) and (not right_lstrip or re.match(r"^[\.\u00b7\d\"')\],;: -]", right_lstrip)):
            return True

    # Chemical/isotope-style tokens, e.g. C14, I131, Na24.
    m = re.search(r"([A-Z][a-z]?)$", left_compact)
    if m and m.group(1) in ELEMENT_SYMBOLS and re.fullmatch(r"\d{1,3}", visible_clean):
        if not right_lstrip or re.match(r"^[\s,.;:)\]-]", right_lstrip):
            return True

    # Scientific notation/exponents such as 10^5 or x^2 are not references.
    if re.search(r"(?:10|[×xX*\u22c5])$", left_compact) and re.fullmatch(r"[\d+\-\u2212]+", visible_clean):
        return True

    return False


def normalize_split_superscript_citations(p):
    """Merge citation text split across adjacent superscript runs before conversion.

    Word often stores a visible citation like 27,31 as empty field-control-like runs
    plus separate superscript text runs such as "2" and "7,31". Without merging,
    the converter creates two EndNote fields. Empty non-superscript runs between
    superscript text runs are ignored during this normalization.
    """
    children = [c for c in list(p) if c.tag == W+'r' and c.find(W+'t') is not None]
    i = 0; merged = 0
    while i < len(children):
        child = children[i]
        if is_superscript(child):
            nonempty_sup = []
            j = i
            while j < len(children):
                txt = run_text(children[j])
                if is_superscript(children[j]):
                    if txt:
                        nonempty_sup.append(j)
                    j += 1
                    continue
                if txt == '':
                    j += 1
                    continue
                break
            combined = ''.join(run_text(children[k]) for k in nonempty_sup)
            if len(nonempty_sup) > 1 and re.fullmatch(rf"\s*{CITE_SEQ}\s*", combined):
                set_run_text(children[nonempty_sup[0]], combined)
                for k in nonempty_sup[1:]:
                    set_run_text(children[k], '')
                merged += 1
            i = j
        else:
            i += 1
    return merged


def is_blocked_plain_citation_context(text, match):
    """Return True when a plain-text numeric match is a biomedical code/metric, not a citation.

    This intentionally applies only to plain-text rescue, not to already-superscript
    citation runs. It uses hard-coded vocabularies for ICD-10 prefixes, air-pollutant
    abbreviations, and element symbols rather than a blanket uppercase-letter rule.
    """
    start, end = match.span(1)
    prefix = contiguous_alpha_before(text, start)
    if not prefix:
        return False

    # Air pollution metrics: PM10, PM2.5/PM2·5, NO2, SO2, O3, CO2, NOx, SOx, etc.
    if prefix in AIR_POLLUTANT_PREFIXES or prefix.upper() in {p.upper() for p in AIR_POLLUTANT_PREFIXES}:
        return True

    # Isotopes/elements: C14, Na24, Fe59, I131, etc.
    if prefix in ELEMENT_SYMBOLS:
        return True

    # ICD-10 codes/ranges: E14, E14–E10, I14–I16, I11.
    # This is limited to valid ICD-10 one-letter prefixes, not arbitrary words.
    if prefix in ICD10_PREFIXES:
        after = text[end:end+4]
        if not after or re.match(r"^[\s,.;:)\]-]", after) or re.match(r"^[\-\u2013\u2014][A-Z]?\d", after):
            return True

    return False


def split_text_with_citations(text, regex, ref_by_num, run, force_super=True, protect_plain_context=False):
    out=[]; prev=0; converted=0; warnings=[]
    for m in regex.finditer(text):
        visible = m.group(1)
        if protect_plain_context and is_blocked_plain_citation_context(text, m):
            warnings.append(f'Skipped plain-text numeric token {visible}: protected biomedical code/metric context')
            continue
        nums = ec.expand_citation_numbers(visible)
        if not nums:
            continue
        unknown=[n for n in nums if n not in ref_by_num]
        if unknown:
            warnings.append(f'Skipped citation {visible}: unknown refs {unknown}')
            continue
        if m.start()>prev:
            out.append(make_text_run(text[prev:m.start()], clone_rpr(run)))
        rpr = ensure_superscript_rpr(run) if force_super else clone_rpr(run)
        out.extend(field_runs_for_visible(visible, nums, ref_by_num, rpr))
        converted += 1
        prev=m.end()
    if converted and prev < len(text):
        out.append(make_text_run(text[prev:], clone_rpr(run)))
    return out, converted, warnings


def convert_paragraph(p, ref_by_num):
    removed_fields = strip_existing_field_code_runs(p)
    normalize_split_superscript_citations(p)
    converted=0; warnings=[]; replacements=[]
    if removed_fields:
        warnings.append(f'Removed {removed_fields} stale field-code run(s) before rebuilding citations')
    for child in list(p):
        if child.tag != W+'r' or child.find(W+'t') is None:
            continue
        if child.find(W+'instrText') is not None or child.find(W+'fldChar') is not None:
            continue
        text=''.join(t.text or '' for t in child.iter(W+'t'))
        if not text:
            continue
        if is_superscript(child):
            if re.fullmatch(rf"\s*{CITE_SEQ}\s*", text):
                if is_blocked_superscript_citation_context(child, text):
                    warnings.append(f'Skipped superscript numeric token {text.strip()}: protected scientific/pollutant context')
                    continue
                new, n, w = split_text_with_citations(text, SUPER_CITE_RE, ref_by_num, child, force_super=True)
            else:
                continue
        else:
            new, n, w = split_text_with_citations(text, PLAIN_CITE_RE, ref_by_num, child, force_super=True, protect_plain_context=True)
        if n:
            replacements.append((child,new))
            converted += n
            warnings.extend(w)
    for old,newruns in replacements:
        parent=old.getparent(); idx=parent.index(old); parent.remove(old)
        for nr in reversed(newruns):
            parent.insert(idx, nr)
    return converted, warnings


def make_bib_paragraph_original(num, raw_text, begin=False, end=False):
    p = etree.Element(W+'p')
    if begin:
        r=etree.SubElement(p,W+'r'); etree.SubElement(r,W+'fldChar',attrib={W+'fldCharType':'begin'})
        r=etree.SubElement(p,W+'r'); instr=etree.SubElement(r,W+'instrText'); instr.set('{%s}space'%NS_XML,'preserve'); instr.text=' ADDIN EN.REFLIST '
        r=etree.SubElement(p,W+'r'); etree.SubElement(r,W+'fldChar',attrib={W+'fldCharType':'separate'})
    p.append(make_text_run(f'{num}  {raw_text}'))
    if end:
        r=etree.SubElement(p,W+'r'); etree.SubElement(r,W+'fldChar',attrib={W+'fldCharType':'end'})
    return p


def convert(input_path, output_path, dry_run=False):
    files, root = ec.read_docx_xml(input_path)
    body=ec.get_body(root); paras=ec.body_paragraphs(body)
    ref_idx,end_idx=ec.find_reference_bounds(paras)
    raw_refs=ec.parse_reference_paragraphs(paras, ref_idx, end_idx)
    print(f'References found: {len(raw_refs)}')
    refs=[ec.parse_ref_text(n,t) for n,t in raw_refs]
    print(f'Parsed references: {sum(r.parsed for r in refs)}/{len(refs)}')
    ref_by_num={r.num:r for r in refs}
    converted=0; warnings=[]
    for p in paras[:ref_idx]:
        n,w=convert_paragraph(p, ref_by_num)
        converted += n; warnings.extend(w)
    print(f'Converted citations: {converted}')
    for w in warnings[:50]: print('  '+w)
    paras=ec.body_paragraphs(body)
    ref_heading=paras[ref_idx]
    for p in paras[ref_idx+1:end_idx]:
        try: body.remove(p)
        except ValueError: pass
    insert_pos=list(body).index(ref_heading)+1
    sorted_refs=sorted(refs, key=lambda r:r.num)
    for i,r in enumerate(sorted_refs):
        body.insert(insert_pos, make_bib_paragraph_original(r.num, r.raw_text, begin=(i==0), end=(i==len(sorted_refs)-1)))
        insert_pos += 1
    xml_bytes=etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    b,s,e=ec.count_field_markers(xml_bytes)
    print(f'Field markers: begin={b} separate={s} end={e}')
    print('Field markers balanced: '+('yes' if b==s==e else 'no'))
    if dry_run:
        print('Dry run only: no output written')
        return b==s==e and len(refs)>0 and converted>0
    files['word/document.xml']=xml_bytes
    with zipfile.ZipFile(output_path,'w',zipfile.ZIP_DEFLATED) as zout:
        for name,data in files.items(): zout.writestr(name,data)
    print(f'Output: {output_path}')
    return b==s==e and len(refs)>0 and converted>0


def main(argv=None):
    import argparse
    ap=argparse.ArgumentParser(description='Convert Lancet-style superscript numeric citations to EndNote fields.')
    ap.add_argument('-i','--input',required=True)
    ap.add_argument('-o','--output',required=True)
    ap.add_argument('--dry-run',action='store_true')
    args=ap.parse_args(argv)
    return 0 if convert(args.input,args.output,args.dry_run) else 1

if __name__ == '__main__':
    raise SystemExit(main())

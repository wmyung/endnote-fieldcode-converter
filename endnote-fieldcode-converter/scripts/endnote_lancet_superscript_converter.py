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


def split_text_with_citations(text, regex, ref_by_num, run, force_super=True):
    out=[]; prev=0; converted=0; warnings=[]
    for m in regex.finditer(text):
        visible = m.group(1)
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
    converted=0; warnings=[]; replacements=[]
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
                new, n, w = split_text_with_citations(text, SUPER_CITE_RE, ref_by_num, child, force_super=True)
            else:
                continue
        else:
            new, n, w = split_text_with_citations(text, PLAIN_CITE_RE, ref_by_num, child, force_super=True)
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

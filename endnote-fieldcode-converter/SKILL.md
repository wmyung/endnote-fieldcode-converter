---
name: endnote-fieldcode-converter
description: convert plain-text numbered citations and reference lists in docx manuscripts into endnote-compatible word field codes with embedded traveling library data. use when a user asks to turn numeric citations such as [1], [1,2], or [1-3] and a references section in a .docx file into endnote add-in fields, create an endnote-ready manuscript, insert ADDIN EN.CITE / ADDIN EN.REFLIST field codes, perform offline conversion, or enrich references with crossref doi lookups.
---

# EndNote Field Code Converter

## Purpose

Use this skill to convert a manuscript DOCX that contains plain-text numeric citations and a numbered reference section into a Word document with EndNote-compatible field codes. The bundled script inserts `ADDIN EN.CITE` fields for in-text citations and an `ADDIN EN.REFLIST` field for the bibliography display text.

## Inputs and outputs

Expected input:
- A `.docx` manuscript.
- In-text citations formatted as `[1]`, `[1,2]`, `[1, 2]`, `[1-3]`, or `[1,3-5]`.
- A heading named `References`, `Reference`, or `REFERENCES`.
- Numbered references beginning with `[1]`, `[2]`, etc. or `1.`, `2.`, etc. after the References heading.

Expected output:
- A new `.docx` file containing Word field-code markers for EndNote citations and a bibliography field.
- The visible citation text remains numeric, for example `[1]` or `[1,2]`, so the manuscript remains readable in Word.
- The output filename should normally end with `_EndNote.docx` unless the user requests another path.

## Required execution rule

Run the bundled Python script instead of trying to manually edit OOXML in conversation. Use Hermes `terminal` for code execution.

Install dependencies if needed:

```bash
python3 -m pip install lxml
```

## Standard workflow

1. Locate the user's input DOCX. If multiple DOCX files are present and the user did not specify one, choose the most likely manuscript file by filename and recent upload context.
2. If the DOCX uses Lancet-style bare superscript citations (`1`, `9–11`) and a reference list numbered as `1  Author...` rather than bracketed citations, use `scripts/endnote_lancet_superscript_converter.py` instead of the default script. It also converts missed plain inline citation numbers directly attached to text or punctuation, such as `disorder25`, `use,15`, and `diagnosis.16`, and preserves visible citations as superscript bare numbers. This Lancet converter requires `-o/--output` even for `--dry-run`, and it does **not** accept `--skip-crossref`; run it as `python3 scripts/endnote_lancet_superscript_converter.py -i manuscript.docx -o dry.docx --dry-run`.
3. If the user asks to change one reference, add a PubMed-specified citation at a specific sentence, or otherwise alter reference numbering, first edit a copy of the plain DOCX, then run the EndNote conversion. Keep the existing citation number attached to its original sentence and append new citations in the same superscript run/sequence (for example, change `27` to `27,31` at the requested sentence), then add the new reference at the end unless the user requests renumbering. Verify the target paragraph text before and after editing so references do not drift to the wrong sentence.
4. If the DOCX already contains EndNote fields and the user asks for BJOG display formatting, do not reconvert from plain text. Use `scripts/format_bjog_endnote_docx.py` to preserve `ADDIN EN.CITE` / `ADDIN EN.REFLIST` fields while changing visible citations to BJOG superscript numeric style. See `references/bjog-endnote-formatting.md`.
5. Run a dry run first when the file is unfamiliar:

```bash
python3 scripts/endnote_converter.py -i manuscript.docx --dry-run --skip-crossref
```

3. Convert offline by default for speed and reproducibility:

```bash
python3 scripts/endnote_converter.py -i manuscript.docx -o manuscript_EndNote.docx --skip-crossref
```

4. Use CrossRef enrichment only when the user asks for DOI recovery/metadata enrichment or when internet access is available and the user accepts slower execution:

```bash
python3 scripts/endnote_converter.py -i manuscript.docx -o manuscript_EndNote.docx
```

5. Return the generated DOCX link and report:
   - number of parsed references,
   - number of converted in-text citation occurrences,
   - whether field markers are balanced,
   - any unparsed references or skipped citations.

## Options

- `--skip-crossref`: do not query CrossRef; fastest and fully offline.
- `--offline`: alias for `--skip-crossref`.
- `--dry-run`: parse and report without writing output.
- `--convert-unknown-citations`: convert numeric citation brackets even if the number is not found in the parsed reference list. Without this flag, unknown numeric brackets are left as plain text to avoid converting bracketed numbers that are not citations.
- `--keep-reference-text`: keep the original plain-text reference entries after the EndNote bibliography field. Default behavior replaces the plain-text reference entries with an `ADDIN EN.REFLIST` field display.

## Quality checks

After conversion, inspect the script output. Treat these as blocking issues unless the user explicitly accepts them:

- `Field markers balanced: no`
- `Parsed references: 0`
- `Converted citations: 0` when the manuscript visibly contains `[N]` citations
- many `Unparsed reference` warnings
- many `Skipped citation` warnings for numbers that should exist in the reference list

For Lancet-style superscript conversion, also verify multi-digit citations at user-flagged or highlighted positions by inspecting the resulting field display and embedded reference metadata. A known pitfall is converting `14`, `16`, or `25` as separate one-digit fields (`1`+`4`, `1`+`6`, `2`+`5`) if the citation-number regex prefers one-digit alternatives first. The Lancet converter should prefer longer numeric alternatives first, e.g. `(?:[12]\d\d|[1-9]\d|[1-9])`, and validation should confirm the display text remains `14`, `16`, `25`, etc.

The Lancet converter first normalizes citation text split across adjacent superscript runs, because Word can store a visible citation such as `27,31` as separate superscript text runs like `2` and `7,31`. Without this normalization the output creates separate EndNote fields with incorrect display text. Verify merged multi-reference displays such as `27,31` and `6–8,25` after conversion.

The Lancet converter protects plain-text rescue against common non-reference tokens using hard-coded vocabularies rather than blanket uppercase-letter rules:
- ICD-10 one-letter prefixes, so `E14–E10`, `I14–I16`, and `I11` are not converted.
- Air-pollution metric prefixes, so `PM10`, `PM2.5`/`PM2·5`, `NO2`, `SO2`, `O3`, `CO2`, `NOx`, and related tokens are not converted.
- Chemical element symbols, so isotope-like tokens such as `C14`, `Na24`, `Fe59`, and `I131` are not converted.
This protection applies only to plain-text rescue; already-superscript citation runs are still converted.

If the user highlighted missed citations in the source DOCX, remove transient highlight formatting before final delivery unless the user asks to keep it, then re-run conversion and verify `yellow_run_count` or equivalent is zero.

If the conversion fails because the references heading is not found, ask the user whether the heading is named differently, or rerun after editing a copy of the DOCX to use a `References` heading.
- The requested PMID-derived title or replacement reference text appears in the reference list/traveling-library payload.
- The target sentence still contains the intended visible citation sequence (for example, `27,31`) and it is superscript when appropriate.
- `ADDIN EN.REFLIST` appears once and field-marker begin/separate/end counts are equal. `ADDIN EN.CITE` string counts may differ from the script's converted-citation count because one visible citation field can contain multiple cited references; rely on balanced field markers plus spot checks of target citations.

If the conversion fails because the references heading is not found, ask the user whether the heading is named differently, or rerun after editing a copy of the DOCX to use a `References` heading.

## Important limitations

- The parser is heuristic. It handles common biomedical numeric reference formats but cannot guarantee perfect CSL/EndNote metadata extraction from every journal style.
- Reference lists may use either bracketed numbering (`[1]`) or dotted numbering (`1.`); the installed local script supports both. If a dry run reports `References found: 0` while the document visibly has a reference list, inspect whether the numbering or heading format differs before using `--convert-unknown-citations`.
- EndNote must be installed in Word for the user to fully update, format, and manage the resulting fields.
- The script creates a traveling-library style XML payload from the parsed reference text. It does not connect to the user's local EndNote library.
- CrossRef lookup can improve DOI and metadata fields, but it depends on internet access and may fail silently for ambiguous titles.
- Complex citations inside Word text boxes, comments, footnotes, headers, or track-changes markup may require additional OOXML handling.

## Troubleshooting

Consult `references/usage.md` for command examples and expected manuscript structure.
Consult `references/validation.md` for post-conversion checks and common failure modes.
Consult `references/bjog-endnote-style.md` when the user asks for the BJOG EndNote `.ens` output style file.
Consult `references/bjog-endnote-formatting.md` and run `scripts/format_bjog_endnote_docx.py` when the user asks to make an EndNote-fielded DOCX match BJOG visible citation/reference formatting while keeping EndNote fields editable.
Consult `references/lancet-superscript-multidigit-pitfall.md` when user-flagged Lancet-style bare superscript citations such as `14`, `16`, or `25` appear missing or have wrong EndNote metadata after conversion.

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
- Numbered references beginning with `[1]`, `[2]`, etc. after the References heading.

Expected output:
- A new `.docx` file containing Word field-code markers for EndNote citations and a bibliography field.
- The visible citation text remains numeric, for example `[1]` or `[1,2]`, so the manuscript remains readable in Word.
- The output filename should normally end with `_EndNote.docx` unless the user requests another path.

## Required execution rule

Run the bundled Python script instead of trying to manually edit OOXML in conversation. Use `container.exec` for code execution.

Install dependencies if needed:

```bash
python3 -m pip install lxml
```

## Standard workflow

1. Locate the user's input DOCX. If multiple DOCX files are present and the user did not specify one, choose the most likely manuscript file by filename and recent upload context.
2. Run a dry run first when the file is unfamiliar:

```bash
python3 /home/oai/skills/endnote-fieldcode-converter/scripts/endnote_converter.py -i manuscript.docx --dry-run --skip-crossref
```

3. Convert offline by default for speed and reproducibility:

```bash
python3 /home/oai/skills/endnote-fieldcode-converter/scripts/endnote_converter.py -i manuscript.docx -o manuscript_EndNote.docx --skip-crossref
```

4. Use CrossRef enrichment only when the user asks for DOI recovery/metadata enrichment or when internet access is available and the user accepts slower execution:

```bash
python3 /home/oai/skills/endnote-fieldcode-converter/scripts/endnote_converter.py -i manuscript.docx -o manuscript_EndNote.docx
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

If the conversion fails because the references heading is not found, ask the user whether the heading is named differently, or rerun after editing a copy of the DOCX to use a `References` heading.

## Important limitations

- The parser is heuristic. It handles common biomedical numeric reference formats but cannot guarantee perfect CSL/EndNote metadata extraction from every journal style.
- EndNote must be installed in Word for the user to fully update, format, and manage the resulting fields.
- The script creates a traveling-library style XML payload from the parsed reference text. It does not connect to the user's local EndNote library.
- CrossRef lookup can improve DOI and metadata fields, but it depends on internet access and may fail silently for ambiguous titles.
- Complex citations inside Word text boxes, comments, footnotes, headers, or track-changes markup may require additional OOXML handling.

## Troubleshooting

Consult `references/usage.md` for command examples and expected manuscript structure.
Consult `references/validation.md` for post-conversion checks and common failure modes.

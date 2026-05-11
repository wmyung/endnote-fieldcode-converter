# EndNote Field Code Converter

A ChatGPT Skill and standalone Python utility for converting Microsoft Word manuscripts with plain-text numeric citations into EndNote-ready `.docx` files.

The converter replaces ordinary bracketed citations such as `[1]`, `[1,2]`, and `[1-3]` with Microsoft Word field-code sequences using `ADDIN EN.CITE`. It also converts the numbered reference list into an `ADDIN EN.REFLIST` bibliography field. The output is designed to be opened in Microsoft Word with EndNote installed, so the references can be updated, formatted, and managed through EndNote.

## What this project does

This project helps when a manuscript contains references only as plain text but needs to be converted into a Word document that behaves like an EndNote-linked manuscript.

Typical use cases include:

- converting journal manuscripts from plain-text Vancouver-style citations to EndNote-compatible fields;
- recovering an EndNote-like traveling library from a numbered reference list;
- preparing biomedical or scientific manuscripts for citation updating and bibliography formatting in EndNote;
- performing fast offline conversion without DOI lookup;
- optionally enriching parsed references with CrossRef DOI metadata.

## Repository contents

```text
.
├── LICENSE
├── README.md
├── RELEASE_NOTES.md
├── PUBLISH_TO_GITHUB.md
├── skill.zip
└── endnote-fieldcode-converter/
    ├── LICENSE
    ├── SKILL.md
    ├── agents/
    │   └── openai.yaml
    ├── references/
    │   ├── usage.md
    │   └── validation.md
    ├── requirements.txt
    └── scripts/
        └── endnote_converter.py
```

## Installation

### Use as a ChatGPT Skill

1. Download `skill.zip` from this repository.
2. Open the ChatGPT Skills page.
3. Upload `skill.zip`.
4. Use the Skill on a `.docx` manuscript that contains bracketed numeric citations and a numbered reference list.

### Use as a standalone command-line script

Clone the repository and install the Python dependency:

```bash
git clone https://github.com/YOUR-USERNAME/endnote-fieldcode-converter.git
cd endnote-fieldcode-converter
python3 -m pip install -r endnote-fieldcode-converter/requirements.txt
```

Run the converter offline without CrossRef lookup:

```bash
python3 endnote-fieldcode-converter/scripts/endnote_converter.py \
  -i manuscript.docx \
  -o manuscript_EndNote.docx \
  --skip-crossref
```

Run with CrossRef metadata lookup:

```bash
python3 endnote-fieldcode-converter/scripts/endnote_converter.py \
  -i manuscript.docx \
  -o manuscript_EndNote.docx
```

Run a dry check without writing an output file:

```bash
python3 endnote-fieldcode-converter/scripts/endnote_converter.py \
  -i manuscript.docx \
  --dry-run \
  --skip-crossref
```

## Supported citation patterns

The converter supports common bracketed numeric citation forms:

```text
[1]
[1,2]
[1, 2]
[1-3]
[1,3-5]
```

By default, the script converts only citation numbers that are found in the parsed reference list. This reduces accidental conversion of bracketed numbers that are not citations, such as years or scale scores.

To force conversion of all numeric bracket patterns, use:

```bash
--convert-unknown-citations
```

## Expected manuscript format

The manuscript should include a references section with a clear heading such as `References`, `Reference`, or `Bibliography`.

Example:

```text
References
[1] Smith J, Doe A. Example article title. Example Journal. 2020;12(3):45-50.
[2] Kim H, Lee S. Another article. Journal Name. 2021;7:100-110.
```

In-text citations should appear before the reference section:

```text
Previous studies reported similar findings [1]. The evidence is mixed [1,2] and has expanded recently [3-5].
```

## Output

The output `.docx` preserves visible numeric citations while replacing their underlying Word XML with EndNote field-code structures:

```text
w:fldChar begin
w:instrText ADDIN EN.CITE ...
w:fldChar separate
visible citation text, for example [1]
w:fldChar end
```

The reference list is replaced by an `ADDIN EN.REFLIST` field unless `--keep-reference-text` is used.

## Command-line options

| Option | Description |
|---|---|
| `-i`, `--input` | Input DOCX path |
| `-o`, `--output` | Output DOCX path. Defaults to `input_EndNote.docx` |
| `--skip-crossref` | Skip CrossRef DOI lookup. This is the fastest fully offline mode |
| `--offline` | Alias for `--skip-crossref` |
| `--dry-run` | Parse and validate without writing output |
| `--convert-unknown-citations` | Convert numeric brackets even if the citation number is not found in the reference list |
| `--keep-reference-text` | Keep the original plain-text reference entries after inserting the EndNote bibliography field |

## Validation checklist

After conversion, inspect the console output. A successful conversion should report:

- references found and parsed;
- converted citation count greater than zero;
- balanced Word field markers;
- output file written.

Treat these as issues requiring manual review:

- `Field markers balanced: no`
- `Parsed references: 0`
- `Converted citations: 0` when the manuscript visibly contains `[N]` citations
- many `Unparsed reference` warnings
- many `Skipped citation` warnings for numbers that should exist in the reference list

## Limitations

- The reference parser is heuristic and may not fully parse every journal style.
- EndNote must be installed in Microsoft Word to fully update, format, and manage the resulting fields.
- The script creates a traveling-library style XML payload from parsed reference text. It does not connect to a user's local EndNote library.
- CrossRef lookup may improve DOI and metadata fields, but it requires internet access and can fail for ambiguous titles.
- Complex citations inside Word text boxes, comments, footnotes, headers, or tracked changes may require additional OOXML handling.

## Suggested GitHub repository description

```text
ChatGPT Skill and Python script for converting plain-text numeric citations in DOCX manuscripts into EndNote field codes.
```

## Suggested GitHub topics

```text
endnote, docx, word, references, citations, manuscript, biomedical-writing, chatgpt-skill, ooxml, crossref
```

## License

This project is released under the MIT License. See [LICENSE](LICENSE).

## Keywords / Search Terms

`endnote` `zotero` `citation-manager` `reference-manager` `docx` `microsoft-word`
`traveling-library` `addin-en-cite` `addin-en-reflist` `field-code`
`crossref-api` `doi-lookup` `manuscript` `academic-writing` `biomedical`
`vancouver-style` `citation-converter` `plain-text-citations`
`zotero-to-endnote` `word-field-codes`

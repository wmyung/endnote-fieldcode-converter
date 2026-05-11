# Usage guide

## Minimal conversion

```bash
python3 scripts/endnote_converter.py -i manuscript.docx -o manuscript_EndNote.docx --skip-crossref
```

## Conversion with CrossRef DOI lookup

```bash
python3 scripts/endnote_converter.py -i manuscript.docx -o manuscript_EndNote.docx
```

## Dry run

```bash
python3 scripts/endnote_converter.py -i manuscript.docx --dry-run --skip-crossref
```

## Expected document structure

The document should contain a visible references heading followed by numbered references:

```text
References
[1] Smith J, Doe A. Example article title. Example Journal. 2020;12(3):45-50.
[2] Kim H, Lee S. Another article. Journal Name. 2021;7:100-110.
```

In-text citations should use bracketed numeric style:

```text
Previous studies reported similar findings [1]. The evidence is mixed [1,2] and has expanded recently [3-5].
```

## Output behavior

The output Word file preserves visible numeric citations while replacing the underlying text with Word field-code sequences:

- `w:fldChar begin`
- `w:instrText ADDIN EN.CITE ...`
- `w:fldChar separate`
- visible display text such as `[1]`
- `w:fldChar end`

The reference list is replaced by an `ADDIN EN.REFLIST` field unless `--keep-reference-text` is used.

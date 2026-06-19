# BJOG EndNote display-formatting notes

Use this when a DOCX already contains EndNote field codes and the user asks to make it BJOG style while preserving EndNote editability.

## Session-derived behavior

For BJOG/Vancouver-style manuscript output, keep Word field-code structure intact and change only display text/runs:

- Preserve `ADDIN EN.CITE` fields.
- Preserve `ADDIN EN.REFLIST` field.
- Change in-text visible citations from bracketed numeric display, e.g. `[1,2]`, to superscript numeric display, e.g. `1,2`.
- Change reference-list visible labels from `[1]` to `1.`.
- Move periods/commas immediately after a citation field to before the superscript field, e.g. `risk [1,2].` -> `risk.^1,2` in Word display.
- Keep the document zip-valid and verify balanced field markers.

## Reusable script

Run:

```bash
python3 scripts/format_bjog_endnote_docx.py \
  -i manuscript_EndNote.docx \
  -o manuscript_EndNote_BJOG_formatted.docx
```

Expected verification output includes:

- `zip_test_bad_file: None`
- `balanced: True`
- unchanged or expected `EN.CITE` / `EN.REFLIST` counts
- `field_markers` with equal begin/separate/end counts

## Pitfalls

- Do not regenerate citations from plain text if EndNote fields already exist; modify field display runs instead.
- Do not remove `w:instrText` XML payloads; they contain EndNote traveling-library data.
- If using a derived `.ens` style instead of official `BJOG.ens`, state the caveat separately; do not imply it is official.

# Release notes

## v0.1.0

Initial public release under the MIT License.

### Features

- Convert plain-text numeric citations in DOCX manuscripts into `ADDIN EN.CITE` Word field codes.
- Convert numbered reference lists into an `ADDIN EN.REFLIST` bibliography field.
- Support `[1]`, `[1,2]`, `[1, 2]`, `[1-3]`, and `[1,3-5]` citation patterns.
- Support offline conversion with `--skip-crossref` or `--offline`.
- Support optional CrossRef DOI and metadata enrichment.
- Provide dry-run mode for preflight parsing.
- Validate Word field marker balance after conversion.
- Skip unknown numeric brackets by default to reduce false-positive citation conversion.

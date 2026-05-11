# Validation guide

## Successful conversion indicators

A successful run should report:

- references found and parsed,
- converted citation count greater than zero,
- balanced Word field markers,
- output file written.

## Common warnings

### References heading not found

The script searches paragraph text for `references`, `reference`, or `bibliography`. If the document uses another label, edit a copy of the DOCX or adjust the script invocation after modifying the heading.

### Unparsed reference

The reference text did not match the heuristic parser. The script still creates a fallback EndNote record using the full reference text as a generic title. Review these records in Word/EndNote after conversion.

### Skipped citation

A numeric bracket was detected before the reference section, but the number was not found in the parsed reference list. By default, the script leaves that bracket unchanged to reduce false positives. Use `--convert-unknown-citations` only when the user wants all numeric brackets converted.

## Manual Word checks

Open the output DOCX in Microsoft Word with EndNote installed, then:

1. Toggle field codes if needed to confirm `ADDIN EN.CITE` exists.
2. Use EndNote's Update Citations and Bibliography command.
3. Check that the visible numbering and bibliography order match the manuscript.
4. Save a separate copy before submitting to a journal.

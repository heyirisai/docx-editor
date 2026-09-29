---
'@eigenpal/docx-editor-core': patch
---

Treat a literal percent on table and cell widths (`w:w="100%"` with `w:type="pct"`) as that percentage of the parent. OOXML `pct` is fiftieths of a percent, and parsing `"100%"` as `100` laid the table out at 2% of the text column.

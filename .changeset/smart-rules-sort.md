---
'@eigenpal/docx-editor-core': patch
---

Tables keep rows wrapped in a row-level content control or customXml (w:tbl > w:sdt > w:tr); they were dropped on open and on save.

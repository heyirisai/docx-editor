---
'@eigenpal/docx-editor-core': patch
---

Repeating table header rows (w:tblHeader) now also repeat above the continuation of a body row that broke across a page boundary, matching Word: a header row repeats on every page that shows part of the table, not only when the break falls between rows.

---
'@eigenpal/docx-editor-core': patch
---

A table with repeating header rows (w:tblHeader) no longer strands a clipped header row at the bottom of a page: as in Word, the header row is never split across a page break, and the whole table moves to the next page when the header plus the first body row does not fit.

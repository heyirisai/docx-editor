---
'@eigenpal/docx-editor-core': patch
---

When the rows of a table-level content control or customXml wrapper are no longer next to each other, each run of rows is saved inside its own copy of the wrapper (keeping its tag, alias and lock) instead of the later rows losing it. Reopened content controls leave out the duplicate id.

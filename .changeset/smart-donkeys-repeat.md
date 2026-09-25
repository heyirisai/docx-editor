---
'@eigenpal/docx-editor-core': minor
---

Support legacy Word form fields (FORMDROPDOWN, FORMCHECKBOX, FORMTEXT) as first-class content controls: they are discovered by `findContentControls` with `source: 'legacy'`, edited with `setContentControlValue`, and clickable in the paged editor. The `w:ffData` block round-trips verbatim. Adds `findGlyphCheckboxes` for read-only detection of typed/Wingdings checkbox glyphs.

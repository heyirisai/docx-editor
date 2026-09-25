---
'@eigenpal/docx-editor-core': patch
---

Legacy form fields with no stored result (FORMDROPDOWN, FORMTEXT, FORMCHECKBOX) now display what Word shows - the current list entry, the text default or blank, the checkbox glyph - instead of rendering empty, and findContentControls reports that text. The synthesized display is not written back on save, so an untouched field still round-trips byte for byte.

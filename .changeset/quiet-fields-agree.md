---
'@eigenpal/docx-editor-core': patch
---

Legacy form-field robustness: `w:ffData` patches follow the namespace prefix the file itself uses (a document binding WordprocessingML to another prefix, or the default namespace, no longer loses dropdown/checkbox edits on save); a projected `FORMTEXT` keeps its `legacyFormField.value`/`hasResult` in step when its content is replaced through `setContentControlContent` or by typing in the editor; and `findGlyphCheckboxes` reports every box character in a run — `☐ Yes ☐ No` typed as one run yields two candidates — with a new `offset` for each.

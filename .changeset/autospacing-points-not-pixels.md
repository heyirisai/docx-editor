---
'@eigenpal/docx-editor-core': patch
---

Fix automatic paragraph spacing (`w:beforeAutospacing`/`w:afterAutospacing`) being applied as 14 pixels instead of Word's 14 points, which made every such paragraph 25% short. On a cover page built from empty spacer paragraphs the error compounded and walked the artwork up the page.

---
'@eigenpal/docx-editor-core': patch
---

Numbered lists number like Word: each level starts at its own start value (including 0, or the value a "Restart at 1" / "Set numbering value" override gives it), and a list that opens with sub-items or skips a level gives the missing parent level that value, so the sub-items read 1.1., 1.2. and the next top-level item is 2. instead of repeating or restarting at 1.

---
'@eigenpal/docx-editor-core': patch
---

Numbered lists number like Word: each level starts at its own start value (including 0), and a list that opens with sub-items or skips a level gives the missing parent level its start value, so the sub-items read 1.1., 1.2. and the next top-level item is 2. A "Restart at 1" or "Set numbering value" override applies when its own level is first used or restarted, so a restarted list that opens with sub-items prints 1. again at its first top-level item, as Word does.

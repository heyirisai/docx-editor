---
'@eigenpal/docx-editor-core': patch
---

Multi-level numbered lists that open with sub-items or skip a level now number like Word: the missing parent level takes its start value, so the sub-items read 1.1., 1.2. and the next top-level item is 2. instead of restarting at 1.

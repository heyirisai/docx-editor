---
'@eigenpal/docx-editor-core': patch
---

Pagination parity with Word for Arial / Times New Roman documents and tall footers. Single line spacing for Arial and Times New Roman now uses Word's line pitch (1.1499 × font size — 11pt Arial → 12.65pt) instead of a value ~3% short, so Arial documents no longer paginate a page late; a line holding only an anchored drawing is sized from its paragraph mark like Word; and the footer band is anchored with its bottom at the `w:footer` distance and grows upward, so footers paint where Word paints them and a tall footer shortens the body page by `footer distance + footer height`.

---
'@eigenpal/docx-editor-core': minor
---

API: `TableCell.content` is now `(Paragraph | Table | BlockSdt)[]` and `TableCell` gains an optional `rowWrappers`. Code that walks cell content as "not a paragraph, so a table" must handle `BlockSdt` (a block-level content control) too.

Fix fixed-layout tables whose rows cover different grid columns (w:gridBefore / w:gridAfter): header cells no longer shift onto hairline spacer columns and collapse after the first edit, and gridBefore/gridAfter/wBefore/wAfter now survive a save.

Fix missing spaces between runs with horizontal text scaling (w:w), e.g. "Customer Records" painting as "CustomerRecords": scaled runs are now measured and painted at their scaled width (their client rects match the layout, with no overlap into neighbouring runs), character spacing (w:spacing) applies after every character as in Word, and a word split across runs no longer wraps mid-word.

Block-level content controls placed directly in a table cell (e.g. Yes/No drop-downs in questionnaire grids) are now parsed, rendered, saved unchanged, and found/answered by the content-control APIs instead of being dropped. So are content controls that wrap a whole cell at row level (`w:tr > w:sdt > w:sdtContent > w:tc`, also inside a row-level `w:customXml`): the wrapped cell used to vanish with all its paragraphs, leaving a short row; it is now a real grid cell whose content is the control (`TableCell.rowWrappers` records the wrapper), and a save re-emits the wrapper around the cell unchanged. A row-level `w:customXml` wrapper is only written back if its captured tags are exactly a `w:customXml` start tag with an optional `w:customXmlPr` (otherwise the wrapper is dropped and the cell kept), and `rowWrappers` stays out of the shared collaboration document, like a control's raw `w:sdtPr`.

Answering a content control (`setContentControlValue` / `setContentControlContent`, headless or in the editor) now keeps the `w14:paraId`, `w14:textId` and paragraph properties of the paragraphs it rewrites, as Word does, so answer anchors keyed by paraId still resolve after a save. When an answer adds paragraphs, the first keeps the original id and the new ones get fresh ids unique in the document.

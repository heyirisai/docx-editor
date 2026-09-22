# Test Fixtures

This directory contains DOCX test fixtures for the Playwright test suite.

## Files

### empty.docx

An empty DOCX document with default Word settings.
Used for testing baseline document state.

### styled-content.docx

A document containing styled content:

- Bold, italic, and underlined text
- Different font sizes
- Paragraph with alignment variations
- Mixed formatting

### with-tables.docx

A document containing tables:

- Simple 3x3 table
- Table with merged cells
- Table with formatted content

### complex-styles.docx

A document with complex styling:

- Custom styles
- Theme colors
- Headers/footers
- Multiple sections

### wrap-none-positioned-image-demo.docx

A synthetic document containing a positioned image anchored with `wp:wrapNone`.
Used to reproduce anchored images that should paint independently without adding
paragraph flow height or text-wrap margins.

### wrap-none-two-seals-title-box-demo.docx

A synthetic title-page document containing two `behindDoc` `wp:wrapNone` images
aligned to the left and right margins around a centered title box. Used to
reproduce multiple non-wrapping anchored images on the same page.

### image-layout-modes-demo.docx

A synthetic document containing exactly one image of each of Word's three core
layout modes — `wp:inline` (in-line), `wp:wrapSquare` (wrap-around float), and
`wp:wrapTopAndBottom` (full-width block) — in a single document. Used by
`e2e/tests/image-layout-modes.spec.ts` to lock in correct rendering of all
three paths side by side.

### issue-472-floating-textbox.docx

A synthetic document containing a `wps:wsp` text box in a `wp:anchor` with
`wp:wrapSquare wrapText="bothSides"`. Used to reproduce issue #472 without
committing the private original document.

### footnote-bottom-overflow.docx

A synthetic document containing dense bottom-of-page footnote references with
long citation-like note text. Used to verify that final footnote reservation and
painted footnote height agree so notes do not run off the page.

### footnote-overlap-regression.docx

A synthetic document containing neutral filler paragraphs with dense short
footnote references around page boundaries. Used to verify that references in
split paragraph fragments are mapped to the page containing their rendered
line, so the painted footnote area does not overlap body text.

### endnotes-tracked-changes.docx

A synthetic document with two body endnote references, separator and
continuation-separator endnotes, and a normal endnote whose body contains a
tracked insertion (`w:ins`). Used to verify the note-body serializer round-trip
(separators + `w:endnoteRef` survive repack) and `getChanges({ includeEndnotes })`.

### empty-table-row-vmerge.docx

A synthetic document containing a table whose middle row is made entirely of
`w:vMerge` continuation cells. Used to verify that DOCX import does not produce
an invalid empty ProseMirror `tableRow`.

### table-cell-selection-drag.docx

A synthetic document containing a simple table with generic sample text. Used to
verify precise drag selection inside a single table cell without promoting the
selection to the whole cell.

### table-header-orphan-rfp.docx

A synthetic RFP-style document whose narrative filler pushes a 7-column table
with a repeating header row (`w:trPr/w:tblHeader`) to within one line of the
page bottom; 15 body rows of 2–4 lines each, rows allowed to split. Used to
verify Word's rule that a header row never sits alone (or clipped) at a page
bottom — the whole table moves to the next page instead.

### table-header-orphan-cantsplit.docx

The same 7-column repeating-header matrix with `w:cantSplit` on every row. Used
to verify that a body row that does not fit moves whole to the next page while
the header still never strands, and that no row is clipped at either boundary.

### table-header-orphan-proposal.docx

A proposal-style narrative document with a 2-column repeating-header table
embedded in long prose (the proposal-template corpus, as opposed to the
RFP/questionnaire matrices above). Used to verify the same header rules hold
when the table is narrow and surrounded by body text.

### table-row-split-header.docx

A synthetic RFP-style document whose 4-column table has a repeating header row
(`w:trPr/w:tblHeader`) and a body row far taller than a page, so the row splits
mid-content (no `w:cantSplit`). Used to verify that the remainder resumes at the
top of the printable area on the next page, below a repeated header — per
ECMA-376 §17.4.78 a header row repeats on every page that displays part of the
table, including one that only shows the tail of a split row.

### table-row-split-noheader.docx

The proposal-style counterpart: a headerless 2-column table inside narrative
prose whose middle row is taller than a page. Used to verify that a split row's
continuation starts exactly at the printable top (nothing painted in the top
margin) and that no line is lost or painted twice across the break.

Both are generated by `bun scripts/create-table-row-split-fixtures.mjs`.

### toc-hyperlink-tabs.docx

A synthetic document with one TOC1 paragraph wrapping
`1[tab]Introduction[tab]5` inside a `<w:hyperlink>`, plus the TOC1 style with
its right-aligned dot-leader tab stop. Used to verify that tabs inside
hyperlinks survive parsing and that TOC entries render with dot leaders and
right-aligned page numbers like Word.

### inline-checkbox-controls.docx

A synthetic document with inline Word checkbox content controls
(`w14:checkbox`) covering unchecked, checked, untagged, locked, data-bound, and
plain-text fallback cases. Used to verify that docx-editor paints the existing
Word glyph in flow while exposing only unlocked/unbound controls as clickable
editor widgets.

## Generating Fixtures

To regenerate fixtures, run:

```bash
bun run e2e/fixtures/generate-fixtures.ts
bun scripts/create-issue-472-floating-textbox-fixture.mjs
bun scripts/create-footnote-bottom-overflow-fixture.mjs
bun scripts/create-footnote-overlap-regression-fixture.mjs
bun scripts/create-empty-table-row-vmerge-fixture.mjs
bun scripts/create-table-cell-selection-drag-fixture.mjs
bun scripts/create-table-header-pagination-fixtures.mjs
bun scripts/create-toc-hyperlink-fixture.mjs
bun scripts/create-inline-checkbox-controls-fixture.mjs e2e/fixtures/inline-checkbox-controls.docx
```

Or manually create them using Microsoft Word or another DOCX editor.

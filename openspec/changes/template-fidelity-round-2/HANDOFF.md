# Template fidelity, round 2 — handoff

Branch `fix/per-section-headers-and-exact-rows`, on top of commit `9a77c32f`.
Everything below is **uncommitted** in the working tree.

Goal: render every template in `~/Documents/Templates` the way Word does,
fixing causes in the engine rather than per file. Round 1 (`9a77c32f`) did
sections / columns / shading. This round started from two the user named:
**COMET** (table problems, page 2's TOC sitting on page 1) and
**Word-Doc-Iris-Proposal-Template** (the art box beside the summary missing).

---

## How things were decided

MS Word 16.x is the oracle, driven headlessly (AppleScript → PDF →
`pdftoppm`/`pdftotext`). Each claim below was measured, not inferred; several
were established with synthetic probe `.docx` files and by editing a real
template and re-rendering it in Word. The harness and its gotchas are written
up in the session memory notes `template-conformance-harness` and
`word-anchor-and-header-band-rules`.

---

## Done

1. **A float with no room beside it reserves a full-width band.**
   `clampFloatingWrapMargins` used to report "no exclusion" for a float as
   wide as the content area, so text ran under it. Word pushes the line
   below, exactly as for `topAndBottom`. This is the root cause of the COMET
   cover: the first-page header's full-page picture displaces the header's
   own text, which is what grows the header band and pushes the body off
   page one.
   _Files:_ `layout-bridge/measuring/measureParagraph.ts` (the clamp),
   `floatingZones.ts`, `measureBlocksPipeline.ts`.

2. **Per-page `w:titlePg` margins.** `extendMarginsForHeaderFooter` took
   `max(header, firstHeader)` for a whole section, so (1) collapsed every
   page of COMET's first section to a 24px strip. Bands are now split into
   `rest` / `first`; the paginator carries `firstPageMargins` and applies
   them only to a section's opening page. A cover with no content area is
   allowed (bounded by the sheet) and behaves like Word: the first block is
   parked off the bottom of that sheet, the body starts overleaf.
   _Files:_ `layout-bridge/headerFooterMargins.ts`,
   `layout-engine/{paginator,index,types,section-breaks}.ts`,
   `editor/computeLayout.ts`.

3. **Shapes and text boxes inside `wpg:wgp` groups paint.** Only pictures did,
   so the "VelocityEHS Advantage" panel in the Iris template was missing and
   the body text beside it ran full width. New `docx/groupFrame.ts` holds the
   child-coordinate mapping, shared with `groupPreview.ts`.
   _Files:_ `docx/groupFrame.ts` (new), `docx/groupPreview.ts`,
   `docx/textBoxParser.ts`, `docx/blockContentParser.ts`.

4. **Stroke-only connectors paint as a line.** `prst="line"` / `wps:cNvCnPr`
   with `cy="0"` is how a footer rule is authored; round 1 lifted it as a
   filled shape, so it painted a full-width bordered box across the bottom of
   every page of COMET and the Iris template. New `Shape.lineShape`
   (`'down' | 'up'`, from `a:xfrm/@flipV`) plumbed to the painter, which
   strokes one rotated edge.

5. **`w:ptab` (absolute-position tab, §17.3.3.19)** parsed, laid out and
   written back. COMET's footers use centre+right `w:ptab` to push the page
   number to the right margin; unparsed, it read "Hilb Group 4".
   _Files:_ `docx/runParser.ts`, `types/content/run.ts`,
   `prosemirror/utils/tabCalculator.ts` (`positionalTabStop`), the PM
   `TabExtension`, `toFlowBlocks/runs.ts`, `measureParagraph.ts`,
   `renderParagraph/line.ts`, `serializer/runSerializer.ts`.

6. **Table style cell defaults + run-property precedence.**
   A table style's own `w:tcPr` (§17.7.6.8) now supplies the cell default for
   the whole table — `vAlign` included, which is what centres COMET's merged
   label column. Its run properties now sit between the document defaults and
   the paragraph style (§17.7.2); above the paragraph style they repainted
   body text white-on-white in a one-row table stale-flagged
   `w:cnfStyle firstRow`. Verified against Word by recolouring the
   conditional and observing that nothing changed where `Normal` set a colour.
   _Files:_ `toProseDoc/tables.ts`, `toProseDoc/paragraph.ts`,
   `prosemirror/styles/styleResolver.ts` (new `ownRunFormatting`).

7. **Floats anchored inside a table cell exclude that cell's text.**
   `measureTableBlock` takes a `measureCellBlocks` callback (both adapters
   pass their float-aware `measureBlocks`), and `renderTable` no longer
   re-measures cells with a second, painter-local zone model. COMET's bio
   photos used to sit on top of the names.

8. **`floatSkipBefore` is `padding-top`, not `margin-top`.** A body fragment
   is absolutely positioned (its own BFC) but a fragment inside a cell is
   only `position: relative`, so the margin collapsed out through the cell
   and moved the whole cell — picture included — instead of the text.

9. **A hanging-indent first line reserves room for its list marker.** The
   marker width was exempted on that path, so a bulleted first line measured
   `hanging` px wider than it painted and the last word was clipped. Most
   visible in a narrow table cell.

10. **Codex review fixes** (see below) and docs: `docs/site/data/word-features.ts`
    (shapes, groups, borders-shading) and `docs/site/content/word-fidelity.mdx`.

### Tests added

- `docx/__tests__/connector-and-group-shapes.test.ts`
- `docx/__tests__/positional-tab.test.ts`
- `prosemirror/conversion/__tests__/table-style-cell-defaults.test.ts`
- new cases in `layout-bridge/__tests__/headerFooterMargins.test.ts` and
  `measuring/__tests__/clampFloatingWrapMargins.test.ts`

---

## Verification

|                                                                          |                                                                                                        |
| ------------------------------------------------------------------------ | ------------------------------------------------------------------------------------------------------ |
| typecheck                                                                | 6/6 clean                                                                                              |
| lint                                                                     | 0 errors, 132 pre-existing warnings                                                                    |
| `bun test`                                                               | 2100 pass / 0 fail                                                                                     |
| Playwright chromium                                                      | `table` 72 pass, `header\|footer\|image` 66 pass, broad grep 144 pass — 0 failures                     |
| `api:check`                                                              | exit 0                                                                                                 |
| parity / css-thin / public-docs / editor-contract / export-parity / i18n | all pass                                                                                               |
| round-trip                                                               | COMET and the Iris template re-pack with byte-identical element counts to HEAD (checked in a worktree) |

Page counts vs Word (`was` = before this round):

```
comet    word 22  was 20  now 21   IMPROVED
nnn      word 13  was 10  now 12   IMPROVED
ideagen  word 22  was 25  now 26   (font drift; pages 1-7 pixel-identical)
cority draftv1 harmonic highered sentinelone sqeleave  exact, unchanged
goodwin irisprop lcps sqeprop  unchanged
```

Per-page ink vs Word moved materially on exactly three pages, all closer:
COMET p2 0.0571→0.0004 (Word 0.0002), p3 0.2608→0.0571 (0.0447),
p4 0.0642→0.2608 (0.2608). No page in any template got materially worse.

---

## Codex review

Run on the uncommitted diff. Five findings; four addressed:

- **P1 — zero-capacity cover swallows the first fragment.** _Kept by design._
  Word does exactly this: on COMET it parks the first "Table of Contents"
  off the bottom of the cover sheet and starts page 2 with the next
  paragraph, which our render now matches to ~10px. Only a section's FIRST
  page can be degenerate, and exactly one fragment is swallowed before the
  paginator advances, so it terminates. **Still worth adding a paginator
  regression test for this** (started, not finished — see below).
- **P2 — document defaults outranked the table style.** _Fixed._
  `resolveParagraphStyle` now also returns `ownRunFormatting` (the style
  chain without the doc defaults) so the table style slots between them.
- **P2 — first-page margins on an even/odd parity padding page.** _Fixed._
  New `paginator.restampSectionFirstPage(state, isFirst)` re-applies the
  page's margins when the section-first stamp moves; `section-breaks.ts`
  uses it on both the demotion and the claim.
- **P2 — `w:ptab` measured against a different width than it painted.**
  _Fixed._ The painter now receives `contentWidthPx` / `indentRightPx`
  unreduced, matching `measureParagraph`.
- **P2 — children of an ALIGNED group align independently, losing their
  relative offsets.** _Documented, not fixed._ Correct for the common case
  (a panel and the bar across its top, which share the group's box); wrong
  for children at different offsets. A real fix needs the anchor model to
  carry an alignment AND a delta, which `ImagePosition` has no field for.
  Noted in `textBoxParser.ts` and in the feature matrix.

---

## Round 3 — what the user reported, and what fixed it

Four complaints: ideagen "has a lot of issues"; COMET table rows with images
moving down and no border showing; the Iris proposal template gaining blank
pages; its footer misaligned, content missing on the last page, and a button
with no corner radius.

1. **§17.6.22 — a section break is governed by the section being ENTERED.**
   `w:type` says how ITS OWN section starts, and an absent type is `nextPage`.
   The dispatcher already read the next section's type but fell back to the
   outgoing break's when it was absent, so COMET's run of three `continuous`
   sections swallowed the page break that opens the Compensation divider and
   the whole sheet vanished. Measured with a three-section probe through Word
   (`sectPr0=nextPage, sectPr1=continuous` keeps section 1 on section 0's page;
   flip them and every section gets its own page).
   _Files:_ `layout-engine/index.ts`.

2. **A paragraph that is nothing but a page break.** It was folded into
   `pageBreakBefore`, which moved its own empty line to the TOP of the next
   page — where it both showed as a blank line and satisfied the following
   paragraph's `pageBreakBefore`, costing an extra sheet each time. Word keeps
   that line on the page it is already on. The break now becomes a standalone
   block, and `fromProseDoc` folds a `pageBreak` back into the paragraph it
   came from, so the round trip stops gaining a paragraph per hard break
   (better than the `<w:pageBreakBefore/>` substitute it used to write).
   _Files:_ `toProseDoc/paragraph.ts`, `toProseDoc.ts`, `fromProseDoc.ts`.

3. **`w:tblBorders` cascades PER SIDE.** A table that writes only
   `top/left/bottom/right = none` still inherits `insideH`/`insideV` from its
   style — COMET's bio tables are Table Grid with the outer box switched off,
   and Word draws the black rule between the photo column and the text.
   _Files:_ `toProseDoc/tables.ts` (`mergeTableBorders`).

4. **The painter re-measured a column fragment at the page width.** With any
   float on the page the painter re-measures each paragraph (only it knows the
   fragment's real Y). It passed the page content width, so in a multi-column
   section the lines were re-broken across the whole page and painted over each
   other — ideagen's two-column "Core Capabilities" block. It now measures at
   `fragment.width`, with the zone margins mapped into the fragment's own
   coordinate space.
   _Files:_ `layout-painter/renderPage.ts` (`zonesForFragment`).

5. **`a:prstGeom` presets are drawn.** Every shape painted as a rectangle, so
   COMET's "15 years of experience" roundel was square and the Iris back
   cover's "EXPLORE OUR PLATFORM" pill had hard corners. New
   `Shape.geometry` / `Shape.cornerAdj` (`ellipse` | `roundRect`, adjust as a
   fraction of the short side) parse, paint as `border-radius` and serialize
   back.

6. **Tables inside `w:txbxContent`.** The element is `EG_BlockLevelElts`; the
   parser passed `null` for the table parser with a "most text boxes don't
   have tables" comment, so the Iris template's "PROOF POINT" panel painted as
   an empty frame. `ShapeBlockContent = Paragraph | Table` now runs through
   parse, PM, flow blocks, measure and paint. Text-box measurement moved into
   core (`layout-bridge/measureTextBox.ts`) so React and Vue cannot drift.

7. **§17.6.11 — `w:footer` is the distance to the footer's BOTTOM edge.** The
   band was anchored by its top, putting every footer one line-height too low;
   in the Iris template the page-anchored icon beside "EHS.COM" ended up on
   its own line above the text. Measured in Word at `w:footer` =
   200/331/720 twips: the last line's bottom lands exactly that far above the
   page edge each time.
   _Files:_ `layout-painter/renderPage.ts`, `layout-bridge/headerFooterMargins.ts`.

8. **A header float with `behindDoc="0"` paints above the footer.** COMET's
   divider pages are a full-bleed header picture whose white right half is what
   hides the footer in Word; painted in DOM order, our footer rule and page
   number sat on top of the artwork. HF images now take the same
   `headerFooterFrontZIndex` band the HF text boxes already used.

9. **`a:lumMod` / `a:lumOff` resolve.** These are what Word's colour picker
   writes for every "Lighter 40% / Darker 25%" theme variant — far more common
   on shapes than `a:tint`/`a:shade`, and ignored. ideagen's "Core
   Capabilities" panel (accent5 at 95% luminance) painted pure white.
   _Files:_ `types/colors.ts`, `docx/drawingUtils.ts`, `utils/colorResolver.ts`.

10. **Wingdings 0xD8 is the arrowhead bullet `➢`**, not a fallback dot. Added
    it and the neighbouring private-use glyphs to `SYMBOL_BULLET_MAP`, and text
    boxes now get the same bullet normalisation the body does.

11. **A picture's own `a:prstGeom` is applied** (§20.1.9.18). Word stores a
    circular headshot as a normal rectangular `wp:extent` plus
    `<a:prstGeom prst="ellipse"/>` on `pic:spPr`, and paints the PRESET rather
    than the frame. Only the shape parser read `a:prstGeom`, so every one of
    these painted as a hard-edged square. The parser now reads the preset for
    pictures too (shared `parsePresetGeometry`, moved from `textBoxParser` to
    `drawingUtils`), it rides through PM / the layout bridge, the painter emits
    a `border-radius` on the `<img>` — or on the crop wrapper, in a header or
    footer — and the serializer writes it back instead of the hardcoded
    `prst="rect"`. Fixes COMET's 10 bio headshots, GOODWIN's 11, and ideagen's
    5 `flowChartConnector` circles.
    _Files:_ `docx/drawingUtils.ts`, `docx/imageParser.ts`,
    `docx/textBoxParser.ts`, `types/content/image.ts`,
    `prosemirror/schema/nodes.ts`, `extensions/nodes/ImageExtension.ts`,
    `conversion/toProseDoc/runs.ts`, `conversion/fromProseDoc/runs.ts`,
    `serializer/runSerializer/drawing.ts`, `layout-engine/types.ts`,
    `layout-bridge/toFlowBlocks{,/runs}.ts`, `layout-painter/renderImage.ts`,
    `floatingImageLayer.ts`, `renderPage.ts`, `renderPage/headerFooter.ts`.

12. **A floating picture inside a TABLE CELL lost its crop and opacity.**
    `CellFloatingImage` re-packed the run into a narrower record and dropped
    `srcRect` / `a:alphaModFix` — the same omission the header/footer path
    already carries a comment about. COMET's headshots are cropped 12.5% top
    and bottom, so they painted as the whole source squashed into a square.
    The record now forwards the picture's visual attrs.
    _File:_ `layout-painter/renderTableCellFloating.ts`.

13. **An anchored picture in a table cell is no longer clipped to the cell —
    nor moved.** Three separate faults on one object, all visible on LCPS's
    header logo, which lost its final letter and the mark drawn above it:
    - `extractCellFloatingImages` clamped the anchor into the cell content box
      (`Math.max(0, Math.min(x, contentWidth - width))`). `layoutInCell="1"`
      says the anchor is RESOLVED against the cell, not that the object is
      confined to it — measured with a probe (`cellclip.docx`): a 2.5in
      picture anchored in a 1in cell paints in full, across its neighbours and
      below the row. The clamp threw away the logo's `posOffset="-114300"`
      (-12px) and pushed it 12px right, into the table's clip edge. Removing
      it puts the logo at x 601..742 — Word's ink box to the pixel.
    - The cell float LAYER carried `overflow: hidden` and was sized to the
      cell's content box; it is now a bare coordinate origin that clips
      nothing, and it hangs off the ROW instead of the cell (the cell must
      keep clipping its own content: over-tall content and v-merge
      continuation slices depend on it).
    - The TABLE element clips only when the fragment is a page-break WINDOW.
      A fragment showing the whole table has no window to enforce, and
      clipping there cost real content.
      _Files:_ `layout-painter/renderTableCellFloating.ts`,
      `layout-painter/renderTable.ts`, `layout-painter/floatingImageLayer.ts`;
      e2e selector updated in `float-text-wrapping.spec.ts`.

### Page counts after round 3 (ours vs Word)

```
cority       8 /  8   exact        comet      22 / 22   EXACT (was 20)
draftv1      4 /  4   exact        irisprop   11 / 11   EXACT (was 13)
highered     3 /  3   exact        harmonic   12 / 12   exact
sqeleave     4 /  4   exact        sentinelone 2 /  2   exact
nnn         12 / 13                sqeprop    13 / 12
goodwin     21 / 19   Trade Gothic missing
lcps        72 / 70   Open Sans missing
ideagen     28 / 22   Gilroy missing
```

Seven of thirteen exact, up from five. Nothing regressed. Every remaining gap
is a template whose brand font is not installed — irisprop is exact because it
is set in Arial.

LCPS's reference was re-exported: the original PDF was 62 pages only because
Word was showing markup, which SCALES the page (comment balloons take a column
of the sheet) — geometry can't be read from an export like that. With
`show revisions and comments` off it is 70 pages. Check this before trusting
any Word PDF from a document that carries comments or tracked changes.

### How close are we to Word? (measured 2026-09-21)

Ink-overlap (IoU) between our render and Word's, both rasterised at 96 dpi, on
the 8 templates whose page count already matches so pages align 1:1 (66 pages).
Three granularities: 1px is glyph-exact, 8px is roughly word placement, 32px is
roughly paragraph/block placement.

```
template      pg      1px     8px    32px   pages >=90% @32px   fonts missing
highered       3    95.4%   98.4%   99.0%   3/3                 none
irisprop      11    30.7%   63.4%   86.6%   5/11                Proxima Nova, Ballinger
sentinelone    2    54.3%   69.8%   85.4%   1/2                 none  (see defect below)
cority         8    54.1%   68.1%   84.6%   5/8                 Fakt Pro, Aptos Display
sqeleave       4    10.8%   40.8%   80.6%   0/4                 Inter, Neue Haas Grotesk
harmonic      12    28.1%   46.6%   79.8%   3/12                Carlito, DengXian
comet         22    39.6%   56.4%   72.4%   7/22                Trade Gothic
draftv1        4    40.9%   54.4%   49.3%   0/4                 Trade Gothic
ALL           66    39.1%   58.4%   78.3%   24/66
```

The one template whose fonts are all installed (`highered`) is at **99%**. That
is the ceiling the engine currently reaches, and it is the strongest evidence
we have that font substitution — not layout logic — is what separates us from
Word on everything else. `sentinelone` also has every font and only scores 85%,
because of a real defect (below).

**Both Word PDFs of documents carrying comments were markup-scaled** and are
useless for geometry: Word shrinks the sheet to make room for balloons. Re-export
with `show revisions and comments of view of active window` set to false. LCPS
was 62 pages that way and is really 70; irisprop is 11 either way. Only those two
templates carry comments, so the rest of `conform/word/*.pdf` is sound.

### Known residuals

- **ideagen** is structurally right now (columns, panels, theme tints) but
  runs 6 pages long purely on text height: its body font is **Gilroy**, which
  is neither installed nor on Google Fonts. Everything downstream — banners
  landing on the wrong page, `behindDoc="0"` bands covering headings — follows
  from that drift, not from a layout bug. (Verified: Word hides text under a
  `behindDoc="0"` shape exactly as we do.)
- ~~**SentinelOne's cover text box does not render**~~ — FIXED, and the
  diagnosis above was wrong. The parser does return the text box; it reaches
  the DOM with the right text and geometry. It was being COVERED by the header's
  full-bleed cover picture (8.1in x 10.9in, opaque white on the right), which
  paints over the body because `.layout-page-header` is a later sibling of
  `.layout-page-content`. See "Round 5" below. Reading `elementsFromPoint` was
  what settled it: the span was the top hit-test target, because the thing
  covering it is `pointer-events: none`.
- **The Iris template's "Driving Success Together" heading** wraps into the
  33px channel beside the PROOF POINT box; Word puts it below. A faithful
  reproduction of that anchor geometry in Word wraps the same way we do, so
  the difference is somewhere in how Word sizes or places that particular box,
  not in the wrap rule (Word's minimum wrap channel is 24px = 0.25in,
  measured, which is what `MIN_WRAP_SEGMENT_WIDTH` already uses).

---

## Round 4 — the Tier 1 corpus runner (`scripts/corpus/tier1.ts`)

`bun run corpus:tier1 -- <dir>` round-trips every `.docx` under a path and
asserts seven things that must hold for EVERY document, with no Word, no
browser and no human judgement: `parse`, `serialize`,
`serialize-deterministic`, `repack`, `reparse`, `text-preserved` and
`roundtrip-stable`. Roughly a dozen documents a second, so a corpus of
thousands is practical. `--json out.json` writes a machine-readable report;
`--baseline b.json` exits 1 only on a regression against it, which is the
shape CI wants (the absolute pass rate on a real corpus will never be 100%).

`roundtrip-stable` is the one that earns its keep: it serializes the model,
repacks it, parses THAT, and serializes again. Any difference means every save
mutates the file — the drift that turns a template into garbage after a dozen
edits.

The first run over `~/Documents/Templates` failed 7/13 on `text-preserved` and
5/13 on `roundtrip-stable`. All are now 13/13. What it found:

| Defect                                                   | Where                                 | Effect before                                                                                                                                                 |
| -------------------------------------------------------- | ------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Block `w:sdt` in a `w:tc` dropped                        | `docx/tableParser.ts`                 | SQE's locked "GET IN TOUCH" CTA box vanished from the render AND the saved file                                                                               |
| `w:csTheme` written instead of `w:cstheme`               | `serializer/runSerializer.ts`         | CT_Fonts spells it all-lowercase; Word ignored the attribute and our own parser lost it on the 2nd save (GOODWIN, NNN)                                        |
| `a:schemeClr val="text1"`                                | `serializer/runSerializer/drawing.ts` | WordprocessingML slot names written into DrawingML; re-read as `dk1`, so ideagen's black cover panel repainted itself each save                               |
| `a:tint`/`a:shade` in the wrong unit                     | same                                  | a hex 0–255 modifier written where 1000ths-of-a-percent belong                                                                                                |
| Gradient / pattern / picture fill destroyed              | `docx/preservedShapeXml.ts`           | `parseFill` flattens every gradient to `{type:'gradient'}` with no stops, so the fill was written as nothing — and `<a:noFill/>` on the save after that (NNN) |
| `w:commentReference` carrier run kept AND re-synthesized | `paragraphParser/content.ts`          | one junk empty run added per comment per save (LCPS, Iris)                                                                                                    |

Two notes on the runner itself, both learned the hard way:

- `text-preserved` must strip `<mc:Fallback>` before comparing. A cover page is
  authored twice — `mc:Choice` (DrawingML) and `mc:Fallback` (legacy VML) — and
  counting the fallback copy scored Higher Education 86.7% for doing exactly
  what Word does.
- A bare percentage is not actionable. Every missing chunk is narrowed to the
  longest substring that is genuinely absent and reported; an empty narrowing
  means the chunk boundary, not the content, was the problem.

### Still open after round 4

- **Recursion depth on nested `w:sdt`** is unbounded in both the body path
  (`blockContentParser.parseBlockSdt`) and the new cell path
  (`tableParser.parseCellSdt`). CLAUDE.md asks for a cap; fix both together or
  neither, since asymmetric limits are worse than none.
- **Codex review did not run** on this change — the workspace is out of
  credits ("ERROR: Your workspace is out of credits"). Self-reviewed against
  the CLAUDE.md sink list instead; the sink grep is clean and the one judgement
  call (allowing `a:blipFill` through `preservedSpPrExtra`) is argued in that
  file's doc comment.
- **Tier 2 and Tier 3** are not built. Tier 2 = the same corpus in a browser,
  asserting layout invariants without an oracle (no overlapping line boxes, no
  content outside the page box, page count stable across two layouts). Tier 3 =
  the Word/LibreOffice oracle already scripted in `scripts/conformance/`.

---

## Round 5 — the header/footer z-band, measured

Two cover pages were missing their text: SentinelOne's "Project / Client /
Delivered On" block and Hilb's "EMPLOYEE BENEFITS ... RFP FOR:". Both were in
the DOM, correctly positioned, black on white, and invisible.

The old rule was "an in-front header/footer float beats all body content",
introduced from ONE measurement (a COMET divider page where Word's full-bleed
header picture hides the footer rule and page number). That generalised from
header-vs-footer to header-vs-body, and header-vs-body is the opposite.

**The probe** (`scratchpad/zprobe/`, built with Python `zipfile` on a real
template's `styles.xml`): a header picture, `behindDoc="0"`,
`relativeHeight="9000000"`, over a plain body paragraph, a body text box with
`relativeHeight="100"`, and footer text. Word's PDF export:

| Object                               | vs the header picture |
| ------------------------------------ | --------------------- |
| body paragraph text                  | **above** it          |
| body text box (`relativeHeight` 100) | **above** it          |
| footer text                          | **below** it          |

So `relativeHeight` orders objects only WITHIN a story, and the header/footer
story as a whole paints below the body story. `behindDoc` orders an object
against its own story's text. The bands in `layout-engine/zOrder.ts` now encode
exactly that, bottom to top:

    HF_BEHIND_Z (-1)  <  HF flow content (auto)  <  HF front floats
                      <  BODY_CONTENT_Z (1e9)   <  PAGE_OVERLAY_Z

`.layout-page-content` carries `BODY_CONTENT_Z`, which also makes it a stacking
context — body floats use raw OOXML `relativeHeight` as a z-index and can no
longer reach the header/footer band by being authored high. `HF_BEHIND_Z` is
safe because `applyPageStyles` puts both `isolation: isolate` and the page
background on `.layout-page`; move either and behind-doc objects fall behind
the page.

Verified after the change: SentinelOne and Hilb cover text visible and matching
Word, COMET's divider page still hides its footer behind the artwork.

### Note for whoever touches this next

The lesson is the one the old comment already half-recorded: a z-order rule
derived from one pair of stories does not transfer to another pair. Probe both
directions before generalising.

---

## Left to do

1. **Tests for round 3.** Arsam asked for fixes first and tests at the end, so
   only three files were added (`docx/__tests__/textbox-table-and-geometry`,
   the border-cascade block in `table-style-cell-defaults`, and
   `layout-painter/__tests__/column-fragment-remeasure`). Still owed: the
   §17.6.22 break-type cases beyond the three in
   `layout-engine/integration/section-breaks.test.ts`, the footer-band
   geometry, `a:lumMod`/`a:lumOff`, the HF float z-band, the Wingdings bullet,
   and the P1 zero-capacity paginator case carried over from round 2
   (harness: `layout-engine/__tests__/force-page-break-empty.test.ts`).
   Round 3 items 11-13 also owe tests: the picture `a:prstGeom` parse →
   paint → save path, the cell-float visual-attr forwarding, and the
   cell-anchored float position/clipping rules.
2. **`! bun changeset`** — interactive, hand-writing `.changeset/*.md` is
   forbidden. Not run for `9a77c32f`, `23cc0970` or `afd81a31` either, so this
   branch still needs exactly one.
3. **Tests for round 4**: the cell `w:sdt` parse → PM → flow → serialize path,
   the `w:cstheme` spelling, the DrawingML scheme-colour inverse map, the
   preserved gradient fill, and the comment-reference carrier rule (including
   the guard that keeps a reference with no matching `commentRangeEnd`).
4. **Aligned-group child offsets** (Codex P2 above) if it shows up in a real
   template.

### Bigger items still open (unchanged from round 1)

- **Font substitution** is now the dominant residual: every template whose
  body font is unavailable drifts (lcps +10 on Open Sans, goodwin +2 on
  Trade Gothic, ideagen +6 on Gilroy), every template whose font resolves
  is exact or ±1. Real OS/2 ratios already extracted from Word's cloud-font
  cache: Open Sans 1.3618, Nunito Sans 1.4290, Poppins 1.7620,
  Titillium Web 1.5210, Segoe UI 1.3301, Aptos Display 1.2847,
  Trade Gothic Next Cond 1.1880.
- **Zero-click font fetch (security).** `getGoogleFontEquivalent` in
  `utils/fontLoader.ts` ends `|| trimmed`, so opening a DOCX sends its own
  brand-font names to Google Fonts with no user action — recorded live on
  GOODWINTEMPLATE. Violates the "no zero-click external fetch" rule in
  CLAUDE.md. Fix with an allowlist; folds naturally into the font work.
- **Per-section header/footer EDITING** — `useHeaderFooterEditing.ts` still
  resolves one pair for the document, so double-clicking a header on a later
  page edits a different section's. Needs React + Vue.
- **Vue divergence** — harmonic renders 13 pages in Vue vs 12 in React and
  Word; `issue-764` / `issue-777` fail on pre-change sources; `issue-781` is
  flaky (2 pass / 4 fail across six runs either way).

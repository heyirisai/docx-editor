/**
 * Table pagination with repeating header rows (`w:trPr/w:tblHeader`).
 *
 * Word's rules, which these suites lock in:
 *
 * 1. A header row never sits alone at the bottom of a page, and is never
 *    itself broken across a page boundary. If the header row(s) plus at
 *    least the first line of the first body row do not fit in the space
 *    left, the whole table starts on the next page.
 * 2. A repeated header on a continuation page renders at FULL height (never
 *    clipped), and the first body row of that page starts below it.
 * 3. A body row that does not fit either moves whole to the next page
 *    (`w:cantSplit`) or breaks at a whole-line boundary — in both cases the
 *    visible band of every row is accounted for in the fragment height, so
 *    no cell content is cut off.
 *
 * Fixtures mirror the two corpora the editor has to serve: RFP /
 * questionnaire tables (wide, 7 columns, many short body rows) and
 * proposal-style narrative documents (a 2-column table with a repeating
 * header embedded in long prose).
 */

import { describe, test, expect } from 'bun:test';

import { layoutDocument } from '../index';
import type {
  Layout,
  ParagraphBlock,
  ParagraphMeasure,
  TableBlock,
  TableFragment,
  TableMeasure,
} from '../types';
import { makeLayoutOptions } from './helpers';

const LINE = 20;
const EPS = 0.001;

// ---------------------------------------------------------------------------
// Synthetic fixture builders
// ---------------------------------------------------------------------------

function para(id: number, lines = 1): ParagraphBlock {
  return {
    kind: 'paragraph',
    id,
    runs: [{ kind: 'text', text: 'x'.repeat(lines * 10) }],
  } as unknown as ParagraphBlock;
}

function paraMeasure(lines: number): ParagraphMeasure {
  return {
    kind: 'paragraph',
    lines: Array.from({ length: lines }, () => ({
      fromRun: 0,
      fromChar: 0,
      toRun: 0,
      toChar: 0,
      width: 10,
      ascent: LINE * 0.8,
      descent: LINE * 0.2,
      lineHeight: LINE,
    })),
    totalHeight: lines * LINE,
  };
}

type RowSpec = {
  /** Text lines in every cell of the row (row height = lines * LINE). */
  lines: number;
  isHeader?: boolean;
  cantSplit?: boolean;
};

/**
 * Build a table block + measure from a row spec. Every cell holds one
 * paragraph of `lines` lines so the row-break geometry has real interior
 * line boundaries (that is what lets a row — or, before the fix, a header
 * row — split mid-content).
 */
function buildTable(
  id: number,
  columnCount: number,
  specs: RowSpec[],
  totalWidth = 624
): { block: TableBlock; measure: TableMeasure } {
  const colWidth = totalWidth / columnCount;
  const columnWidths = Array.from({ length: columnCount }, () => colWidth);

  const block = {
    kind: 'table',
    id,
    columnWidths,
    rows: specs.map((spec, r) => ({
      id: id + (r + 1) * 100,
      isHeader: spec.isHeader,
      cantSplit: spec.cantSplit,
      cells: Array.from({ length: columnCount }, (_, c) => ({
        id: id + (r + 1) * 100 + c + 1,
        blocks: [para(id + (r + 1) * 100 + c + 1, spec.lines)],
      })),
    })),
  } as unknown as TableBlock;

  const measure: TableMeasure = {
    kind: 'table',
    columnWidths,
    totalWidth,
    totalHeight: specs.reduce((sum, s) => sum + s.lines * LINE, 0),
    rows: specs.map((spec) => ({
      height: spec.lines * LINE,
      cells: Array.from({ length: columnCount }, () => ({
        blocks: [paraMeasure(spec.lines)],
        width: colWidth,
        height: spec.lines * LINE,
      })),
    })),
  };

  return { block, measure };
}

/** ~15 body rows of 2–4 lines each, preceded by a 2-line header row. */
function bodySpecs(count: number, cantSplit = false): RowSpec[] {
  return Array.from({ length: count }, (_, i) => ({
    lines: 2 + (i % 3),
    cantSplit: cantSplit || undefined,
  }));
}

// ---------------------------------------------------------------------------
// Assertions
// ---------------------------------------------------------------------------

type PlacedFragment = { pageIndex: number; frag: TableFragment };

function tableFragments(layout: Layout, blockId: number): PlacedFragment[] {
  const out: PlacedFragment[] = [];
  layout.pages.forEach((page: Layout['pages'][number], pageIndex: number) => {
    for (const frag of page.fragments) {
      if (frag.kind === 'table' && frag.blockId === blockId) {
        out.push({ pageIndex, frag: frag as TableFragment });
      }
    }
  });
  return out;
}

function countHeaderRows(block: TableBlock): number {
  let n = 0;
  for (const row of block.rows) {
    if (row.isHeader) n++;
    else break;
  }
  return n;
}

/**
 * Assert Word's header/pagination contract over every fragment of a table.
 */
function expectWordHeaderPagination(
  layout: Layout,
  block: TableBlock,
  measure: TableMeasure
): PlacedFragment[] {
  const frags = tableFragments(layout, block.id as unknown as number);
  expect(frags.length).toBeGreaterThan(0);

  const headerRowCount = countHeaderRows(block);
  const rowH = measure.rows.map((r) => r.height);
  const headerHeight = rowH.slice(0, headerRowCount).reduce((a, b) => a + b, 0);

  for (const { pageIndex, frag } of frags) {
    // (1) A header row is never cut at a page bottom.
    if (frag.bottomClip !== undefined) {
      expect(frag.toRow - 1).toBeGreaterThanOrEqual(headerRowCount);
    }

    // (2) A fragment never resumes in the middle of a header row.
    if (frag.topClip !== undefined && frag.topClip > 0) {
      expect(frag.fromRow).toBeGreaterThanOrEqual(headerRowCount);
    }

    // (3) No page ends with a header-row-only fragment while the table
    //     continues on the next page.
    if (frag.continuesOnNext && headerRowCount > 0) {
      expect(frag.toRow).toBeGreaterThan(frag.fromRow + (frag.fromRow === 0 ? headerRowCount : 0));
      if (frag.fromRow === 0) {
        expect(frag.toRow).toBeGreaterThan(headerRowCount);
      }
    }

    // (4) A continuation that starts at a row boundary repeats the FULL
    //     header and carries at least one body row beneath it.
    if (frag.continuesFromPrev && !frag.topClip && headerRowCount > 0) {
      expect(frag.headerRowCount).toBe(headerRowCount);
      expect(frag.toRow).toBeGreaterThan(frag.fromRow);
    }

    // (5) No clipping: the fragment is tall enough for the repeated header
    //     plus every visible row band it claims to paint.
    const overhead = frag.headerRowCount ? headerHeight : 0;
    let content = 0;
    for (let r = frag.fromRow; r < frag.toRow; r++) {
      let h = rowH[r];
      if (r === frag.fromRow && frag.topClip) h -= frag.topClip;
      if (r === frag.toRow - 1 && frag.bottomClip !== undefined) h -= rowH[r] - frag.bottomClip;
      content += h;
    }
    expect(frag.height).toBeGreaterThanOrEqual(overhead + content - EPS);

    // (6) A fragment that could fit on an empty page must not overflow the
    //     page's content area (an overflowing fragment renders clipped).
    const page = layout.pages[pageIndex];
    const contentBottom = page.size.h - page.margins.bottom - (page.footnoteReservedHeight ?? 0);
    const capacity = contentBottom - page.margins.top;
    if (frag.height <= capacity) {
      expect(frag.y + frag.height).toBeLessThanOrEqual(contentBottom + 0.5);
    }
  }

  return frags;
}

/** Rows covered by fragments must tile the table exactly once, in order. */
function expectRowsCoveredOnce(frags: PlacedFragment[], rowCount: number): void {
  let cursor = 0;
  for (const { frag } of frags) {
    expect(frag.fromRow).toBe(cursor);
    cursor = frag.bottomClip !== undefined ? frag.toRow - 1 : frag.toRow;
  }
  expect(cursor).toBe(rowCount);
}

// ---------------------------------------------------------------------------
// Suites
// ---------------------------------------------------------------------------

const options = makeLayoutOptions();
const CONTENT_HEIGHT = options.pageSize.h - options.margins.top - options.margins.bottom;

/** Filler paragraph that leaves `freeLines` lines of space on page 1. */
function filler(freeLines: number): { block: ParagraphBlock; measure: ParagraphMeasure } {
  const lines = Math.floor((CONTENT_HEIGHT - freeLines * LINE) / LINE);
  return { block: para(1, lines), measure: paraMeasure(lines) };
}

describe('Layout engine — repeating table header pagination (RFP/questionnaire corpus)', () => {
  test('a 7-column table whose 2-line header lands one line above the page bottom moves whole to the next page', () => {
    const f = filler(1);
    const { block, measure } = buildTable(1000, 7, [
      { lines: 2, isHeader: true },
      ...bodySpecs(15),
    ]);

    const layout = layoutDocument([f.block, block], [f.measure, measure], options);
    const frags = expectWordHeaderPagination(layout, block, measure);

    // The table does not start on page 1 at all — the header alone would be
    // the only thing that fit there.
    expect(frags[0].pageIndex).toBe(1);
    expect(frags[0].frag.fromRow).toBe(0);
    expect(frags[0].frag.toRow).toBeGreaterThan(1);
    expectRowsCoveredOnce(frags, measure.rows.length);
  });

  test('the same table with `w:cantSplit` body rows keeps every row whole and never strands the header', () => {
    const f = filler(1);
    const { block, measure } = buildTable(2000, 7, [
      { lines: 2, isHeader: true, cantSplit: true },
      ...bodySpecs(15, true),
    ]);

    const layout = layoutDocument([f.block, block], [f.measure, measure], options);
    const frags = expectWordHeaderPagination(layout, block, measure);

    for (const { frag } of frags) {
      expect(frag.bottomClip).toBeUndefined();
      expect(frag.topClip).toBeUndefined();
    }
    expect(frags[0].pageIndex).toBe(1);
    expectRowsCoveredOnce(frags, measure.rows.length);
  });

  test('the header never strands for any amount of leftover space on page 1', () => {
    for (let freeLines = 0; freeLines <= 6; freeLines++) {
      const f = filler(freeLines);
      const { block, measure } = buildTable(3000 + freeLines, 7, [
        { lines: 2, isHeader: true },
        ...bodySpecs(15),
      ]);
      const layout = layoutDocument([f.block, block], [f.measure, measure], options);
      const frags = expectWordHeaderPagination(layout, block, measure);
      expectRowsCoveredOnce(frags, measure.rows.length);
    }
  });

  test('continuation pages repeat the full header above the first body row', () => {
    const f = filler(1);
    const { block, measure } = buildTable(4000, 7, [
      { lines: 2, isHeader: true },
      ...bodySpecs(15),
    ]);
    const layout = layoutDocument([f.block, block], [f.measure, measure], options);
    const frags = expectWordHeaderPagination(layout, block, measure);

    const continuations = frags.filter(({ frag }) => frag.continuesFromPrev && !frag.topClip);
    expect(continuations.length).toBeGreaterThan(0);
    for (const { frag } of continuations) {
      // Full header height is reserved on top of the body rows.
      const headerHeight = measure.rows[0].height;
      const bodyHeight = frag.height - headerHeight;
      expect(frag.headerRowCount).toBe(1);
      expect(bodyHeight).toBeGreaterThanOrEqual(measure.rows[frag.fromRow].height - EPS);
    }
  });
});

describe('Layout engine — repeating table header pagination (proposal-style corpus)', () => {
  test('a 2-column table with a repeating header inside long narrative prose paginates like Word', () => {
    // Narrative before the table leaves one line of room; narrative after it
    // continues the document, as in a proposal template.
    const intro = filler(1);
    const { block, measure } = buildTable(5000, 2, [
      { lines: 2, isHeader: true },
      ...bodySpecs(15),
    ]);
    const outroLines = 30;
    const outro = para(9000, outroLines);

    const layout = layoutDocument(
      [intro.block, block, outro],
      [intro.measure, measure, paraMeasure(outroLines)],
      options
    );
    const frags = expectWordHeaderPagination(layout, block, measure);

    expect(frags[0].pageIndex).toBe(1);
    expect(frags[0].frag.toRow).toBeGreaterThan(1);
    expectRowsCoveredOnce(frags, measure.rows.length);
  });

  test('a tall body row still splits mid-content, and the split band is never clipped', () => {
    // One body row taller than a page forces a mid-row break.
    const { block, measure } = buildTable(6000, 2, [
      { lines: 2, isHeader: true },
      { lines: Math.ceil((CONTENT_HEIGHT * 1.6) / LINE) },
      { lines: 3 },
    ]);
    const layout = layoutDocument([block], [measure], options);
    const frags = expectWordHeaderPagination(layout, block, measure);

    expect(frags.some(({ frag }) => frag.bottomClip !== undefined)).toBe(true);
    expectRowsCoveredOnce(frags, measure.rows.length);
  });
});

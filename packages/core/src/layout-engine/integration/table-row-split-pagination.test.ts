/**
 * A table BODY row that splits across a page boundary (Word's "allow row to
 * break across pages" — the default when `w:cantSplit` is absent).
 *
 * Word's geometry for the continuation, which this suite locks in:
 *
 * 1. Nothing of the table is ever painted inside the page margin: every
 *    fragment starts at or below the printable top, and a continuation
 *    fragment starts exactly AT the printable top.
 * 2. The remainder of the split row resumes at the top of the printable
 *    area — below the repeated header row(s) when the table has any. Per
 *    ECMA-376 §17.4.78, `w:tblHeader` rows repeat "at the top of each new
 *    page on which part of this table is displayed"; a page that shows the
 *    tail of a split row displays part of the table, so the header repeats
 *    there too (it is not suppressed by the mid-row break).
 * 3. A fragment's height is exactly the repeated-header height plus the
 *    visible band of every row it paints — never more (it would overflow the
 *    page) and never less (the painter's window would clip live content).
 * 4. The split row's two halves tile the row exactly: the continuation
 *    resumes at the offset where the previous page stopped, so no line is
 *    lost and none is painted twice.
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
// Synthetic fixtures (mirrors of e2e/fixtures/table-row-split-*.docx)
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

type RowSpec = { lines: number; isHeader?: boolean };

/** Table whose every cell holds one paragraph, so rows have real line boundaries. */
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
 * Assert Word's split-row geometry across every fragment of one table.
 * Returns the fragments so callers can add case-specific expectations.
 */
function expectWordSplitRowGeometry(
  layout: Layout,
  block: TableBlock,
  measure: TableMeasure
): PlacedFragment[] {
  const frags = tableFragments(layout, block.id as unknown as number);
  expect(frags.length).toBeGreaterThan(1);

  const headerRowCount = countHeaderRows(block);
  const rowH = measure.rows.map((r) => r.height);
  const headerHeight = rowH.slice(0, headerRowCount).reduce((a, b) => a + b, 0);
  /** Painted band per row, summed over the fragments that show it. */
  const painted = new Array(rowH.length).fill(0);

  frags.forEach(({ pageIndex, frag }, i) => {
    const page = layout.pages[pageIndex];
    const printableTop = page.margins.top;
    const contentBottom = page.size.h - page.margins.bottom - (page.footnoteReservedHeight ?? 0);
    const overhead = frag.headerRowCount ? headerHeight : 0;

    // (1) Nothing is painted inside the top margin (where the running header
    //     lives): the fragment's own origin is never above the printable top.
    expect(frag.y).toBeGreaterThanOrEqual(printableTop - EPS);

    // (2) A continuation starts exactly AT the printable top, and its first
    //     painted body line sits at the printable top plus the repeated
    //     header height (the repeated header occupies [y, y + overhead)).
    if (frag.continuesFromPrev) {
      expect(frag.y).toBeCloseTo(printableTop, 6);
      const firstBodyLineTop = frag.y + overhead;
      expect(firstBodyLineTop).toBeCloseTo(printableTop + overhead, 6);
      // A table with header rows repeats them on every continuation page,
      // including one that resumes a row broken mid-content.
      if (headerRowCount > 0) {
        expect(frag.headerRowCount).toBe(headerRowCount);
        expect(overhead).toBe(headerHeight);
      }
    }

    // (3) The fragment's height is exactly its repeated header plus the
    //     visible band of every row it paints — no painted content exceeds
    //     the allotted clip, and no claimed space goes unpainted.
    let content = 0;
    for (let r = frag.fromRow; r < frag.toRow; r++) {
      let band = rowH[r];
      if (r === frag.fromRow && frag.topClip) band -= frag.topClip;
      if (r === frag.toRow - 1 && frag.bottomClip !== undefined) band -= rowH[r] - frag.bottomClip;
      expect(band).toBeGreaterThan(0);
      painted[r] += band;
      content += band;
    }
    expect(frag.height).toBeCloseTo(overhead + content, 6);

    // A fragment that fits an empty page must stay inside the printable area.
    if (frag.height <= contentBottom - printableTop) {
      expect(frag.y + frag.height).toBeLessThanOrEqual(contentBottom + 0.5);
    }

    // (4) A split row resumes exactly where the previous fragment stopped.
    const prev = frags[i - 1]?.frag;
    if (prev?.bottomClip !== undefined) {
      expect(frag.fromRow).toBe(prev.toRow - 1);
      expect(frag.topClip ?? 0).toBeCloseTo(prev.bottomClip, 6);
    } else if (prev) {
      expect(frag.fromRow).toBe(prev.toRow);
      expect(frag.topClip).toBeUndefined();
    }
  });

  // (4, cont.) Every row is painted exactly once in total — no lost lines,
  // no duplicated lines across the boundary.
  rowH.forEach((h, r) => expect(painted[r]).toBeCloseTo(h, 6));

  return frags;
}

// ---------------------------------------------------------------------------
// Suites
// ---------------------------------------------------------------------------

const options = makeLayoutOptions();
const CONTENT_HEIGHT = options.pageSize.h - options.margins.top - options.margins.bottom;
/** Lines of a tall body row: 1.6 pages, so it must break across the boundary. */
const TALL_ROW_LINES = Math.ceil((CONTENT_HEIGHT * 1.6) / LINE);

describe('Layout engine — body row split across a page boundary (headerless)', () => {
  test('the continuation resumes at the printable top with nothing in the margin', () => {
    const { block, measure } = buildTable(7000, 2, [
      { lines: 3 },
      { lines: TALL_ROW_LINES },
      { lines: 3 },
    ]);

    const layout = layoutDocument([block], [measure], options);
    const frags = expectWordSplitRowGeometry(layout, block, measure);

    // Row 1 really did break mid-content.
    const split = frags.find(({ frag }) => frag.bottomClip !== undefined);
    expect(split).toBeDefined();
    expect(split?.frag.toRow).toBe(2);
    const continuation = frags.find(({ frag }) => (frag.topClip ?? 0) > 0);
    expect(continuation).toBeDefined();
    expect(continuation?.frag.fromRow).toBe(1);
    // Headerless table: no header band is reserved above the resumed row.
    expect(continuation?.frag.headerRowCount).toBeUndefined();
  });
});

describe('Layout engine — body row split under a repeating header row', () => {
  test('the continuation repeats the header and resumes the row below it', () => {
    const { block, measure } = buildTable(7100, 4, [
      { lines: 2, isHeader: true },
      { lines: TALL_ROW_LINES },
      { lines: 3 },
    ]);

    const layout = layoutDocument([block], [measure], options);
    const frags = expectWordSplitRowGeometry(layout, block, measure);

    const continuation = frags.find(({ frag }) => (frag.topClip ?? 0) > 0);
    expect(continuation).toBeDefined();
    // ECMA-376 §17.4.78: the header repeats on every page that displays part
    // of the table — a mid-row break does not suppress it.
    expect(continuation?.frag.headerRowCount).toBe(1);
    // The resumed body row starts one header height below the printable top.
    const page = layout.pages[continuation!.pageIndex];
    expect(continuation!.frag.y + measure.rows[0].height).toBeCloseTo(
      page.margins.top + measure.rows[0].height,
      6
    );
  });

  test('a row spanning three pages repeats the header on every continuation', () => {
    const veryTall = Math.ceil((CONTENT_HEIGHT * 2.4) / LINE);
    const { block, measure } = buildTable(7200, 4, [
      { lines: 2, isHeader: true },
      { lines: veryTall },
      { lines: 2 },
    ]);

    const layout = layoutDocument([block], [measure], options);
    const frags = expectWordSplitRowGeometry(layout, block, measure);

    expect(frags.length).toBeGreaterThanOrEqual(3);
    for (const { frag } of frags.slice(1)) {
      expect(frag.headerRowCount).toBe(1);
    }
  });

  test('no line is lost or painted twice when the split lands mid-page-1', () => {
    // Prose before the table so the split offset is not a whole page.
    const introLines = Math.floor(CONTENT_HEIGHT / LINE / 2);
    const intro = para(7300, introLines);
    const { block, measure } = buildTable(7400, 4, [
      { lines: 2, isHeader: true },
      { lines: TALL_ROW_LINES },
      { lines: 4 },
    ]);

    const layout = layoutDocument(
      [intro, block],
      [paraMeasure(introLines), measure],
      makeLayoutOptions()
    );
    expectWordSplitRowGeometry(layout, block, measure);
  });
});

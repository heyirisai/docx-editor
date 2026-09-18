/**
 * A full-bleed cover page, as real templates build one.
 *
 * The pattern is a one-cell table the width of the page with
 * `<w:trHeight w:hRule="exact">` set slightly TALLER than the page, inside a
 * section with zero margins — so the fill bleeds off the edge with no white
 * sliver. Three things went wrong with it:
 *
 *  - an exact-height row was broken across pages at a line boundary, which is
 *    not a thing a fixed-height row has;
 *  - the rows after it were carried to a fresh page, so a 0.05in accent bar
 *    got a page of its own;
 *  - the empty paragraph that carries the section's `sectPr` then demanded a
 *    page too, and got a blank one.
 */
import { describe, expect, test } from 'bun:test';
import { layoutDocument } from '../index';
import type {
  FlowBlock,
  ParagraphBlock,
  ParagraphMeasure,
  SectionBreakBlock,
  TableBlock,
  TableMeasure,
} from '../types';

const PAGE = { w: 816, h: 1056 };
const NO_MARGINS = { top: 0, right: 0, bottom: 0, left: 0 };

/** A one-column table whose rows are all `hRule="exact"`. */
function exactTable(id: string, heights: number[]): { block: TableBlock; measure: TableMeasure } {
  return {
    block: {
      kind: 'table',
      id,
      rows: heights.map((h, i) => ({
        id: `${id}-r${i}`,
        height: h,
        heightRule: 'exact' as const,
        cells: [{ id: `${id}-c${i}`, blocks: [], colSpan: 1, rowSpan: 1 }],
      })),
      columnWidths: [PAGE.w],
    } as unknown as TableBlock,
    measure: {
      kind: 'table',
      totalWidth: PAGE.w,
      totalHeight: heights.reduce((a, b) => a + b, 0),
      columnWidths: [PAGE.w],
      rows: heights.map((h) => ({ height: h, cells: [{ height: h, blocks: [], measures: [] }] })),
    } as unknown as TableMeasure,
  };
}

function para(id: string, height: number) {
  const block: ParagraphBlock = {
    kind: 'paragraph',
    id,
    pmStart: 0,
    pmEnd: 0,
    runs: [{ kind: 'text', text: id }],
    attrs: {},
  } as ParagraphBlock;
  const measure: ParagraphMeasure = {
    kind: 'paragraph',
    lines: [
      {
        fromRun: 0,
        fromChar: 0,
        toRun: 0,
        toChar: 0,
        width: 100,
        ascent: 10,
        descent: 3,
        lineHeight: height,
      },
    ],
    totalHeight: height,
  } as ParagraphMeasure;
  return { block, measure };
}

const run = (blocks: FlowBlock[], measures: unknown[]) =>
  layoutDocument(blocks, measures as never, {
    pageSize: PAGE,
    margins: NO_MARGINS,
    finalPageSize: PAGE,
    finalMargins: NO_MARGINS,
  });

describe('a fixed-height row taller than the page', () => {
  test('is clipped onto one page instead of broken across two', () => {
    // 1067px on a 1056px page — the deliberate 11px of bleed.
    const t = exactTable('cover', [1067]);
    const result = run([t.block], [t.measure]);

    expect(result.pages).toHaveLength(1);
    const frags = result.pages[0].fragments;
    expect(frags).toHaveLength(1);
    expect(frags[0].height).toBeGreaterThan(PAGE.h); // overflows, and the page clips it
    expect((frags[0] as { continuesOnNext?: boolean }).continuesOnNext).toBeFalsy();
  });

  test('takes the fixed-height rows after it with it', () => {
    // The trailing accent bar sits at an offset the tall row already decided —
    // off the page. It used to get a page of its own.
    const t = exactTable('cover', [1067, 5]);
    const result = run([t.block], [t.measure]);

    expect(result.pages).toHaveLength(1);
    expect(result.pages[0].fragments).toHaveLength(1);
    const frag = result.pages[0].fragments[0] as { fromRow: number; toRow: number };
    expect(frag.fromRow).toBe(0);
    expect(frag.toRow).toBe(2);
  });

  test('rows that each fit a page still paginate instead of being swallowed', () => {
    // Half a page of text, then two exact rows that each fit a page but not
    // together. Coalescing here would clip the second one away.
    const p = para('intro', 600);
    const t = exactTable('bands', [600, 600]);
    const result = run([p.block, t.block], [p.measure, t.measure]);

    const tableFrags = result.pages.flatMap((pg) =>
      pg.fragments.filter((f) => f.blockId === 'bands')
    );
    expect(tableFrags.length).toBeGreaterThan(1);
    expect(tableFrags.every((f) => f.height <= PAGE.h)).toBe(true);
    // Both rows are still placed somewhere.
    const last = tableFrags[tableFrags.length - 1] as { toRow: number };
    expect(last.toRow).toBe(2);
  });

  test('a row that merely does not fit HERE still moves on as a whole', () => {
    // Half a page of text, then an exact row taller than what is left but
    // shorter than a page: it belongs on the next page, unbroken.
    const p = para('intro', 600);
    const t = exactTable('band', [600]);
    const result = run([p.block, t.block], [p.measure, t.measure]);

    expect(result.pages).toHaveLength(2);
    expect(result.pages[1].fragments).toHaveLength(1);
    expect(result.pages[1].fragments[0].height).toBe(600);
  });
});

describe('w:titlePg is a section property', () => {
  test("the section's own first page is flagged, not just the document's", () => {
    const a = para('a', 100);
    const sb: SectionBreakBlock = {
      kind: 'sectionBreak',
      id: 'sb',
      type: 'nextPage',
      margins: NO_MARGINS,
    };
    const b = para('b', 1000);
    const c = para('c', 1000);
    const result = run(
      [a.block, sb, b.block, c.block],
      [a.measure, { kind: 'sectionBreak' }, b.measure, c.measure]
    );

    expect(result.pages[0].isSectionFirstPage).toBe(true);
    // First page of section 1 — page 2 of the document.
    expect(result.pages[1].sectionIndex).toBe(1);
    expect(result.pages[1].isSectionFirstPage).toBe(true);
    // Its continuation is not.
    expect(result.pages[2].isSectionFirstPage).toBe(false);
  });
});

describe('a parity break', () => {
  test('flags the page the section opens on, not a sheet before it', () => {
    // `forcePageBreak` is idempotent on an empty page, so today an `oddPage`
    // break from page 1 does not actually insert a blank page 2 — but whichever
    // page the section opens on is the one `w:titlePg` applies to, and no
    // earlier page may claim the role.
    const a = para('a', 100);
    const sb: SectionBreakBlock = {
      kind: 'sectionBreak',
      id: 'sb',
      type: 'oddPage',
      margins: NO_MARGINS,
    };
    const b = para('b', 100);
    const result = run([a.block, sb, b.block], [a.measure, { kind: 'sectionBreak' }, b.measure]);

    const opening = result.pages.find((p) => p.fragments.some((f) => f.blockId === 'b'));
    expect(opening?.isSectionFirstPage).toBe(true);
    expect(opening?.sectionIndex).toBe(1);
    // Exactly one page in that section carries the flag.
    const flagged = result.pages.filter((p) => p.sectionIndex === 1 && p.isSectionFirstPage);
    expect(flagged).toHaveLength(1);
  });
});

describe('pages know which section they belong to', () => {
  test('sectionIndex advances across a nextPage break', () => {
    const a = para('a', 100);
    const sb: SectionBreakBlock = {
      kind: 'sectionBreak',
      id: 'sb',
      type: 'nextPage',
      margins: NO_MARGINS,
    };
    const b = para('b', 100);
    const result = run([a.block, sb, b.block], [a.measure, { kind: 'sectionBreak' }, b.measure]);

    expect(result.pages).toHaveLength(2);
    expect(result.pages[0].sectionIndex ?? 0).toBe(0);
    expect(result.pages[1].sectionIndex).toBe(1);
  });

  test('a single-section document stays on section 0', () => {
    const a = para('a', 100);
    const b = para('b', 1000);
    const result = run([a.block, b.block], [a.measure, b.measure]);
    for (const page of result.pages) expect(page.sectionIndex ?? 0).toBe(0);
  });
});

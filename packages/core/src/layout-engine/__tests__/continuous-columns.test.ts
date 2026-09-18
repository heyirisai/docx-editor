/**
 * ECMA-376 §17.6.4 — a `continuous` section break that changes `w:cols`
 * starts a new column set on the SAME page. Word balances that set across its
 * columns and resumes full-width content below the tallest one.
 *
 * Before this, only the document's terminal section was balanced: a
 * mid-document two-column section stacked into column one, its fragments kept
 * the page's full content width (so centred headings centred across the page,
 * over the other column), and the content that followed carried on inside the
 * shortened region.
 */

import { describe, expect, test } from 'bun:test';
import { layoutDocument } from '../index';
import type {
  FlowBlock,
  Measure,
  ParagraphBlock,
  ParagraphMeasure,
  SectionBreakBlock,
} from '../types';

const PAGE = { w: 816, h: 1056 };
const MARGINS = { top: 96, right: 96, bottom: 96, left: 96 };
/** 816 - 96 - 96 */
const CONTENT_WIDTH = 624;

function para(id: string, height: number): { block: ParagraphBlock; measure: ParagraphMeasure } {
  return {
    block: {
      kind: 'paragraph',
      id,
      pmStart: 0,
      pmEnd: 0,
      runs: [{ kind: 'text', text: id }],
      attrs: {},
    },
    measure: {
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
    },
  };
}

/**
 * lead | [continuous break] 6 x 60px in 2 columns | [continuous break] tail.
 * The break block carries the properties of the section it CLOSES, so the
 * two-column config sits on the SECOND break.
 */
function twoColumnDocument(): { blocks: FlowBlock[]; measures: Measure[] } {
  const blocks: FlowBlock[] = [];
  const measures: Measure[] = [];
  const push = (p: { block: FlowBlock; measure: Measure }) => {
    blocks.push(p.block);
    measures.push(p.measure);
  };
  const breakBlock = (id: string, columns?: SectionBreakBlock['columns']) => {
    const sb: SectionBreakBlock = { kind: 'sectionBreak', id, type: 'continuous' };
    if (columns) sb.columns = columns;
    blocks.push(sb);
    measures.push(undefined as unknown as Measure);
  };

  push(para('lead', 100));
  breakBlock('sb0');
  for (let i = 0; i < 6; i++) push(para(`c${i}`, 60));
  breakBlock('sb1', { count: 2, gap: 24 });
  push(para('tail', 60));
  return { blocks, measures };
}

describe('continuous multi-column sections', () => {
  test('balances across columns, then resumes below the tallest', () => {
    const { blocks, measures } = twoColumnDocument();
    const layout = layoutDocument(blocks, measures, { pageSize: PAGE, margins: MARGINS });

    expect(layout.pages.length).toBe(1);
    const at = (id: string) => layout.pages[0].fragments.find((f) => f.blockId === id)!;

    // 6 x 60px over 2 columns balances 3/3 rather than filling column one.
    const columnX = [...new Set(['c0', 'c1', 'c2', 'c3', 'c4', 'c5'].map((id) => at(id).x))];
    expect(columnX.length).toBe(2);
    expect(['c0', 'c1', 'c2'].map((id) => at(id).x)).toEqual([96, 96, 96]);
    expect(['c3', 'c4', 'c5'].map((id) => at(id).x)).toEqual([420, 420, 420]);
    // The second column restarts at the region top, not the page top.
    expect(at('c3').y).toBe(at('c0').y);

    // Fragments span ONE column: (624 - 24) / 2.
    expect(at('c0').width).toBe(300);
    // ...while single-column content keeps the full content width.
    expect(at('lead').width).toBe(CONTENT_WIDTH);
    expect(at('tail').width).toBe(CONTENT_WIDTH);

    // The section that follows clears the tallest column and gets the rest of
    // the page back — it neither overlaps the columns nor spills to page 2.
    expect(at('tail').x).toBe(96);
    expect(at('tail').y).toBeGreaterThanOrEqual(at('c2').y + 60);
  });

  test('a section with no w:cols stays full width', () => {
    const blocks: FlowBlock[] = [];
    const measures: Measure[] = [];
    const a = para('a', 60);
    blocks.push(a.block);
    measures.push(a.measure);
    blocks.push({ kind: 'sectionBreak', id: 'sb', type: 'continuous' } as SectionBreakBlock);
    measures.push(undefined as unknown as Measure);
    const b = para('b', 60);
    blocks.push(b.block);
    measures.push(b.measure);

    const layout = layoutDocument(blocks, measures, { pageSize: PAGE, margins: MARGINS });
    for (const fragment of layout.pages[0].fragments) {
      expect(fragment.x).toBe(96);
      expect((fragment as { width: number }).width).toBe(CONTENT_WIDTH);
    }
  });
});

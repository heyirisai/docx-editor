/**
 * Paragraph spacing, and the one suppression rule Word actually has.
 *
 * `getSpacingBefore`/`getSpacingAfter` are deliberately dumb: a paragraph's
 * spacing applies whether or not it has text (ECMA-376 §17.3.1.33). The
 * suppression Word performs is `w:contextualSpacing` (§17.3.1.9), and it lives
 * in `applyContextualSpacing`, which runs over the block list before these
 * getters do. The `isEmptyParagraph` heuristic that used to sit here — zero
 * style-inherited spacing on an empty paragraph — was a third, unspecified rule
 * on top, and it collapsed the empty spacer paragraphs templates use to push a
 * cover block down the page.
 *
 * So the pair below is the real regression guard: spacing survives on an empty
 * paragraph, AND contextual spacing still suppresses what it is supposed to.
 */
import { describe, expect, test } from 'bun:test';
import { getSpacingAfter, getSpacingBefore } from './paragraphSpacing';
import { layoutDocument } from './index';
import type { FlowBlock, ParagraphBlock, ParagraphMeasure } from './types';

const block = (runs: ParagraphBlock['runs'], spacing?: { before?: number; after?: number }) =>
  ({ kind: 'paragraph', id: 'p', runs, attrs: { spacing } }) as unknown as ParagraphBlock;

describe('paragraph spacing', () => {
  test('a style-spaced empty paragraph keeps its spacing', () => {
    const spacer = block([], { before: 24, after: 36 });
    expect(getSpacingBefore(spacer)).toBe(24);
    expect(getSpacingAfter(spacer)).toBe(36);
  });

  test('an empty paragraph holding one empty run keeps it too', () => {
    const spacer = block([{ kind: 'text', text: '' }] as ParagraphBlock['runs'], {
      before: 12,
      after: 12,
    });
    expect(getSpacingBefore(spacer)).toBe(12);
    expect(getSpacingAfter(spacer)).toBe(12);
  });

  test('a paragraph with text is unaffected', () => {
    const para = block([{ kind: 'text', text: 'Cover' }] as ParagraphBlock['runs'], {
      before: 8,
      after: 10,
    });
    expect(getSpacingBefore(para)).toBe(8);
    expect(getSpacingAfter(para)).toBe(10);
  });

  test('absent spacing reads as zero', () => {
    expect(getSpacingBefore(block([]))).toBe(0);
    expect(getSpacingAfter(block([]))).toBe(0);
  });
});

// --- The rule Word does have -------------------------------------------------

const line = (height: number): ParagraphMeasure => ({
  kind: 'paragraph',
  lines: [
    {
      fromRun: 0,
      fromChar: 0,
      toRun: 0,
      toChar: 1,
      width: 100,
      ascent: 8,
      descent: 2,
      lineHeight: height,
    },
  ],
  totalHeight: height,
});

const para = (id: string, attrs: Record<string, unknown>): ParagraphBlock =>
  ({
    kind: 'paragraph',
    id,
    pmStart: 0,
    pmEnd: 1,
    runs: [{ kind: 'text', text: 'x' }],
    attrs,
  }) as unknown as ParagraphBlock;

const LINE_HEIGHT = 20;

/** Vertical gap between the two paragraphs, in px. One line each, so any
 *  excess over the line height is spacing. */
function gapBetween(blocks: ParagraphBlock[]): number {
  const layout = layoutDocument(
    blocks as FlowBlock[],
    blocks.map(() => line(LINE_HEIGHT)),
    { pageSize: { w: 600, h: 4000 }, margins: { top: 0, right: 0, bottom: 0, left: 0 } }
  );
  const tops = layout.pages[0].fragments.filter((f) => f.kind === 'paragraph').map((f) => f.y);
  return tops[1] - tops[0];
}

describe('contextual spacing suppression (17.3.1.9)', () => {
  test('suppresses the gap between same-style paragraphs that opt in', () => {
    const attrs = { styleId: 'ListParagraph', contextualSpacing: true, spacing: { after: 200 } };
    expect(gapBetween([para('a', attrs), para('b', { ...attrs })])).toBe(LINE_HEIGHT);
  });

  test('leaves the gap when the paragraphs do not share a style', () => {
    expect(
      gapBetween([
        para('a', { styleId: 'ListParagraph', contextualSpacing: true, spacing: { after: 200 } }),
        para('b', { styleId: 'Normal', contextualSpacing: true, spacing: { after: 200 } }),
      ])
    ).toBeGreaterThan(LINE_HEIGHT);
  });

  test('leaves the gap when neither paragraph opts in', () => {
    const attrs = { styleId: 'Normal', spacing: { after: 200 } };
    expect(gapBetween([para('a', attrs), para('b', { ...attrs })])).toBeGreaterThan(LINE_HEIGHT);
  });

  test('an empty spacer paragraph still pushes the next block down', () => {
    // The bug the removed heuristic caused: a template pushes its cover block
    // down the page with empty paragraphs carrying only style-inherited
    // spacing, and collapsing that stacked them at line height.
    const spacer = para('spacer', { styleId: 'Normal', spacing: { after: 400 } });
    spacer.runs = [];
    expect(gapBetween([spacer, para('cover', { styleId: 'Title' })])).toBeGreaterThan(LINE_HEIGHT);
  });
});

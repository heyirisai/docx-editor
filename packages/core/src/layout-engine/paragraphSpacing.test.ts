/**
 * Empty paragraphs keep their spacing.
 *
 * Templates push a cover block down the page with empty spacer paragraphs that
 * carry only style-inherited spacing. Collapsing that (Word does suppress it in
 * some cases, but the paginator already collapses against the previous block's
 * `after`) stacked the spacers at line height and left the cover's anchored art
 * overlapping the text below.
 */
import { describe, expect, test } from 'bun:test';
import { getSpacingAfter, getSpacingBefore } from './paragraphSpacing';
import type { ParagraphBlock } from './types';

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

/**
 * Regression test for #381 — empty paragraphs flagged with
 * `suppressEmptyParagraphHeight` measure as zero-height anchors.
 *
 * Used by the HF measurement pipeline to handle the canonical OOXML
 * "trailing empty paragraph after a table" pattern: the paragraph mark
 * exists in the document model but Word renders it as a zero-height
 * anchor, not a full line-height of phantom space.
 */

import { describe, test, expect } from 'bun:test';
import { measureParagraph } from '../measureParagraph';
import type { ParagraphBlock } from '../../../layout-engine/types';

function emptyPara(attrs: ParagraphBlock['attrs'] = {}): ParagraphBlock {
  return {
    kind: 'paragraph',
    id: 'p',
    runs: [],
    attrs,
  };
}

describe('measureParagraph — suppressEmptyParagraphHeight (#381)', () => {
  test('empty paragraph without flag uses default empty-line height', () => {
    const measure = measureParagraph(emptyPara(), 600);
    expect(measure.totalHeight).toBeGreaterThan(0);
    expect(measure.lines.length).toBe(1);
    expect(measure.lines[0].lineHeight).toBeGreaterThan(0);
  });

  test('empty paragraph with flag measures as zero height', () => {
    const measure = measureParagraph(emptyPara({ suppressEmptyParagraphHeight: true }), 600);
    expect(measure.totalHeight).toBe(0);
    expect(measure.lines.length).toBe(1);
    expect(measure.lines[0].lineHeight).toBe(0);
    expect(measure.lines[0].ascent).toBe(0);
    expect(measure.lines[0].descent).toBe(0);
  });

  // Non-empty paragraphs ignore the flag by construction — the
  // suppress branch is gated on `runs.length === 0`. Not worth a
  // separate test (would need a Canvas-backed environment for real
  // text measurement).
});

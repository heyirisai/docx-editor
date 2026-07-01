/**
 * Regression: footer floating images that resolve fully off the page must not
 * inflate the HF visual bounds.
 *
 * Some templates accumulate stale anchored drawings in a footer — e.g. a
 * decorative "wave" image anchored `relativeFrom="paragraph"` with a large
 * (~10in) vertical offset. Word positions it ~10in below the footer paragraph,
 * which lands it past the page bottom, so Word clips it: the footer looks flat.
 *
 * The painter (`renderHeaderFooterContent`) culls such off-page floats; the
 * measurement (`calculateHeaderFooterVisualBounds`) must agree, or `visualBottom`
 * grows to the invisible wave and shoves the in-flow footer text onto the next
 * page (the reported bug).
 */

import { describe, test, expect } from 'bun:test';
import { calculateHeaderFooterVisualBounds } from '../headerFooterLayout';
import type { FlowBlock, ImageRun, Measure, ParagraphBlock } from '../../layout-engine/types';

// US-Letter at 96dpi.
const metrics = {
  section: 'footer' as const,
  pageSize: { w: 816, h: 1056 },
  margins: { top: 96, right: 96, bottom: 96, left: 96, header: 48, footer: 48 },
};

// 10.27in vertical offset in EMUs (914400 EMU/in) → ~986px, well past the page
// bottom once added to the footer flow origin.
const OFF_PAGE_OFFSET_EMU = Math.round(10.27 * 914400);

function waveRun(posOffset: number): ImageRun {
  return {
    kind: 'image',
    src: 'wave.jpeg',
    width: 832, // 8.67in
    height: 58, // 0.6in
    position: { vertical: { relativeTo: 'paragraph', posOffset } },
  };
}

function footerParagraph(runs: ImageRun[]): { block: ParagraphBlock; measure: Measure } {
  return {
    block: {
      kind: 'paragraph',
      id: 'p',
      runs: [{ kind: 'text', text: '© 2026 HeyIris' }, ...runs],
    },
    measure: { kind: 'paragraph', lines: [], totalHeight: 20 },
  };
}

function bounds(items: Array<{ block: FlowBlock; measure: Measure }>) {
  const blocks = items.map((i) => i.block);
  const measures = items.map((i) => i.measure);
  const total = measures.reduce((s, m) => s + (m.kind === 'paragraph' ? m.totalHeight : 0), 0);
  return calculateHeaderFooterVisualBounds(blocks, measures, total, metrics);
}

describe('calculateHeaderFooterVisualBounds — off-page footer floats', () => {
  test('an off-page wave image does not extend visualBottom', () => {
    // Eight stacked off-page waves (the real template) + the copyright text.
    const waves = Array.from({ length: 8 }, () => waveRun(OFF_PAGE_OFFSET_EMU));
    const { visualBottom } = bounds([footerParagraph(waves)]);
    // Bounds stay at the in-flow text extent (20), NOT ~986 + 58 = 1044.
    expect(visualBottom).toBe(20);
  });

  test('an on-page footer float still extends visualBottom', () => {
    // A small offset keeps the image within the page — it must still count so a
    // genuine decorative footer image is not clipped.
    const { visualBottom } = bounds([footerParagraph([waveRun(0)])]);
    // paragraphStartY (0) + offset 0 + height 58 = 58.
    expect(visualBottom).toBe(58);
  });
});

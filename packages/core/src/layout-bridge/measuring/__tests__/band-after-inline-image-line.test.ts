import { describe, test, expect } from 'bun:test';

// bun:test has no DOM; give measureTextWidth a deterministic canvas stub
// (8px per character) before the lazy getCanvasContext() first runs.
if (typeof document === 'undefined') {
  (globalThis as Record<string, unknown>).document = {
    createElement: () => ({
      getContext: () => ({
        font: '',
        measureText: (text: string) => ({ width: text.length * 8 }),
      }),
    }),
  };
}

import { measureParagraph } from '../measureParagraph';
import type { FloatingImageZone } from '../floatingZones';

/**
 * Band clearance for a line must use where that line really starts. An inline
 * image grows its line well past the text height; the next line starts below
 * the grown line, so a topAndBottom band just under the image must push it.
 * Pre-fix the running Y advanced by the text-only height, so the next line was
 * tested near the paragraph top, missed the band and overlapped the picture.
 */
describe('topAndBottom band after a line grown by an inline image', () => {
  const BAND: FloatingImageZone = {
    leftMargin: 0,
    rightMargin: 0,
    topY: 120,
    bottomY: 200,
    fullWidthBlock: true,
  };

  const measure = (floatingZones?: FloatingImageZone[]) =>
    measureParagraph(
      {
        kind: 'paragraph',
        id: 'p1',
        pmStart: 0,
        pmEnd: 20,
        runs: [
          { kind: 'image', width: 590, height: 100 },
          { kind: 'text', text: 'After the picture', fontSize: 12, fontFamily: 'Arial' },
        ],
        attrs: {},
      } as never,
      600,
      { floatingZones }
    );

  test('the text line below the image is pushed below the band', () => {
    // Fixture check: without the band, line 2 spans the band's top edge.
    const free = measure();
    expect(free.lines).toHaveLength(2);
    const [imgLine, textLine] = free.lines;
    expect(imgLine!.lineHeight).toBeLessThan(BAND.topY);
    expect(imgLine!.lineHeight + textLine!.lineHeight).toBeGreaterThan(BAND.topY);

    const banded = measure([BAND]);
    const [line1, line2] = banded.lines;
    expect(line1!.floatSkipBefore ?? 0).toBe(0);
    const line2Top = line1!.lineHeight + (line2!.floatSkipBefore ?? 0);
    expect(line2Top).toBeCloseTo(BAND.bottomY, 3);
  });
});

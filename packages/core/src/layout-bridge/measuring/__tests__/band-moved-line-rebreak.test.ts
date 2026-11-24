import { describe, expect, test } from 'bun:test';

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
 * A large-font line that dips into a topAndBottom band is moved below the
 * band. When a side-wrapped float sits beside that destination, the line must
 * be broken against the narrower width there — not keep the text it was
 * broken with at full width above the band, which would run into the float.
 */
const WIDTH = 400;
const SIDE_MARGIN = 200;
const TEXT = 'one two three four five six seven eight nine ten eleven twelve thirteen fourteen';

const paragraph = {
  kind: 'paragraph',
  id: 'p1',
  pmStart: 0,
  pmEnd: 200,
  runs: [
    {
      kind: 'text',
      text: TEXT,
      fontSize: 20,
      fontFamily: 'Arial',
    },
  ],
  attrs: {},
} as never;

describe('line moved below a topAndBottom band re-breaks against a side float', () => {
  const lineHeight = measureParagraph(paragraph, WIDTH).lines[0]!.lineHeight;
  // Line 2 starts at `lineHeight`. The band begins just past the 11pt
  // placement estimate (~14.7px) for that line, so only its real height
  // reaches the band and the line is moved after it was broken.
  const bandTop = lineHeight + 15;
  const bandBottom = bandTop + 100;
  const zones: FloatingImageZone[] = [
    { leftMargin: 0, rightMargin: 0, topY: bandTop, bottomY: bandBottom, fullWidthBlock: true },
    { leftMargin: SIDE_MARGIN, rightMargin: 0, topY: bandBottom, bottomY: bandBottom + 400 },
  ];
  const result = measureParagraph(paragraph, WIDTH, { floatingZones: zones });

  test('the fixture moves line 2 below the band', () => {
    expect(lineHeight * 2).toBeGreaterThan(bandTop);
    const line2 = result.lines[1]!;
    expect(line2.floatSkipBefore ?? 0).toBeGreaterThan(0);
    expect(lineHeight + (line2.floatSkipBefore ?? 0)).toBeGreaterThanOrEqual(bandBottom - 0.5);
    expect(line2.leftOffset).toBe(SIDE_MARGIN);
  });

  test('every line fits the width available where it is placed', () => {
    for (const line of result.lines) {
      const available = WIDTH - (line.leftOffset ?? 0) - (line.rightOffset ?? 0);
      expect(line.width).toBeLessThanOrEqual(available + 0.5);
    }
  });

  test('line 1 is unchanged and all text is still laid out', () => {
    expect(result.lines[0]!.floatSkipBefore ?? 0).toBe(0);
    expect(result.lines[0]!.leftOffset ?? 0).toBe(0);
    const last = result.lines[result.lines.length - 1]!;
    expect(last.toChar).toBe(TEXT.length);
    for (let i = 1; i < result.lines.length; i++) {
      expect(result.lines[i]!.fromChar).toBe(result.lines[i - 1]!.toChar);
    }
  });
});

import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import type { TextRun } from '../../layout-engine/types';
import { renderTextRun } from '../renderParagraph/runs';
import { resetCanvasContext } from '../../layout-bridge/measuring/measureContainer';

/**
 * OOXML w:w (horizontal scale) must widen the run's LAYOUT advance, not just
 * its paint. `transform: scaleX()` alone leaves the box at the unscaled width,
 * so a 110% word's glyphs overhang the following run — a single-space run
 * after it disappears ("Business" + " " + "Context" painted as
 * "BusinessContext").
 *
 * Recipe: box = unscaled width W/s, scaled by s (visual/client rect = W), plus
 * margin-right W - W/s so the next run starts at W (the measured advance).
 */

// happy-dom has no 2D canvas: stub one at 0.5em per glyph.
let restoreGetContext: (() => void) | undefined;
beforeAll(() => {
  GlobalRegistrator.register();
  const proto = HTMLCanvasElement.prototype as unknown as { getContext: unknown };
  const original = proto.getContext;
  proto.getContext = function () {
    return {
      font: '',
      measureText(text: string) {
        const m = /([\d.]+)px/.exec(this.font as string);
        const px = m ? parseFloat(m[1]!) : 16;
        return { width: text.length * px * 0.5 };
      },
    };
  };
  restoreGetContext = () => {
    proto.getContext = original;
  };
  resetCanvasContext();
});
afterAll(() => {
  restoreGetContext?.();
  resetCanvasContext();
  GlobalRegistrator.unregister();
});

const CHAR_PX = ((11 * 96) / 72) * 0.5;

function run(overrides: Partial<TextRun>): TextRun {
  return { kind: 'text', text: 'Business', fontSize: 11, fontFamily: 'Calibri', ...overrides };
}

describe('w:w horizontal scale painter recipe', () => {
  /** Advance the run occupies in the line, and its painted (client-rect) width. */
  const advance = (el: HTMLElement) =>
    parseFloat(el.style.width) + parseFloat(el.style.marginRight);
  const painted = (el: HTMLElement, scale: number) => parseFloat(el.style.width) * scale;

  test('a scaled run advances by its scaled width and paints exactly that wide', () => {
    const el = renderTextRun(run({ horizontalScale: 110 }), document);
    expect(el.style.transform).toBe('scaleX(1.1)');
    expect(el.style.transformOrigin).toBe('left center');
    expect(el.style.display).toBe('inline-block');
    const scaled = 8 * CHAR_PX * 1.1;
    expect(advance(el)).toBeCloseTo(scaled, 3);
    expect(painted(el, 1.1)).toBeCloseTo(scaled, 3);
    // The text node stays the span's direct child (click/caret mapping).
    expect(el.firstChild?.nodeType).toBe(3);
  });

  test('a condensed (w:w < 100) run pulls the next run in', () => {
    const el = renderTextRun(run({ horizontalScale: 80 }), document);
    expect(advance(el)).toBeCloseTo(8 * CHAR_PX * 0.8, 3);
    expect(parseFloat(el.style.marginRight)).toBeLessThan(0);
  });

  test('a scaled single-space run keeps a positive advance, tracking unscaled', () => {
    const letterSpacing = -10 / 15; // w:spacing w:val="-10"
    const el = renderTextRun(run({ text: ' ', horizontalScale: 110, letterSpacing }), document);
    expect(advance(el)).toBeCloseTo(CHAR_PX * 1.1 + letterSpacing, 3);
    expect(advance(el)).toBeGreaterThan(0);
    // scaleX would scale CSS letter-spacing too; Word adds w:spacing unscaled.
    expect(parseFloat(el.style.letterSpacing)).toBeCloseTo(letterSpacing / 1.1, 4);
  });

  test('an unscaled run keeps its natural inline box', () => {
    const el = renderTextRun(run({ letterSpacing: 0.4 }), document);
    expect(el.style.transform).toBe('');
    expect(el.style.width).toBe('');
    expect(el.style.marginRight).toBe('');
    expect(el.style.display).toBe('');
    expect(el.style.letterSpacing).toBe('0.4px');
  });
});

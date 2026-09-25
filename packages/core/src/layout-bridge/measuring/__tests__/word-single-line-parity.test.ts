/**
 * Word single-line-spacing parity (lineRule="auto", line=240).
 *
 * Word's "single" line pitch is GDI's tmHeight + tmExternalLeading for the
 * run font. For Arial and Times New Roman the hhea lineGap survives as
 * external leading, so Word lays both out at 1.1499 × font size — the
 * well-known 11pt Arial → 12.65pt and 12pt TNR → 13.8pt values. The
 * resolver used to drop the gap (1.1172 / 1.1074), which under-measured every
 * Arial / TNR line by ~3% and paginated Arial documents a page too late (a
 * Google-Docs-authored questionnaire whose 51 empty Arial spacer lines pushed
 * "Table of Contents" onto page 1 instead of page 2 in the editor).
 *
 * These cases pin the three paragraph flavours from that document plus the
 * paragraph-mark rule for text-less lines. Expected values are analytic
 * (font size × ratio × line multiplier + before/after), in points.
 */

import { describe, test, expect } from 'bun:test';

// bun:test has no DOM; give measureTextWidth a deterministic canvas stub
// (0.5em per character) before the lazy getCanvasContext() first runs.
// getFontMetrics reads lineHeight from the OS/2 ratio table, not the canvas.
if (typeof document === 'undefined') {
  (globalThis as Record<string, unknown>).document = {
    createElement: () => ({
      getContext: () => ({
        font: '',
        measureText(text: string) {
          const m = /([\d.]+)px/.exec(this.font as string);
          const px = m ? parseFloat(m[1]) : 16;
          return { width: text.length * px * 0.5 };
        },
      }),
    }),
  };
}

import { measureParagraph } from '../measureParagraph';
import { resolveFontFamily } from '../../../utils/fontResolver';
import type { ParagraphBlock } from '../../../layout-engine/types';

const PT_TO_PX = 96 / 72;
const PX_TO_PT = 72 / 96;
/** Arial / Times New Roman: (usWinAscent + usWinDescent + hhea.lineGap) / 2048 = 2355/2048. */
const ARIAL_RATIO = 2355 / 2048;

function block(
  partial: Partial<ParagraphBlock> & { attrs?: ParagraphBlock['attrs'] }
): ParagraphBlock {
  return {
    kind: 'paragraph',
    id: 'p',
    pmStart: 0,
    pmEnd: 1,
    runs: [],
    ...partial,
  } as ParagraphBlock;
}

function heightPt(b: ParagraphBlock): number {
  const m = measureParagraph(b, 600);
  return m.totalHeight * PX_TO_PT;
}

describe('font resolver: Word single-line ratios include GDI external leading', () => {
  test('Arial and Times New Roman resolve to 1.1499 (2355/2048)', () => {
    expect(resolveFontFamily('Arial').singleLineRatio).toBeCloseTo(ARIAL_RATIO, 3);
    expect(resolveFontFamily('Times New Roman').singleLineRatio).toBeCloseTo(ARIAL_RATIO, 3);
  });

  test('fonts whose hhea box spans the usWin box keep the plain usWin sum', () => {
    // Calibri: usWin 1950+550 = 2500, hhea 1536/−512 gap 452 → leading 0.
    expect(resolveFontFamily('Calibri').singleLineRatio).toBeCloseTo(2500 / 2048, 3);
  });
});

describe('Word line heights for lineRule=auto paragraphs (Arial)', () => {
  test('empty 11pt Arial paragraph, line=240 → 12.65pt', () => {
    const h = heightPt(
      block({
        attrs: {
          defaultFontSize: 11,
          defaultFontFamily: 'Arial',
          spacing: { before: 0, after: 0, line: 1, lineUnit: 'multiplier', lineRule: 'auto' },
        },
      })
    );
    expect(h).toBeCloseTo(11 * ARIAL_RATIO, 2); // 12.649pt
  });

  test('empty 8.5pt Arial spacer line (docDefault sz=17 on the paragraph mark) → 9.77pt', () => {
    const h = heightPt(
      block({
        attrs: {
          defaultFontSize: 8.5,
          defaultFontFamily: 'Arial',
          spacing: { before: 0, after: 0, line: 1, lineUnit: 'multiplier', lineRule: 'auto' },
        },
      })
    );
    expect(h).toBeCloseTo(8.5 * ARIAL_RATIO, 2); // 9.774pt
  });

  test('7pt Arial text with line=190 auto → 7 × 1.1499 × 190/240 = 6.37pt', () => {
    const h = heightPt(
      block({
        runs: [{ kind: 'text', text: '40, rue de Sèvres', fontSize: 7, fontFamily: 'Arial' }],
        attrs: {
          defaultFontSize: 7,
          defaultFontFamily: 'Arial',
          spacing: {
            before: 0,
            after: 0,
            line: 190 / 240,
            lineUnit: 'multiplier',
            lineRule: 'auto',
          },
        },
      })
    );
    expect(h).toBeCloseTo(7 * ARIAL_RATIO * (190 / 240), 2); // 6.372pt
  });

  test('12pt Arial TOC entry with after=100 twips → 13.80 + 5 = 18.80pt', () => {
    const h = heightPt(
      block({
        runs: [
          { kind: 'text', text: 'Introduction', fontSize: 12, fontFamily: 'Arial', bold: true },
        ],
        attrs: {
          defaultFontSize: 12,
          defaultFontFamily: 'Arial',
          spacing: {
            before: 0,
            after: 100 / 15, // 100 twips in px
            line: 1,
            lineUnit: 'multiplier',
            lineRule: 'auto',
          },
        },
      })
    );
    expect(h).toBeCloseTo(12 * ARIAL_RATIO + 5, 2); // 18.799pt
  });

  test('18pt bold Arial title, line=240 → 20.70pt', () => {
    const h = heightPt(
      block({
        runs: [
          {
            kind: 'text',
            text: 'Full - Provider Security Requirements',
            fontSize: 18,
            fontFamily: 'Arial',
            bold: true,
          },
        ],
        attrs: {
          defaultFontSize: 18,
          defaultFontFamily: 'Arial',
          spacing: { before: 0, after: 0, line: 1, lineUnit: 'multiplier', lineRule: 'auto' },
        },
      })
    );
    expect(h).toBeCloseTo(18 * ARIAL_RATIO, 2); // 20.698pt
  });

  test('12pt Times New Roman line → 13.80pt', () => {
    const h = heightPt(
      block({
        runs: [{ kind: 'text', text: 'Hello', fontSize: 12, fontFamily: 'Times New Roman' }],
        attrs: { spacing: { line: 1, lineUnit: 'multiplier', lineRule: 'auto' } },
      })
    );
    expect(h).toBeCloseTo(12 * ARIAL_RATIO, 2);
  });
});

describe('text-less lines take the paragraph mark font (w:pPr/w:rPr)', () => {
  test('a header paragraph holding only an anchored drawing measures as its 12pt mark, not the 11pt default', () => {
    const b = block({
      runs: [
        {
          kind: 'image',
          src: 'logo.jpg',
          width: 151,
          height: 69,
          position: {
            horizontal: { relativeTo: 'page', posOffset: 360045 },
            vertical: { relativeTo: 'page', posOffset: 1033780 },
          },
          displayMode: 'float',
          wrapType: 'none',
        } as never,
      ],
      attrs: {
        defaultFontSize: 12,
        defaultFontFamily: 'Arial',
        spacing: { before: 0, after: 0, line: 1, lineUnit: 'multiplier', lineRule: 'auto' },
      },
    });
    const m = measureParagraph(b, 600);
    expect(m.lines).toHaveLength(1);
    expect(m.totalHeight).toBeCloseTo(12 * ARIAL_RATIO * PT_TO_PX, 1); // 18.40px
  });
});

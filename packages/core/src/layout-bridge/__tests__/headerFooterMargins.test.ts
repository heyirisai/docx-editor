/**
 * Unit tests for extendMarginsForHeaderFooter (issue #705).
 *
 * The body-margin push must be driven by the header/footer's IN-FLOW band
 * height (`flowHeight`), not its float-inclusive `visualBottom`/`height`. A
 * page/margin-anchored letterhead in a header has a huge `visualBottom` but a
 * tiny `flowHeight`; counting `visualBottom` drove the effective top margin
 * past the page and the paginator hard-threw, blanking the document.
 */

import { describe, test, expect } from 'bun:test';
import { extendMarginsForHeaderFooter } from '../headerFooterMargins';
import type { PageMargins } from '../../layout-engine/types';
import type { HeaderFooterContent } from '../../layout-painter/renderPage';

// A4 in px (16838 twips tall), with the margins from the #705 repro doc.
const PAGE = { w: 794, h: 1123 };
const MARGINS: PageMargins = {
  top: 187,
  right: 57,
  bottom: 113,
  left: 94,
  header: 47,
  footer: 47,
};

function hf(partial: Partial<HeaderFooterContent>): HeaderFooterContent {
  return {
    blocks: [],
    measures: [],
    height: 0,
    ...partial,
  };
}

describe('extendMarginsForHeaderFooter', () => {
  test('no header/footer content → margins returned unchanged', () => {
    const { margins, finalMargins } = extendMarginsForHeaderFooter({
      pageSize: PAGE,
      margins: MARGINS,
      finalMargins: MARGINS,
    });
    expect(margins).toBe(MARGINS);
    expect(finalMargins).toBe(MARGINS);
  });

  test('#705 — page-anchored letterhead (tiny flowHeight, huge visualBottom) does NOT push the body', () => {
    // A full-page letterhead anchored in the first-page header: visualBottom
    // ~1760px (paints down the page) but only ~30px of actual in-flow text.
    const letterheadHeader = hf({ flowHeight: 30, height: 1760, visualBottom: 1760 });
    const { margins } = extendMarginsForHeaderFooter({
      pageSize: PAGE,
      margins: MARGINS,
      finalMargins: MARGINS,
      headers: [letterheadHeader],
    });
    // flowHeight 30 < availableHeaderSpace (187-47=140) → no extension, no clamp.
    expect(margins.top).toBe(MARGINS.top);
    // Content area stays healthy (this is the case that used to throw).
    expect(PAGE.h - margins.top - margins.bottom).toBeGreaterThan(0);
  });

  test('genuinely tall in-flow header pushes the body top down (flowHeight drives it)', () => {
    const tallHeader = hf({ flowHeight: 300, height: 300 });
    const { margins } = extendMarginsForHeaderFooter({
      pageSize: PAGE,
      margins: MARGINS,
      finalMargins: MARGINS,
      headers: [tallHeader],
    });
    // top = max(187, headerDistance 47 + flowHeight 300) = 347.
    expect(margins.top).toBe(347);
  });

  test('falls back to `height` when `flowHeight` is undefined', () => {
    const legacyHeader = hf({ height: 300 });
    const { margins } = extendMarginsForHeaderFooter({
      pageSize: PAGE,
      margins: MARGINS,
      finalMargins: MARGINS,
      headers: [legacyHeader],
    });
    expect(margins.top).toBe(347);
  });

  test('footer band pushes the bottom margin up', () => {
    const tallFooter = hf({ flowHeight: 200 });
    const { margins } = extendMarginsForHeaderFooter({
      pageSize: PAGE,
      margins: MARGINS,
      finalMargins: MARGINS,
      footers: [tallFooter],
    });
    // §17.6.11: `w:footer` is the distance from the page bottom to the
    // footer's BOTTOM edge, so the band occupies `distance + flowHeight` and
    // the body clears 47 + 200 from the page bottom. Measured in Word at three
    // `w:footer` values — see `renderPage`'s `footerBandTop`.
    expect(margins.bottom).toBe(247);
  });

  test('the max band across header variants (default + first-page) wins', () => {
    const { margins } = extendMarginsForHeaderFooter({
      pageSize: PAGE,
      margins: MARGINS,
      finalMargins: MARGINS,
      headers: [hf({ flowHeight: 120 }), hf({ flowHeight: 320 })],
    });
    expect(margins.top).toBe(47 + 320);
  });

  test('clamp: an absurd in-flow header degrades to a thin content band instead of throwing', () => {
    const warnings: string[] = [];
    const { margins } = extendMarginsForHeaderFooter({
      pageSize: PAGE,
      margins: MARGINS,
      finalMargins: MARGINS,
      headers: [hf({ flowHeight: 5000 })],
      warn: (m) => warnings.push(m),
    });
    const content = PAGE.h - margins.top - margins.bottom;
    expect(content).toBeGreaterThan(0); // never <= 0 → paginator never throws
    expect(margins.top + margins.bottom).toBeLessThanOrEqual(PAGE.h);
    expect(warnings).toHaveLength(1);
    expect(warnings[0]).toContain('clamping margins');
  });

  test('per-sectionBreak margins are extended in place', () => {
    const bodyBlocks = [
      {
        kind: 'sectionBreak' as const,
        margins: { top: 187, right: 57, bottom: 113, left: 94, header: 47, footer: 47 },
      },
    ];
    extendMarginsForHeaderFooter({
      pageSize: PAGE,
      margins: MARGINS,
      finalMargins: MARGINS,
      bodyBlocks: bodyBlocks as never,
      headers: [hf({ flowHeight: 300 })],
    });
    expect(bodyBlocks[0].margins.top).toBe(347);
  });
});

describe('w:titlePg — per-page bands', () => {
  test("a cover's tall first-page header pushes ONLY its own page", () => {
    // A full-page cover picture wrapped `square` in the first-page header:
    // its own header text is displaced below the artwork, so the band is
    // taller than the sheet. The rest of the section keeps a short header.
    const { margins, firstPageMargins } = extendMarginsForHeaderFooter({
      pageSize: PAGE,
      margins: MARGINS,
      finalMargins: MARGINS,
      sections: [
        {
          header: hf({ flowHeight: 30 }),
          firstHeader: hf({ flowHeight: 1200 }),
          titlePg: true,
        },
      ],
    });
    // Pages 2+ keep the authored margin — `max(header, firstHeader)` used to
    // hand them the cover's band and collapse every page of the section.
    expect(margins.top).toBe(MARGINS.top);
    // Page 1 gets a top margin bounded by the sheet, so nothing fits on it:
    // Word starts the body overleaf.
    expect(firstPageMargins).toBeDefined();
    expect(firstPageMargins!.top).toBe(PAGE.h);
    expect(firstPageMargins!.bottom).toBe(0);
  });

  test('titlePg with no first-page header means NO header on page one', () => {
    const { margins, firstPageMargins } = extendMarginsForHeaderFooter({
      pageSize: PAGE,
      margins: MARGINS,
      finalMargins: MARGINS,
      sections: [{ header: hf({ flowHeight: 400 }), titlePg: true }],
    });
    expect(margins.top).toBe(47 + 400);
    // No band on page one, so it keeps the authored margin.
    expect(firstPageMargins).toEqual(MARGINS);
  });

  test('no titlePg → no per-page override at all', () => {
    const { firstPageMargins } = extendMarginsForHeaderFooter({
      pageSize: PAGE,
      margins: MARGINS,
      finalMargins: MARGINS,
      sections: [{ header: hf({ flowHeight: 400 }) }],
    });
    expect(firstPageMargins).toBeUndefined();
  });
});

/**
 * Header/footer floats that wrap text displace the body in Word. A float that
 * leaves no usable width beside it in the text column (a page-wide banner, or a
 * `topAndBottom` object) becomes a full-width exclusion band for the body on
 * every page that uses that header/footer. `wrapNone` floats (in front of /
 * behind text) displace nothing.
 *
 * Synthetic geometry: A4 at 96dpi, 1in top/bottom margins, 1.25in side margins,
 * header distance 0.5in — a page-wide picture anchored to the page horizontally
 * and to its (empty) header paragraph vertically, nudged up so it starts just
 * above the page edge and ends ~0.43in below the top margin.
 */

import { describe, expect, test } from 'bun:test';
import {
  createHeaderFooterBodyExclusions,
  headerFooterBodyExclusions,
} from '../headerFooterFloatExclusions';
import type { HeaderFooterContent } from '../../layout-painter/renderPage';
import type {
  ImageRun,
  ParagraphBlock,
  ParagraphMeasure,
  TextBoxBlock,
  TextBoxMeasure,
} from '../../layout-engine/types';

const EMU_PER_PX = 9525;
const PAGE = {
  size: { w: 794, h: 1123 },
  margins: { top: 96, right: 120, bottom: 96, left: 120, header: 48, footer: 48 },
};
const HEADER_DISTANCE = 48;
const PARAGRAPH_HEIGHT = 18;

function picture(overrides: Partial<ImageRun> = {}): ImageRun {
  return {
    kind: 'image',
    src: 'data:image/png;base64,iVBORw0KGgo=',
    width: 797,
    height: 144,
    wrapType: 'square',
    cssFloat: 'none',
    displayMode: 'block',
    distTop: 0,
    distBottom: 0,
    distLeft: 12,
    distRight: 12,
    position: {
      horizontal: { relativeTo: 'page', posOffset: 0 },
      vertical: { relativeTo: 'paragraph', posOffset: -58 * EMU_PER_PX },
    },
    ...overrides,
  };
}

function story(...runs: ImageRun[]): HeaderFooterContent {
  const block: ParagraphBlock = { kind: 'paragraph', id: 'hp', runs };
  const measure: ParagraphMeasure = { kind: 'paragraph', lines: [], totalHeight: PARAGRAPH_HEIGHT };
  return {
    blocks: [block],
    measures: [measure],
    height: PARAGRAPH_HEIGHT,
    flowHeight: PARAGRAPH_HEIGHT,
  };
}

describe('headerFooterBodyExclusions', () => {
  test('a page-wide square-wrapped header picture blocks the body down to its bottom', () => {
    const bands = headerFooterBodyExclusions(story(picture()), 'header', PAGE, HEADER_DISTANCE);
    expect(bands).toHaveLength(1);
    // Painted at headerDistance + paragraphY (0) + posOffset (-58px).
    expect(bands[0]!.top).toBeCloseTo(HEADER_DISTANCE - 58, 5);
    expect(bands[0]!.bottom).toBeCloseTo(HEADER_DISTANCE - 58 + 144, 5);
  });

  test('distT / distB widen the band', () => {
    const bands = headerFooterBodyExclusions(
      story(picture({ distTop: 4, distBottom: 10 })),
      'header',
      PAGE,
      HEADER_DISTANCE
    );
    expect(bands[0]!.top).toBeCloseTo(HEADER_DISTANCE - 58 - 4, 5);
    expect(bands[0]!.bottom).toBeCloseTo(HEADER_DISTANCE - 58 + 144 + 10, 5);
  });

  test('tight, through and topAndBottom wraps displace too', () => {
    for (const wrapType of ['tight', 'through', 'topAndBottom']) {
      const bands = headerFooterBodyExclusions(
        story(picture({ wrapType })),
        'header',
        PAGE,
        HEADER_DISTANCE
      );
      expect(bands).toHaveLength(1);
    }
  });

  test('wrapNone floats (in front of text / behind text) displace nothing', () => {
    for (const wrapType of ['inFront', 'behind']) {
      const bands = headerFooterBodyExclusions(
        story(picture({ wrapType, displayMode: 'float' })),
        'header',
        PAGE,
        HEADER_DISTANCE
      );
      expect(bands).toEqual([]);
    }
  });

  test('inline header pictures displace nothing', () => {
    const inline = picture({ wrapType: 'inline', position: undefined });
    expect(headerFooterBodyExclusions(story(inline), 'header', PAGE, HEADER_DISTANCE)).toEqual([]);
  });

  test('a narrow float with room beside it in the text column is not a full-width band', () => {
    // 150px logo at the left margin: the text column keeps ~400px to its right.
    const logo = picture({
      width: 150,
      height: 120,
      cssFloat: 'left',
      displayMode: 'float',
      position: {
        horizontal: { relativeTo: 'margin', posOffset: 0 },
        vertical: { relativeTo: 'paragraph', posOffset: 0 },
      },
    });
    expect(headerFooterBodyExclusions(story(logo), 'header', PAGE, HEADER_DISTANCE)).toEqual([]);
  });

  test('a one-sided wrap blocks the column when its only allowed side has no room', () => {
    // 300px picture flush with the right edge of the text column: ~240px of
    // room on its left, none on its right.
    const flushRight = (cssFloat: ImageRun['cssFloat']) =>
      picture({
        width: 300,
        cssFloat,
        displayMode: cssFloat === 'none' ? 'block' : 'float',
        position: {
          horizontal: { relativeTo: 'margin', posOffset: 254 * EMU_PER_PX },
          vertical: { relativeTo: 'paragraph', posOffset: 0 },
        },
      });
    // wrapText="right" (text on the right only — the float sits left): blocked.
    expect(
      headerFooterBodyExclusions(story(flushRight('left')), 'header', PAGE, HEADER_DISTANCE)
    ).toHaveLength(1);
    // wrapText="bothSides": text can use the left side, so no full-width band.
    expect(
      headerFooterBodyExclusions(story(flushRight('none')), 'header', PAGE, HEADER_DISTANCE)
    ).toEqual([]);
  });

  test('a float lying wholly off the page is culled like the painter does', () => {
    const offPage = picture({
      position: {
        horizontal: { relativeTo: 'page', posOffset: 0 },
        vertical: { relativeTo: 'paragraph', posOffset: 2000 * EMU_PER_PX },
      },
    });
    expect(headerFooterBodyExclusions(story(offPage), 'header', PAGE, HEADER_DISTANCE)).toEqual([]);
  });

  test('a footer float is placed from the footer band, which grows up from the footer distance', () => {
    const banner = picture({
      height: 140,
      position: {
        horizontal: { relativeTo: 'page', posOffset: 0 },
        vertical: { relativeTo: 'paragraph', posOffset: -100 * EMU_PER_PX },
      },
    });
    const bands = headerFooterBodyExclusions(story(banner), 'footer', PAGE, 48);
    const flowTop = PAGE.size.h - 48 - PARAGRAPH_HEIGHT;
    expect(bands).toHaveLength(1);
    expect(bands[0]!.top).toBeCloseTo(flowTop - 100, 5);
    expect(bands[0]!.bottom).toBeCloseTo(flowTop - 100 + 140, 5);
  });

  test('a page-wide wrapping text box in the header is a band as well', () => {
    const box: TextBoxBlock = {
      kind: 'textBox',
      id: 'tb',
      width: 760,
      height: 130,
      content: [],
      displayMode: 'float',
      wrapType: 'square',
      wrapText: 'bothSides',
      position: {
        horizontal: { relativeTo: 'page', posOffset: 10 * EMU_PER_PX },
        vertical: { relativeTo: 'page', posOffset: 0 },
      },
    };
    const measure = { kind: 'textBox', width: 760, height: 130 } as unknown as TextBoxMeasure;
    const content: HeaderFooterContent = {
      blocks: [box],
      measures: [measure],
      height: 130,
      flowHeight: 0,
    };
    const bands = headerFooterBodyExclusions(content, 'header', PAGE, HEADER_DISTANCE);
    expect(bands).toEqual([{ top: 0, bottom: 130 }]);
  });
});

describe('createHeaderFooterBodyExclusions', () => {
  test('returns undefined when no header/footer float displaces the body', () => {
    expect(
      createHeaderFooterBodyExclusions({
        header: story(picture({ wrapType: 'inFront', displayMode: 'float' })),
      })
    ).toBeUndefined();
    expect(createHeaderFooterBodyExclusions({})).toBeUndefined();
  });

  test('page 1 uses the first-page stories under titlePg, later pages the default ones', () => {
    const exclusions = createHeaderFooterBodyExclusions({
      header: story(picture()),
      firstPageHeader: story(),
      titlePg: true,
      headerDistance: HEADER_DISTANCE,
    });
    expect(exclusions).toBeDefined();
    expect(exclusions!({ number: 1, ...PAGE })).toEqual([]);
    expect(exclusions!({ number: 2, ...PAGE })).toHaveLength(1);
  });

  test('without titlePg every page uses the default header', () => {
    const exclusions = createHeaderFooterBodyExclusions({
      header: story(picture()),
      firstPageHeader: story(),
      headerDistance: HEADER_DISTANCE,
    });
    expect(exclusions!({ number: 1, ...PAGE })).toHaveLength(1);
    expect(exclusions!({ number: 7, ...PAGE })).toHaveLength(1);
  });
});

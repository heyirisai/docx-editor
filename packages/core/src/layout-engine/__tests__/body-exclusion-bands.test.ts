/**
 * `LayoutOptions.pageBodyExclusions`: full-width bands a page's body must stay
 * clear of (from header/footer floats that wrap text and leave no room beside
 * them). A band over the top edge of the body starts that page's body below
 * it, one over the bottom edge ends the body above it, and the reduced area
 * paginates — content that no longer fits flows to the next page.
 */

import { describe, expect, test } from 'bun:test';
import { layoutDocument } from '../index';
import { resolveBodyArea } from '../paginator';
import type { FlowBlock, Measure, ParagraphFragment } from '../types';
import {
  makeLayoutOptions,
  makeLine,
  makeParagraphBlock,
  makeParagraphMeasure,
} from '../integration/helpers';

// US Letter, 1in margins: body area [96, 960] (864px).
const LINE = 24;

function paragraphs(count: number, linesEach = 1): { blocks: FlowBlock[]; measures: Measure[] } {
  const blocks: FlowBlock[] = [];
  const measures: Measure[] = [];
  for (let i = 0; i < count; i++) {
    blocks.push(makeParagraphBlock(i, `P${i}`, i * 10));
    measures.push(
      makeParagraphMeasure(Array.from({ length: linesEach }, () => makeLine(0, 0, 0, 2, 50, LINE)))
    );
  }
  return { blocks, measures };
}

function firstY(fragments: { y: number }[]): number {
  return Math.min(...fragments.map((f) => f.y));
}

describe('resolveBodyArea', () => {
  test('no bands leaves the area alone', () => {
    expect(resolveBodyArea(96, 960, undefined)).toEqual({ top: 96, bottom: 960 });
    expect(resolveBodyArea(96, 960, [])).toEqual({ top: 96, bottom: 960 });
  });

  test('a band over the top edge moves the top to its bottom', () => {
    expect(resolveBodyArea(96, 960, [{ top: -9, bottom: 137 }])).toEqual({ top: 137, bottom: 960 });
  });

  test('overlapping bands at the top are cleared together', () => {
    const bands = [
      { top: 120, bottom: 180 },
      { top: 0, bottom: 130 },
    ];
    expect(resolveBodyArea(96, 960, bands).top).toBe(180);
  });

  test('a band over the bottom edge moves the bottom to its top', () => {
    expect(resolveBodyArea(96, 960, [{ top: 900, bottom: 1056 }])).toEqual({
      top: 96,
      bottom: 900,
    });
  });

  test('a band starting less than a line inside an edge still covers it', () => {
    // A page-wide wrapping picture 4px below the top margin: the 4px above it
    // cannot hold a line, so the body starts below the picture.
    expect(resolveBodyArea(96, 960, [{ top: 100, bottom: 244 }])).toEqual({
      top: 244,
      bottom: 960,
    });
    expect(resolveBodyArea(96, 960, [{ top: 800, bottom: 950 }])).toEqual({
      top: 96,
      bottom: 800,
    });
    // A gap of a line or more is left to the body.
    expect(resolveBodyArea(96, 960, [{ top: 120, bottom: 244 }])).toEqual({
      top: 96,
      bottom: 960,
    });
  });

  test('bands wholly above, below or inside the body are ignored', () => {
    const bands = [
      { top: 0, bottom: 60 },
      { top: 400, bottom: 500 },
      { top: 1000, bottom: 1056 },
    ];
    expect(resolveBodyArea(96, 960, bands)).toEqual({ top: 96, bottom: 960 });
  });

  test('bands that would leave (almost) no body are ignored for the page', () => {
    expect(resolveBodyArea(96, 960, [{ top: 0, bottom: 950 }])).toEqual({ top: 96, bottom: 960 });
  });

  test('degenerate bands are ignored', () => {
    const bands = [
      { top: 50, bottom: 50 },
      { top: Number.NaN, bottom: 200 },
      { top: 200, bottom: 100 },
    ];
    expect(resolveBodyArea(96, 960, bands)).toEqual({ top: 96, bottom: 960 });
  });
});

describe('layoutDocument with pageBodyExclusions', () => {
  test('every page starts its body below a band covering the top margin', () => {
    const { blocks, measures } = paragraphs(80);
    const layout = layoutDocument(
      blocks,
      measures,
      makeLayoutOptions({ pageBodyExclusions: () => [{ top: -9, bottom: 137 }] })
    );
    expect(layout.pages.length).toBeGreaterThan(1);
    for (const page of layout.pages) {
      expect(firstY(page.fragments)).toBe(137);
      // The authored margins are untouched — the header still anchors to them.
      expect(page.margins.top).toBe(96);
    }
  });

  test('the reduced body area paginates: content that no longer fits flows on', () => {
    // 36 lines fill exactly one 864px body; a 41px band pushes the last two off.
    const { blocks, measures } = paragraphs(36);
    const plain = layoutDocument(blocks, measures, makeLayoutOptions());
    expect(plain.pages).toHaveLength(1);

    const banded = layoutDocument(
      blocks,
      measures,
      makeLayoutOptions({ pageBodyExclusions: () => [{ top: 0, bottom: 137 }] })
    );
    expect(banded.pages).toHaveLength(2);
    const lastOnPage1 = Math.max(...banded.pages[0]!.fragments.map((f) => f.y + f.height));
    expect(lastOnPage1).toBeLessThanOrEqual(960);
    expect(banded.pages[1]!.fragments).toHaveLength(2);
    expect(firstY(banded.pages[1]!.fragments)).toBe(137);
  });

  test('a paragraph split by the band continues below it on the next page', () => {
    const { blocks, measures } = paragraphs(1, 40);
    const layout = layoutDocument(
      blocks,
      measures,
      makeLayoutOptions({ pageBodyExclusions: () => [{ top: 0, bottom: 137 }] })
    );
    expect(layout.pages).toHaveLength(2);
    const [first, second] = layout.pages.map((p) => p.fragments[0] as ParagraphFragment);
    expect(first!.y).toBe(137);
    expect(second!.y).toBe(137);
    // (960 - 137) / 24 = 34 whole lines fit under the band.
    expect(first!.toLine).toBe(34);
  });

  test('a band over the bottom edge ends the body above it', () => {
    const { blocks, measures } = paragraphs(36);
    const layout = layoutDocument(
      blocks,
      measures,
      makeLayoutOptions({ pageBodyExclusions: () => [{ top: 900, bottom: 1056 }] })
    );
    expect(layout.pages).toHaveLength(2);
    const bottom = Math.max(...layout.pages[0]!.fragments.map((f) => f.y + f.height));
    expect(bottom).toBeLessThanOrEqual(900);
  });

  test('a page records the raised body bottom so footnotes hang from it', () => {
    const { blocks, measures } = paragraphs(10);
    const layout = layoutDocument(
      blocks,
      measures,
      makeLayoutOptions({
        pageBodyExclusions: () => [{ top: 900, bottom: 1056 }],
        footnoteReservedHeights: new Map([[1, 60]]),
      })
    );
    expect(layout.pages[0]!.bodyBottom).toBe(900);
    // Without a band the page keeps the margin-derived bottom (no field).
    const plain = layoutDocument(blocks, measures, makeLayoutOptions());
    expect(plain.pages[0]!.bodyBottom).toBeUndefined();
  });

  test('bands apply only to the pages that report them', () => {
    const { blocks, measures } = paragraphs(80);
    const layout = layoutDocument(
      blocks,
      measures,
      makeLayoutOptions({
        pageBodyExclusions: (page) => (page.number === 1 ? [] : [{ top: 0, bottom: 137 }]),
      })
    );
    expect(firstY(layout.pages[0]!.fragments)).toBe(96);
    expect(firstY(layout.pages[1]!.fragments)).toBe(137);
  });

  test('the callback is asked once per page, with that page geometry', () => {
    const seen: number[] = [];
    const { blocks, measures } = paragraphs(80);
    layoutDocument(
      blocks,
      measures,
      makeLayoutOptions({
        pageBodyExclusions: (page) => {
          seen.push(page.number);
          expect(page.size).toEqual({ w: 816, h: 1056 });
          expect(page.margins.top).toBe(96);
          return undefined;
        },
      })
    );
    expect(seen).toEqual(seen.map((_, i) => i + 1));
    expect(seen.length).toBeGreaterThan(1);
  });

  test('an explicit page break after a full page does not leave a blank page', () => {
    const blocks: FlowBlock[] = [
      makeParagraphBlock(0, 'A', 0),
      { kind: 'pageBreak', id: 'pb' } as FlowBlock,
      { kind: 'pageBreak', id: 'pb2' } as FlowBlock,
      makeParagraphBlock(1, 'B', 10),
    ];
    const measures: Measure[] = [
      makeParagraphMeasure([makeLine(0, 0, 0, 1, 10, LINE)]),
      { kind: 'pageBreak' } as Measure,
      { kind: 'pageBreak' } as Measure,
      makeParagraphMeasure([makeLine(0, 0, 0, 1, 10, LINE)]),
    ];
    const layout = layoutDocument(
      blocks,
      measures,
      makeLayoutOptions({ pageBodyExclusions: () => [{ top: 0, bottom: 137 }] })
    );
    // The second break lands on a fresh (empty) page and collapses, exactly as
    // without bands: the "fresh page" check follows the moved body top.
    expect(layout.pages).toHaveLength(2);
  });
});

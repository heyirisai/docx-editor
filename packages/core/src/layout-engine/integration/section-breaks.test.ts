import { describe, test, expect } from 'bun:test';

import { layoutDocument } from '../index';
import type { FlowBlock, Measure } from '../types';

import { makeParagraphBlock, makeLine, makeParagraphMeasure, makeLayoutOptions } from './helpers';

describe('Section Breaks', () => {
  test('nextPage section break forces new page', () => {
    const blocks: FlowBlock[] = [
      makeParagraphBlock(0, 'Before section', 1),
      { kind: 'sectionBreak', id: 1, type: 'nextPage' },
      makeParagraphBlock(2, 'After section', 18),
    ];
    const measures: Measure[] = [
      makeParagraphMeasure([makeLine(0, 0, 0, 14, 120, 24)]),
      { kind: 'sectionBreak' },
      makeParagraphMeasure([makeLine(0, 0, 0, 13, 110, 24)]),
    ];

    const layout = layoutDocument(blocks, measures, makeLayoutOptions());

    expect(layout.pages.length).toBe(2);
    expect(layout.pages[0].fragments.some((f) => f.blockId === 0)).toBe(true);
    expect(layout.pages[1].fragments.some((f) => f.blockId === 2)).toBe(true);
  });

  test('continuous section break does not force new page', () => {
    // §17.6.22: the type that governs a break belongs to the section being
    // ENTERED, which for the last break is the body `sectPr` — hence
    // `bodyBreakType`, not the break block's own `type`.
    const blocks: FlowBlock[] = [
      makeParagraphBlock(0, 'Before section', 1),
      { kind: 'sectionBreak', id: 1, type: 'nextPage' },
      makeParagraphBlock(2, 'After section', 18),
    ];
    const measures: Measure[] = [
      makeParagraphMeasure([makeLine(0, 0, 0, 14, 120, 24)]),
      { kind: 'sectionBreak' },
      makeParagraphMeasure([makeLine(0, 0, 0, 13, 110, 24)]),
    ];

    const layout = layoutDocument(blocks, measures, {
      ...makeLayoutOptions(),
      bodyBreakType: 'continuous',
    });

    expect(layout.pages.length).toBe(1);
    expect(layout.pages[0].fragments.length).toBe(2);
  });

  /**
   * §17.6.22: a `sectPr`'s `w:type` describes how ITS OWN section starts, so
   * the break between section N and N+1 is governed by section N+1's type,
   * and an absent type is `nextPage` — never "carry the previous break's
   * type forward".
   *
   * Measured against Word 16 with a three-section probe:
   *   `sectPr0=nextPage, sectPr1=continuous, body=absent` -> AAA BBB | CCC
   *   `sectPr0=continuous, sectPr1=nextPage, body=absent` -> AAA | BBB | CCC
   */
  describe('the entered section owns the break type (§17.6.22)', () => {
    const threeSections = (first: 'nextPage' | 'continuous', second: 'nextPage' | 'continuous') => {
      const blocks: FlowBlock[] = [
        makeParagraphBlock(0, 'AAA', 1),
        { kind: 'sectionBreak', id: 1, type: first },
        makeParagraphBlock(2, 'BBB', 18),
        { kind: 'sectionBreak', id: 3, type: second },
        makeParagraphBlock(4, 'CCC', 35),
      ];
      const measures: Measure[] = [
        makeParagraphMeasure([makeLine(0, 0, 0, 3, 40, 24)]),
        { kind: 'sectionBreak' },
        makeParagraphMeasure([makeLine(0, 0, 0, 3, 40, 24)]),
        { kind: 'sectionBreak' },
        makeParagraphMeasure([makeLine(0, 0, 0, 3, 40, 24)]),
      ];
      return layoutDocument(blocks, measures, makeLayoutOptions());
    };

    test('a continuous SECOND section keeps its content on the first page', () => {
      const layout = threeSections('nextPage', 'continuous');
      expect(layout.pages.length).toBe(2);
      expect(layout.pages[0].fragments.map((f) => f.blockId)).toEqual([0, 2]);
      expect(layout.pages[1].fragments.map((f) => f.blockId)).toEqual([4]);
    });

    test('a nextPage SECOND section breaks even after a continuous one', () => {
      const layout = threeSections('continuous', 'nextPage');
      expect(layout.pages.length).toBe(3);
      expect(layout.pages.map((pg) => pg.fragments.map((f) => f.blockId))).toEqual([[0], [2], [4]]);
    });

    test("an absent type is nextPage, not the previous break's type", () => {
      const blocks: FlowBlock[] = [
        makeParagraphBlock(0, 'AAA', 1),
        { kind: 'sectionBreak', id: 1, type: 'continuous' },
        makeParagraphBlock(2, 'BBB', 18),
        // No `type` — §17.6.22 default. Pre-fix this inherited the
        // `continuous` above and COMET lost a whole divider sheet.
        { kind: 'sectionBreak', id: 3 },
        makeParagraphBlock(4, 'CCC', 35),
      ];
      const measures: Measure[] = [
        makeParagraphMeasure([makeLine(0, 0, 0, 3, 40, 24)]),
        { kind: 'sectionBreak' },
        makeParagraphMeasure([makeLine(0, 0, 0, 3, 40, 24)]),
        { kind: 'sectionBreak' },
        makeParagraphMeasure([makeLine(0, 0, 0, 3, 40, 24)]),
      ];
      const layout = layoutDocument(blocks, measures, makeLayoutOptions());
      // The untyped break defaults to `nextPage`, so BBB gets its own sheet;
      // inheriting the `continuous` above would have kept AAA and BBB together.
      expect(layout.pages.map((pg) => pg.fragments.map((f) => f.blockId))).toEqual([[0], [2], [4]]);
    });
  });
});

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
    const blocks: FlowBlock[] = [
      makeParagraphBlock(0, 'Before section', 1),
      { kind: 'sectionBreak', id: 1, type: 'continuous' },
      makeParagraphBlock(2, 'After section', 18),
    ];
    const measures: Measure[] = [
      makeParagraphMeasure([makeLine(0, 0, 0, 14, 120, 24)]),
      { kind: 'sectionBreak' },
      makeParagraphMeasure([makeLine(0, 0, 0, 13, 110, 24)]),
    ];

    const layout = layoutDocument(blocks, measures, makeLayoutOptions());

    expect(layout.pages.length).toBe(1);
    expect(layout.pages[0].fragments.length).toBe(2);
  });
});

describe('Section breaks — page section tagging', () => {
  // A section break block carries the CURRENT section's properties; its
  // `type` says how that section starts. The last section's start type is
  // `bodyBreakType`.
  test('nextPage sections open their page; a continuous section takes over from the next page', () => {
    const tallLines = Array.from({ length: 60 }, () => makeLine(0, 0, 0, 4, 40, 24));
    const blocks: FlowBlock[] = [
      makeParagraphBlock(0, 'Cover', 1),
      { kind: 'sectionBreak', id: 1, type: 'nextPage' },
      makeParagraphBlock(2, 'Body one', 10),
      { kind: 'sectionBreak', id: 3, type: 'nextPage' },
      makeParagraphBlock(4, 'Body two runs over a page', 20),
    ];
    const measures: Measure[] = [
      makeParagraphMeasure([makeLine(0, 0, 0, 5, 50, 24)]),
      { kind: 'sectionBreak' },
      makeParagraphMeasure([makeLine(0, 0, 0, 8, 80, 24)]),
      { kind: 'sectionBreak' },
      makeParagraphMeasure(tallLines),
    ];

    const layout = layoutDocument(
      blocks,
      measures,
      makeLayoutOptions({ bodyBreakType: 'continuous' })
    );

    expect(layout.pages.length).toBe(3);
    // Page 2 still starts in section 1 (section 2 is continuous and begins
    // mid-page); the page section 2 spills onto is its, but not its first.
    expect(layout.pages.map((p) => p.sectionIndex)).toEqual([0, 1, 2]);
    expect(layout.pages.map((p) => p.firstPageOfSection === true)).toEqual([true, true, false]);
  });
});

/**
 * `computePerBlockPageGeometry` gives each block its own section's page, the
 * way `computePerBlockWidths` gives it its own section's width: a section
 * break carries the properties of the section it ends.
 */

import { describe, expect, test } from 'bun:test';
import { computePerBlockPageGeometry } from '../sectionGeometry';
import type { FlowBlock, PageMargins } from '../../layout-engine/types';

const margins = (top: number): PageMargins => ({ top, right: 50, bottom: top, left: 50 });

describe('computePerBlockPageGeometry', () => {
  test("each block gets its own section's page size and margins", () => {
    const blocks = [
      { kind: 'paragraph', id: 'a' },
      { kind: 'sectionBreak', id: 'sb', margins: margins(60), pageSize: { w: 700, h: 900 } },
      { kind: 'paragraph', id: 'b' },
    ] as unknown as FlowBlock[];
    const geometry = computePerBlockPageGeometry(
      blocks,
      { pageSize: { w: 800, h: 1000 }, margins: margins(100) },
      { pageSize: { w: 800, h: 1000 }, margins: margins(200) }
    );
    expect(geometry.map((g) => g.marginTop)).toEqual([60, 60, 200]);
    expect(geometry[0].pageHeight).toBe(900);
    expect(geometry[2].contentHeight).toBe(600);
  });
});

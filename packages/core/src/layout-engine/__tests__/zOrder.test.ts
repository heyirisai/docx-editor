/**
 * Page z-order bands — document content (z-index from OOXML
 * relativeHeight) must never tie with or eclipse a front page-border
 * overlay (`w:pgBorders w:zOrder="front"`).
 */
import { describe, expect, test } from 'bun:test';
import { PAGE_OVERLAY_Z, contentZIndex } from '../zOrder';

describe('page z-order bands', () => {
  test('typical relativeHeight passes through unchanged (ordering preserved)', () => {
    expect(contentZIndex(251658240)).toBe(251658240);
    expect(contentZIndex(1)).toBe(1);
  });

  test('content never reaches the front-overlay band', () => {
    // relativeHeight is unsigned 32-bit — its maximum exceeds CSS int32.
    expect(contentZIndex(4294967295)).toBeLessThan(PAGE_OVERLAY_Z);
    expect(contentZIndex(PAGE_OVERLAY_Z)).toBe(PAGE_OVERLAY_Z - 1);
  });
});

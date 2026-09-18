/**
 * Page z-order bands. Two invariants, both by construction rather than by
 * convention: document content (z-index straight from OOXML `relativeHeight`)
 * must never tie with or eclipse a front page-border overlay
 * (`w:pgBorders w:zOrder="front"`), and must never reach the header/footer
 * band — a header float and a body float come from different OOXML stories, so
 * their `relativeHeight` values are not comparable.
 */
import { describe, expect, test } from 'bun:test';
import { HF_FRONT_Z_BASE, PAGE_OVERLAY_Z, contentZIndex, headerFooterFrontZIndex } from '../zOrder';

/** relativeHeight is unsigned 32-bit — its maximum exceeds CSS int32. */
const MAX_RELATIVE_HEIGHT = 4294967295;

describe('page z-order bands', () => {
  test('typical relativeHeight passes through unchanged (ordering preserved)', () => {
    expect(contentZIndex(251658240)).toBe(251658240);
    expect(contentZIndex(1)).toBe(1);
  });

  test('content never reaches the front-overlay band', () => {
    expect(contentZIndex(MAX_RELATIVE_HEIGHT)).toBeLessThan(PAGE_OVERLAY_Z);
    expect(contentZIndex(PAGE_OVERLAY_Z)).toBeLessThan(PAGE_OVERLAY_Z);
  });

  test('content never reaches the header/footer band', () => {
    // A body float authored above the band base used to clamp at
    // PAGE_OVERLAY_Z - 1 and paint over the header.
    expect(contentZIndex(HF_FRONT_Z_BASE)).toBeLessThan(HF_FRONT_Z_BASE);
    expect(contentZIndex(MAX_RELATIVE_HEIGHT)).toBeLessThan(headerFooterFrontZIndex(0));
    expect(contentZIndex(MAX_RELATIVE_HEIGHT)).toBeLessThan(
      headerFooterFrontZIndex(MAX_RELATIVE_HEIGHT)
    );
  });

  test('header/footer floats keep their order among themselves', () => {
    expect(headerFooterFrontZIndex(2)).toBeGreaterThan(headerFooterFrontZIndex(1));
    expect(headerFooterFrontZIndex(MAX_RELATIVE_HEIGHT)).toBeLessThan(PAGE_OVERLAY_Z);
  });
});

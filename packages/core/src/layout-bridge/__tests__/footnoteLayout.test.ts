import { describe, expect, test } from 'bun:test';
import { calculateFootnoteReservedHeights, FOOTNOTE_SEPARATOR_HEIGHT } from '../footnoteLayout';

describe('footnote layout reservation', () => {
  test('adds the shared separator height to each page reservation', () => {
    const reserved = calculateFootnoteReservedHeights(
      new Map([
        [1, [10, 11]],
        [3, [12]],
      ]),
      new Map([
        [10, { height: 14 }],
        [11, { height: 18 }],
        [12, { height: 9 }],
      ])
    );

    expect(reserved.get(1)).toBe(14 + 18 + FOOTNOTE_SEPARATOR_HEIGHT);
    expect(reserved.get(3)).toBe(9 + FOOTNOTE_SEPARATOR_HEIGHT);
  });
});

import { describe, expect, test } from 'bun:test';
import { clampFloatingWrapMargins } from '../measureParagraph';

describe('clampFloatingWrapMargins', () => {
  test('reports a full-width band when the float leaves no room beside it', () => {
    // Word pushes the line below such a float; reporting plain zero margins
    // dropped the exclusion entirely and text ran under the artwork.
    expect(clampFloatingWrapMargins(698, 0, 671)).toEqual({
      leftMargin: 0,
      rightMargin: 0,
      fullWidthBlock: true,
    });
  });

  test('preserves valid side margins', () => {
    expect(clampFloatingWrapMargins(200, 0, 671)).toEqual({ leftMargin: 200, rightMargin: 0 });
    expect(clampFloatingWrapMargins(0, 150, 671)).toEqual({ leftMargin: 0, rightMargin: 150 });
  });

  test('reports a full-width band when combined margins cover the line', () => {
    expect(clampFloatingWrapMargins(400, 300, 671)).toEqual({
      leftMargin: 0,
      rightMargin: 0,
      fullWidthBlock: true,
    });
  });
});

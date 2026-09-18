import { describe, expect, test } from 'bun:test';
import { resolveShadingBackgroundHex } from '../colorResolver';
import type { ShadingProperties, Theme } from '../../types/document';

/**
 * ECMA-376 §17.3.5 — `w:shd` is a two-colour pattern: `w:fill` is the
 * background, `w:color` the pattern foreground, `w:val` the coverage. The
 * render paths used to read `w:fill` alone, so a table header shaded
 * `w:val="solid" w:color="2E5090" w:fill="auto"` painted nothing — and the
 * white header text the author paired with it vanished into the page.
 */

const shd = (p: Partial<ShadingProperties>): ShadingProperties => p as ShadingProperties;

describe('resolveShadingBackgroundHex', () => {
  test('clear paints the fill', () => {
    expect(
      resolveShadingBackgroundHex(shd({ pattern: 'clear', fill: { rgb: 'D9E2F3' } }), null)
    ).toBe('D9E2F3');
  });

  test('solid paints the pattern colour, not the fill', () => {
    expect(
      resolveShadingBackgroundHex(
        shd({ pattern: 'solid', color: { rgb: '2E5090' }, fill: { auto: true } }),
        null
      )
    ).toBe('2E5090');
  });

  test('solid with no colour falls back to the fill', () => {
    expect(
      resolveShadingBackgroundHex(shd({ pattern: 'solid', fill: { rgb: 'ABCDEF' } }), null)
    ).toBe('ABCDEF');
  });

  test('nil paints nothing', () => {
    expect(
      resolveShadingBackgroundHex(shd({ pattern: 'nil', fill: { rgb: 'FF0000' } }), null)
    ).toBeUndefined();
  });

  test('a percentage pattern mixes foreground over background', () => {
    // 50% black over white.
    expect(
      resolveShadingBackgroundHex(
        shd({ pattern: 'pct50', color: { rgb: '000000' }, fill: { rgb: 'FFFFFF' } }),
        null
      )
    ).toBe('808080');
    // 25% -> a quarter of the way to black.
    expect(
      resolveShadingBackgroundHex(
        shd({ pattern: 'pct25', color: { rgb: '000000' }, fill: { rgb: 'FFFFFF' } }),
        null
      )
    ).toBe('BFBFBF');
  });

  test('a percentage pattern over an auto background still shows its colour', () => {
    expect(
      resolveShadingBackgroundHex(
        shd({ pattern: 'pct15', color: { rgb: '123456' }, fill: { auto: true } }),
        null
      )
    ).toBe('123456');
  });

  test('a named line pattern reads as its background fill', () => {
    expect(
      resolveShadingBackgroundHex(
        shd({ pattern: 'horzStripe', color: { rgb: '000000' }, fill: { rgb: 'EEEEEE' } }),
        null
      )
    ).toBe('EEEEEE');
  });

  test('theme fills resolve through the theme', () => {
    const theme = {
      colorScheme: { accent1: '4472C4' },
    } as unknown as Theme;
    expect(
      resolveShadingBackgroundHex(shd({ pattern: 'clear', fill: { themeColor: 'accent1' } }), theme)
    ).toBe('4472C4');
  });

  test('no shading, or an unset one, paints nothing', () => {
    expect(resolveShadingBackgroundHex(undefined, null)).toBeUndefined();
    expect(
      resolveShadingBackgroundHex(shd({ pattern: 'clear', fill: { auto: true } }), null)
    ).toBeUndefined();
  });
});

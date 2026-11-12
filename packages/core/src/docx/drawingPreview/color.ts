/**
 * DrawingML colour (`a:srgbClr`, `a:schemeClr`, … plus their modifiers) as a
 * concrete sRGB value with opacity, for painting a shape preview.
 *
 * The document model's `ColorValue` keeps only tint/shade; shapes lean on
 * `lumMod`/`lumOff` and `alpha` (a theme colour "lighter 40%" is
 * `lumMod 60000 + lumOff 40000`), so they are resolved here instead.
 */

import type { Theme, ThemeColorScheme } from '../../types/document';
import { getThemeColor } from '../themeParser';
import { getAttribute, getChildElements, getLocalName, type XmlElement } from '../xmlParser';

export interface ResolvedColor {
  /** `#rrggbb` — always six validated hex digits, safe to write into markup. */
  hex: string;
  /** 0..1 */
  alpha: number;
}

const COLOR_ELEMENTS = new Set(['srgbClr', 'schemeClr', 'sysClr', 'prstClr', 'hslClr', 'scrgbClr']);

/** `a:schemeClr@val` → theme slot, after the default `w:clrSchemeMapping`. */
const SCHEME_SLOT: Record<string, keyof ThemeColorScheme> = {
  dk1: 'dk1',
  lt1: 'lt1',
  dk2: 'dk2',
  lt2: 'lt2',
  tx1: 'dk1',
  bg1: 'lt1',
  tx2: 'dk2',
  bg2: 'lt2',
  accent1: 'accent1',
  accent2: 'accent2',
  accent3: 'accent3',
  accent4: 'accent4',
  accent5: 'accent5',
  accent6: 'accent6',
  hlink: 'hlink',
  folHlink: 'folHlink',
};

const PRESET: Record<string, string> = {
  black: '000000',
  white: 'FFFFFF',
  red: 'FF0000',
  green: '008000',
  lime: '00FF00',
  blue: '0000FF',
  yellow: 'FFFF00',
  cyan: '00FFFF',
  aqua: '00FFFF',
  magenta: 'FF00FF',
  fuchsia: 'FF00FF',
  gray: '808080',
  grey: '808080',
  dkGray: 'A9A9A9',
  ltGray: 'D3D3D3',
  silver: 'C0C0C0',
  orange: 'FFA500',
  purple: '800080',
  navy: '000080',
  maroon: '800000',
  olive: '808000',
  teal: '008080',
};

const SYSTEM: Record<string, string> = {
  windowText: '000000',
  window: 'FFFFFF',
  btnFace: 'F0F0F0',
  btnText: '000000',
  highlight: '0078D7',
  highlightText: 'FFFFFF',
  grayText: '6D6D6D',
};

const HEX6 = /^[0-9a-fA-F]{6}$/;

type Rgb = [number, number, number];

function hexToRgb(hex: string): Rgb {
  return [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255) as Rgb;
}

function rgbToHex([r, g, b]: Rgb): string {
  const h = (v: number) =>
    Math.round(Math.max(0, Math.min(1, v)) * 255)
      .toString(16)
      .padStart(2, '0');
  return `#${h(r)}${h(g)}${h(b)}`;
}

const toLinear = (c: number) => (c <= 0.04045 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4));
const toSrgb = (c: number) =>
  c <= 0.0031308 ? c * 12.92 : 1.055 * Math.pow(Math.max(0, c), 1 / 2.4) - 0.055;

function rgbToHsl([r, g, b]: Rgb): Rgb {
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h: number;
  if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
  else if (max === g) h = (b - r) / d + 2;
  else h = (r - g) / d + 4;
  return [h / 6, s, l];
}

function hslToRgb([h, s, l]: Rgb): Rgb {
  if (s === 0) return [l, l, l];
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const hue = (t: number) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  return [hue(h + 1 / 3), hue(h), hue(h - 1 / 3)];
}

const clamp01 = (v: number) => Math.max(0, Math.min(1, v));

/** Apply the modifier children of a colour element, in document order. */
function applyModifiers(base: Rgb, el: XmlElement): { rgb: Rgb; alpha: number } {
  let rgb = base;
  let alpha = 1;
  for (const mod of getChildElements(el)) {
    const raw = Number(getAttribute(mod, null, 'val'));
    if (!Number.isFinite(raw)) continue;
    const v = raw / 100000;
    switch (getLocalName(mod.name ?? '')) {
      case 'alpha':
        alpha = clamp01(v);
        break;
      case 'alphaMod':
        alpha = clamp01(alpha * v);
        break;
      case 'alphaOff':
        alpha = clamp01(alpha + v);
        break;
      case 'lumMod': {
        const [h, s, l] = rgbToHsl(rgb);
        rgb = hslToRgb([h, s, clamp01(l * v)]);
        break;
      }
      case 'lumOff': {
        const [h, s, l] = rgbToHsl(rgb);
        rgb = hslToRgb([h, s, clamp01(l + v)]);
        break;
      }
      case 'satMod': {
        const [h, s, l] = rgbToHsl(rgb);
        rgb = hslToRgb([h, clamp01(s * v), l]);
        break;
      }
      case 'tint':
        // Fraction of the colour kept; the rest moves toward white (linear RGB).
        rgb = rgb.map((c) => toSrgb(1 - (1 - toLinear(c)) * clamp01(v))) as Rgb;
        break;
      case 'shade':
        rgb = rgb.map((c) => toSrgb(toLinear(c) * clamp01(v))) as Rgb;
        break;
      default:
        break;
    }
  }
  return { rgb, alpha };
}

/**
 * The first colour element directly under `parent` (e.g. an `a:solidFill`,
 * an `a:gs`, a style `a:fillRef`). `phClr` — the placeholder colour theme
 * styles use — resolves to `placeholder` when given.
 */
export function resolveDrawingColor(
  parent: XmlElement | null | undefined,
  theme: Theme | null | undefined,
  placeholder?: ResolvedColor
): ResolvedColor | null {
  const el = getChildElements(parent).find((c) => COLOR_ELEMENTS.has(getLocalName(c.name ?? '')));
  if (!el) return null;
  const kind = getLocalName(el.name ?? '');
  const val = getAttribute(el, null, 'val') ?? '';
  let hex: string | null = null;
  let baseAlpha = 1;

  switch (kind) {
    case 'srgbClr':
      hex = HEX6.test(val) ? val : null;
      break;
    case 'schemeClr': {
      if (val === 'phClr') {
        if (!placeholder) return null;
        hex = placeholder.hex.slice(1);
        baseAlpha = placeholder.alpha;
        break;
      }
      const slot = SCHEME_SLOT[val];
      const themed = slot ? getThemeColor(theme ?? null, slot) : null;
      hex = themed && HEX6.test(themed) ? themed : null;
      break;
    }
    case 'sysClr': {
      const last = getAttribute(el, null, 'lastClr') ?? '';
      hex = HEX6.test(last) ? last : (SYSTEM[val] ?? null);
      break;
    }
    case 'prstClr':
      hex = Object.prototype.hasOwnProperty.call(PRESET, val) ? PRESET[val] : null;
      break;
    case 'scrgbClr': {
      const pct = (a: string) => clamp01(Number(getAttribute(el, null, a)) / 100000 || 0);
      hex = rgbToHex([toSrgb(pct('r')), toSrgb(pct('g')), toSrgb(pct('b'))]).slice(1);
      break;
    }
    case 'hslClr': {
      const hue = (Number(getAttribute(el, null, 'hue')) || 0) / 21600000;
      const sat = clamp01((Number(getAttribute(el, null, 'sat')) || 0) / 100000);
      const lum = clamp01((Number(getAttribute(el, null, 'lum')) || 0) / 100000);
      hex = rgbToHex(hslToRgb([hue, sat, lum])).slice(1);
      break;
    }
  }
  if (!hex) return null;

  const { rgb, alpha } = applyModifiers(hexToRgb(hex), el);
  return { hex: rgbToHex(rgb), alpha: clamp01(alpha * baseAlpha) };
}

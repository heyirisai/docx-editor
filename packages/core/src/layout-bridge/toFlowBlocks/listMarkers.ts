/**
 * List Marker Resolution
 *
 * Helpers for rendering OOXML list markers from the counter stack:
 *   - format numbers as decimal/roman/letter per ECMA-376 §17.9.16 numFmt
 *   - resolve lvlText templates ("%1.%2.") against the counter stack
 *   - drive the per-paragraph counter increment from each level's start
 *     value, including startOverride.
 */

import type { NumberFormat } from '../../types/document';
import type { ParagraphAttrs as PMParagraphAttrs } from '../../prosemirror/schema/nodes';
import { convertBulletToUnicode } from '../../docx/blockContentParser';
import { clampListStart, padDecimal } from '../../docx/numberingParser';

/**
 * Counter slot of a level nothing on the abstractNum has numbered yet. It takes
 * the level's abstract start (`w:start`), not a numId's startOverride.
 */
const NEVER = -2;

/**
 * Counter slot of a level a shallower level restarted. It takes the numId's
 * start for the level, startOverride included. A real value can be 0 (a level
 * whose start is 0), so both markers are negative.
 */
const UNSET = -1;

export function formatNumberedMarker(counters: number[], level: number): string {
  const parts: number[] = [];
  for (let i = 0; i <= level; i += 1) {
    const value = counters[i] ?? UNSET;
    if (value < 0) break;
    parts.push(value);
  }
  if (parts.length === 0) return '1.';
  return `${parts.join('.')}.`;
}

const ROMAN_PAIRS: Array<[number, string]> = [
  [1000, 'M'],
  [900, 'CM'],
  [500, 'D'],
  [400, 'CD'],
  [100, 'C'],
  [90, 'XC'],
  [50, 'L'],
  [40, 'XL'],
  [10, 'X'],
  [9, 'IX'],
  [5, 'V'],
  [4, 'IV'],
  [1, 'I'],
];

function toRoman(n: number, upper: boolean): string {
  if (n <= 0) return '';
  let value = n;
  let out = '';
  for (const [num, sym] of ROMAN_PAIRS) {
    while (value >= num) {
      out += sym;
      value -= num;
    }
  }
  return upper ? out : out.toLowerCase();
}

// Spreadsheet-style: 1→A, 26→Z, 27→AA, 28→AB, ...
function toLetter(n: number, upper: boolean): string {
  if (n <= 0) return '';
  let value = n;
  let out = '';
  while (value > 0) {
    const rem = (value - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    value = Math.floor((value - 1) / 26);
  }
  return upper ? out : out.toLowerCase();
}

/** Formats with no zero. Word prints a 0 in them as a space (see {@link formatCounter}). */
const NO_ZERO_FORMATS: ReadonlySet<NumberFormat> = new Set([
  'upperRoman',
  'lowerRoman',
  'upperLetter',
  'lowerLetter',
]);

function formatCounter(value: number, fmt: NumberFormat | undefined): string {
  // 0 prints as "0" in the decimal formats. Roman and letter have no zero, so
  // Word prints a space in its place and keeps the punctuation after it:
  // "%1." reads " ." for a lowerLetter level at 0.
  if (value < 0) return '';
  if (value === 0 && fmt && NO_ZERO_FORMATS.has(fmt)) return ' ';
  switch (fmt) {
    case 'upperRoman':
      return toRoman(value, true);
    case 'lowerRoman':
      return toRoman(value, false);
    case 'upperLetter':
      return toLetter(value, true);
    case 'lowerLetter':
      return toLetter(value, false);
    case 'decimalZero':
      return padDecimal(value, 2);
    case 'decimalZero3':
      return padDecimal(value, 3);
    case 'decimalZero4':
      return padDecimal(value, 4);
    case 'decimalZero5':
      return padDecimal(value, 5);
    case 'none':
      return '';
    default:
      // decimal and unsupported formats fall back to decimal
      return String(value);
  }
}

/**
 * Resolve an OOXML lvlText template like "%1.%2." against the counter stack
 * and per-level numFmt list (ECMA-376 §17.9.11).
 *
 * When a referenced counter has no value yet (e.g. "%2" referenced from a
 * level-0 paragraph), the placeholder AND the punctuation immediately
 * following it are dropped — matches Word's behavior so "%1.%2." renders
 * "1." rather than "1..".
 *
 * Exported for unit testing.
 */
export function resolveListTemplate(
  template: string,
  counters: number[],
  levelNumFmts: NumberFormat[] | undefined
): string {
  return resolveTemplate(template, counters, levelNumFmts, 1);
}

/**
 * {@link resolveListTemplate} with the smallest counter value that counts as
 * "has a number": 1 for the exported helper, 0 for the marker computation,
 * whose counter stack marks numberless levels with {@link UNSET}.
 */
function resolveTemplate(
  template: string,
  counters: number[],
  levelNumFmts: NumberFormat[] | undefined,
  minValue: number
): string {
  return template.replace(/%(\d)([.):\]])?/g, (_, digit, punct = '') => {
    const idx = parseInt(digit, 10) - 1;
    if (idx < 0) return '';
    const value = counters[idx] ?? UNSET;
    if (value < minValue) return '';
    const fmt = levelNumFmts?.[idx] ?? 'decimal';
    const formatted = formatCounter(value, fmt);
    return formatted ? formatted + punct : '';
  });
}

/**
 * Start value of `ilvl` for this paragraph's numId: the parsed per-level start
 * (lvlOverride startOverride, else `w:start`, else 0), falling back for
 * paragraphs without it (editor-created lists) to the paragraph's own
 * startOverride, then 1. Clamped to 0–32767 whatever the attrs carry, so the
 * counters (and the roman formatter's work) stay bounded by the cap plus the
 * list's length.
 */
function levelStart(pmAttrs: PMParagraphAttrs, ilvl: number, paraLevel: number): number {
  const start =
    pmAttrs.listLevelStarts?.[ilvl] ??
    (ilvl === paraLevel ? pmAttrs.listStartOverride : undefined) ??
    1;
  return clampListStart(start);
}

/**
 * Start value of `ilvl` without the numId's startOverride: the level's own
 * `w:start` (ListRendering.levelAbstractStarts), else 1 for paragraphs
 * without it (editor-created lists). Clamped like {@link levelStart}.
 */
function abstractStart(pmAttrs: PMParagraphAttrs, ilvl: number): number {
  return clampListStart(pmAttrs.listLevelAbstractStarts?.[ilvl] ?? 1);
}

/**
 * Advance the counter stack for a list paragraph and return the rendered
 * marker. Mutates `counters` in place. Returns null when no marker should
 * be drawn (numId is missing or 0 — "no numbering" per ECMA-376).
 */
export function computeListMarker(
  pmAttrs: PMParagraphAttrs,
  listCounters: Map<number, number[]>,
  seenNumIds: Set<string>
): string | null {
  const numPr = pmAttrs.numPr;
  if (!numPr) return null;
  const numId = numPr.numId;
  if (numId == null || numId === 0) return null;

  // Bullets don't consume a numbering slot — they share a numId with numbered
  // levels in some templates, and incrementing here would skip numbers.
  // Run the Symbol-font glyph mapper here too so bullets in table cells and
  // text boxes get the same Unicode conversion that body bullets get from
  // the parser-side resolveBulletMarker (idempotent for already-Unicode chars).
  if (pmAttrs.listIsBullet) {
    return convertBulletToUnicode(pmAttrs.listMarker || '');
  }

  const level = numPr.ilvl ?? 0;

  // numFmt="none" (ECMA-376 §17.18.59): the level shows no auto-number — its
  // `lvlText` is taken literally (empty `lvlText` ⇒ no marker at all). Word
  // renders such paragraphs as plain text, so don't fall through to the
  // generic decimal formatter below and fabricate a "1." marker. (#718)
  const levelFmt = pmAttrs.listLevelNumFmts?.[level] ?? pmAttrs.listNumFmt;
  if (levelFmt === 'none') {
    return pmAttrs.listMarker ? pmAttrs.listMarker : null;
  }

  const counterKey = pmAttrs.listAbstractNumId ?? numId;
  const counters = listCounters.get(counterKey) ?? new Array<number>(9).fill(NEVER);

  // A numId's startOverride for a level restarts that level the first time the
  // numId numbers a paragraph AT that level (ECMA-376 §17.9.27), even when
  // another numId on the same abstractNum has been counting it. Word's
  // "Restart at 1" is such an override on level 0. Shallower levels the
  // paragraph skips are left alone, as in Word.
  const seenKey = `${numId}:${level}`;
  let restartedHere = false;
  if (!seenNumIds.has(seenKey)) {
    seenNumIds.add(seenKey);
    restartedHere = pmAttrs.listStartOverride != null;
  }

  // A shallower level with no number yet takes a start value and counts as
  // used, as in Word: its abstract `w:start` if nothing has numbered it yet,
  // or the numId's start (startOverride included) if an even shallower level
  // restarted it. So a list that opens with sub-items numbers them "1.1.",
  // "1.2." and its first top-level item "2." ("5.1.", "5.2.", "6." when level
  // 0's w:start is 5), while a "Restart at 1" list that opens with sub-items
  // prints "1.1.", "1.2." and then "1." again. A skipped level prints "1.1.1."
  // (not "1.0.1."), and the next item at that level is "1.2.".
  for (let i = 0; i < level; i += 1) {
    const value = counters[i] ?? NEVER;
    if (value === NEVER) counters[i] = abstractStart(pmAttrs, i);
    else if (value === UNSET) counters[i] = levelStart(pmAttrs, i, level);
  }

  const current = counters[level] ?? NEVER;
  if (restartedHere || current === UNSET) counters[level] = levelStart(pmAttrs, level, level);
  else if (current === NEVER) counters[level] = abstractStart(pmAttrs, level);
  else counters[level] = current + 1;
  for (let i = level + 1; i < counters.length; i += 1) {
    counters[i] = UNSET;
  }
  listCounters.set(counterKey, counters);

  // Parsed lvlText template (e.g. "%1." or "%1.%2.") resolves against the
  // counter stack. Editor-created lists with no template fall back to the
  // generic decimal formatter.
  if (pmAttrs.listMarker && pmAttrs.listMarker.includes('%')) {
    return resolveTemplate(pmAttrs.listMarker, counters, pmAttrs.listLevelNumFmts ?? undefined, 0);
  }
  if (pmAttrs.listMarker) {
    return pmAttrs.listMarker;
  }
  return formatNumberedMarker(counters, level);
}

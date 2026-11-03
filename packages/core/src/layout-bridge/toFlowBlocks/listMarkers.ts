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
import { padDecimal } from '../../docx/numberingParser';

/**
 * Counter slot of a level with no number yet: unused so far, or restarted by a
 * shallower level. A real value can be 0 (a level whose start is 0).
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

function formatCounter(value: number, fmt: NumberFormat | undefined): string {
  // 0 prints as "0" in the decimal formats; roman and letter have no zero.
  if (value < 0) return '';
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
 * startOverride, then 1.
 */
function levelStart(pmAttrs: PMParagraphAttrs, ilvl: number, paraLevel: number): number {
  const start =
    pmAttrs.listLevelStarts?.[ilvl] ??
    (ilvl === paraLevel ? pmAttrs.listStartOverride : undefined) ??
    1;
  return Math.max(0, start);
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
  const counters = listCounters.get(counterKey) ?? new Array<number>(9).fill(UNSET);

  // The first paragraph of a numId restarts every level its lvlOverride gives a
  // startOverride (and the levels below it), even when another numId on the
  // same abstractNum has been counting them (ECMA-376 §17.9.27). Word's
  // "Restart at 1" is such an override on level 0.
  const seenKey = String(numId);
  if (!seenNumIds.has(seenKey)) {
    seenNumIds.add(seenKey);
    const restarted =
      pmAttrs.listStartOverrideLevels ?? (pmAttrs.listStartOverride != null ? [level] : []);
    for (const ilvl of restarted) {
      for (let i = Math.max(0, ilvl); i < counters.length; i += 1) counters[i] = UNSET;
    }
  }

  // A shallower level with no number yet (unused so far in this list, or
  // restarted by an even shallower level) takes its start value and counts as
  // used, as in Word. A list that opens with sub-items numbers them "1.1.",
  // "1.2." and its first top-level item "2." ("5.1.", "5.2.", "6." when level
  // 0 starts at 5). A skipped level prints "1.1.1." (not "1.1."), and the next
  // item at that level is "1.2.".
  for (let i = 0; i < level; i += 1) {
    if ((counters[i] ?? UNSET) === UNSET) counters[i] = levelStart(pmAttrs, i, level);
  }

  const current = counters[level] ?? UNSET;
  counters[level] = current === UNSET ? levelStart(pmAttrs, level, level) : current + 1;
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

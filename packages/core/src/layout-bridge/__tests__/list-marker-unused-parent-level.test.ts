/**
 * A list paragraph whose shallower levels have no number yet. This happens
 * when a list opens below level 0, or when a level is skipped. Word gives each
 * such level its start value and counts it as used, so the next paragraph at
 * that level continues from it instead of restarting. Before the fix we left
 * those levels at 0. That dropped "%1." from the label ("1." for "1.1.") and
 * restarted the parent at 1 where Word prints 2.
 */

import { describe, test, expect } from 'bun:test';
import { toProseDoc } from '../../prosemirror/conversion/toProseDoc';
import { toFlowBlocks } from '../toFlowBlocks';
import { computeListRendering, parseNumbering } from '../../docx/numberingParser';
import type { Document, Paragraph } from '../../types/document';

const TEMPLATES = ['%1.', '%1.%2.', '%1.%2.%3.'];

function item(
  numId: number,
  ilvl: number,
  opts: { abstractNumId?: number; startOverride?: number } = {}
): Paragraph {
  return {
    type: 'paragraph',
    formatting: { numPr: { numId, ilvl } },
    content: [{ type: 'run', content: [{ type: 'text', text: `item ${numId}/${ilvl}` }] }],
    listRendering: {
      marker: TEMPLATES[ilvl],
      level: ilvl,
      numId,
      isBullet: false,
      numFmt: 'decimal',
      levelNumFmts: Array(ilvl + 1).fill('decimal'),
      abstractNumId: opts.abstractNumId ?? numId,
      startOverride: opts.startOverride,
    },
  };
}

function markers(content: Paragraph[]): string[] {
  const doc: Document = { package: { document: { content } } };
  const out: string[] = [];
  for (const block of toFlowBlocks(toProseDoc(doc), {})) {
    if (block.kind === 'paragraph' && block.attrs?.listMarker) out.push(block.attrs.listMarker);
  }
  return out;
}

describe('toFlowBlocks — shallower list levels with no number yet', () => {
  test('a list that opens at level 1 numbers its first level-0 item 2', () => {
    expect(markers([item(1, 1), item(1, 1), item(1, 1), item(1, 0), item(1, 1)])).toEqual([
      '1.1.',
      '1.2.',
      '1.3.',
      '2.',
      '2.1.',
    ]);
  });

  test('a separate heading list before it does not change that', () => {
    // A one-item heading list on its own abstractNum, then a second list that
    // opens with sub-items. The second list keeps its own counters.
    expect(
      markers([
        item(1, 0, { abstractNumId: 10 }),
        item(2, 1, { abstractNumId: 20 }),
        item(2, 1, { abstractNumId: 20 }),
        item(2, 0, { abstractNumId: 20 }),
        item(2, 1, { abstractNumId: 20 }),
        item(2, 0, { abstractNumId: 20 }),
        item(2, 1, { abstractNumId: 20 }),
      ])
    ).toEqual(['1.', '1.1.', '1.2.', '2.', '2.1.', '3.', '3.1.']);
  });

  test('a skipped level prints its start value and the next item there continues', () => {
    expect(markers([item(1, 0), item(1, 2), item(1, 1), item(1, 2)])).toEqual([
      '1.',
      '1.1.1.',
      '1.2.',
      '1.2.1.',
    ]);
  });

  test('a level restarted by its parent is filled in again when skipped', () => {
    expect(markers([item(1, 0), item(1, 1), item(1, 0), item(1, 2)])).toEqual([
      '1.',
      '1.1.',
      '2.',
      '2.1.1.',
    ]);
  });

  test('a startOverride on the opening level still applies to that level', () => {
    expect(markers([item(1, 1, { startOverride: 7 }), item(1, 1), item(1, 0)])).toEqual([
      '1.7.',
      '1.8.',
      '2.',
    ]);
  });
});

/**
 * The same paragraphs resolved through a parsed numbering part, so level start
 * values (`w:start`, `w:lvlOverride`/`w:startOverride`) reach the markers the
 * way they do for a loaded document. One abstractNum with three levels,
 * decimal unless `level0Fmt` says otherwise for level 0; `level0Start` is
 * level 0's `w:start` (omitted when null), levels 1 and 2 start at 1.
 * `overrides` maps a numId to its per-level startOverrides; numId 1 is a plain
 * instance unless `overrides` gives it some.
 */
function numberingXml(
  level0Start: number | null,
  overrides: Record<number, Record<number, number>> = {},
  level0Fmt = 'decimal'
): string {
  const lvl = (ilvl: number, start: number | null) =>
    `<w:lvl w:ilvl="${ilvl}">${start === null ? '' : `<w:start w:val="${start}"/>`}` +
    `<w:numFmt w:val="${ilvl === 0 ? level0Fmt : 'decimal'}"/>` +
    `<w:lvlText w:val="${TEMPLATES[ilvl]}"/></w:lvl>`;
  const nums = { 1: {}, ...overrides };
  const numXml = Object.entries(nums).map(
    ([numId, levels]) =>
      `<w:num w:numId="${numId}"><w:abstractNumId w:val="7"/>` +
      Object.entries(levels)
        .map(
          ([ilvl, start]) =>
            `<w:lvlOverride w:ilvl="${ilvl}"><w:startOverride w:val="${start}"/></w:lvlOverride>`
        )
        .join('') +
      `</w:num>`
  );
  return (
    `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:abstractNum w:abstractNumId="7">${lvl(0, level0Start)}${lvl(1, 1)}${lvl(2, 1)}</w:abstractNum>` +
    `${numXml.join('')}</w:numbering>`
  );
}

function parsedMarkers(xml: string, paras: Array<[numId: number, ilvl: number]>): string[] {
  const numbering = parseNumbering(xml);
  return markers(
    paras.map(([numId, ilvl]) => {
      const listRendering = computeListRendering({ numId, ilvl }, numbering);
      if (!listRendering) throw new Error(`no list rendering for ${numId}/${ilvl}`);
      return { ...item(numId, ilvl), listRendering };
    })
  );
}

/**
 * Each expected list is Word's own `list string` for the same paragraphs,
 * read back from minimal fixtures opened in Word for Mac. "Restart at 1" is a
 * new numId on the same abstractNum with a level-0 startOverride of 1. Word
 * fills a parent level nothing has numbered yet from the abstract `w:start`,
 * and applies a startOverride only when its own level first numbers a
 * paragraph of the numId (or a shallower level restarts it).
 */
describe('toFlowBlocks — startOverride and unused parent levels, as Word numbers them', () => {
  test('a Restart at 1 list after a numbered list continues level 0 until it reaches level 0', () => {
    expect(
      parsedMarkers(numberingXml(1, { 2: { 0: 1 } }), [
        [1, 0],
        [1, 0],
        [2, 1],
        [2, 1],
        [2, 0],
      ])
    ).toEqual(['1.', '2.', '2.1.', '2.2.', '1.']);
  });

  test('a Restart at 1 list that opens with sub-items repeats 1. at level 0', () => {
    expect(
      parsedMarkers(numberingXml(1, { 2: { 0: 1 } }), [
        [2, 1],
        [2, 1],
        [2, 0],
        [2, 1],
      ])
    ).toEqual(['1.1.', '1.2.', '1.', '1.1.']);
  });

  test('a Restart at 1 list after a numbered list numbers its later sub-items from the restart', () => {
    expect(
      parsedMarkers(numberingXml(1, { 2: { 0: 1 } }), [
        [1, 0],
        [1, 0],
        [2, 1],
        [2, 1],
        [2, 0],
        [2, 1],
      ])
    ).toEqual(['1.', '2.', '2.1.', '2.2.', '1.', '1.1.']);
  });

  test('a level-0 startOverride of 5 leaves opening sub-items under the abstract start', () => {
    expect(
      parsedMarkers(numberingXml(1, { 3: { 0: 5 } }), [
        [3, 1],
        [3, 1],
        [3, 0],
        [3, 1],
      ])
    ).toEqual(['1.1.', '1.2.', '5.', '5.1.']);
  });

  test('a level-0 w:start of 3 with a startOverride of 5 numbers opening sub-items from 3', () => {
    expect(
      parsedMarkers(numberingXml(3, { 3: { 0: 5 } }), [
        [3, 1],
        [3, 1],
        [3, 0],
      ])
    ).toEqual(['3.1.', '3.2.', '5.']);
  });

  test('a level-1 startOverride of 5 fills a restarted level 1 and applies again when level 1 is used', () => {
    expect(
      parsedMarkers(numberingXml(1, { 1: { 1: 5 } }), [
        [1, 0],
        [1, 2],
        [1, 1],
      ])
    ).toEqual(['1.', '1.5.1.', '1.5.']);
  });
});

describe('toFlowBlocks — shallower list levels take their start value', () => {
  test('a level-0 start of 5 numbers the opening sub-items 5.1., 5.2. and the parent 6.', () => {
    expect(
      parsedMarkers(numberingXml(5), [
        [1, 1],
        [1, 1],
        [1, 0],
        [1, 1],
      ])
    ).toEqual(['5.1.', '5.2.', '6.', '6.1.']);
  });

  test('a startOverride re-applies when a parent level restarts its level', () => {
    expect(
      parsedMarkers(numberingXml(1, { 2: { 1: 5 } }), [
        [2, 0],
        [2, 1],
        [2, 1],
        [2, 0],
        [2, 1],
      ])
    ).toEqual(['1.', '1.5.', '1.6.', '2.', '2.5.']);
  });

  test('a skipped level starts at its own start value', () => {
    expect(
      parsedMarkers(numberingXml(5), [
        [1, 0],
        [1, 2],
        [1, 1],
      ])
    ).toEqual(['5.', '5.1.1.', '5.2.']);
  });

  test('a level without w:start starts at 0, as in Word', () => {
    // ECMA-376 §17.9.26: an omitted w:start means 0.
    expect(
      parsedMarkers(numberingXml(null), [
        [1, 0],
        [1, 1],
        [1, 0],
      ])
    ).toEqual(['0.', '0.1.', '1.']);
    expect(
      parsedMarkers(numberingXml(null), [
        [1, 1],
        [1, 0],
      ])
    ).toEqual(['0.1.', '1.']);
  });

  test('a letter or roman level at 0 keeps its punctuation, as in Word', () => {
    // Roman and letter numerals have no zero. Word prints a space for the
    // number and keeps the "." after it.
    for (const [fmt, next] of [
      ['lowerLetter', 'a.'],
      ['upperLetter', 'A.'],
      ['lowerRoman', 'i.'],
      ['upperRoman', 'I.'],
    ]) {
      expect(
        parsedMarkers(numberingXml(null, {}, fmt), [
          [1, 0],
          [1, 0],
        ])
      ).toEqual([' .', next]);
    }
  });
});

describe('toFlowBlocks — huge list start values', () => {
  // Word documents no upper bound for a list start, so the layout caps it at
  // 32767. Uncapped, a start of 2147483647 made the roman formatter build a
  // marker of over two million characters per paragraph.
  const HUGE = 2147483647;
  const ROMAN_32767 = `${'M'.repeat(32)}DCCLXVII`;

  test('a huge w:start renders a short roman marker quickly', () => {
    const begin = performance.now();
    const out = parsedMarkers(numberingXml(HUGE, {}, 'upperRoman'), [
      [1, 0],
      [1, 0],
      [1, 1],
    ]);
    const elapsed = performance.now() - begin;
    expect(out).toEqual([
      `${ROMAN_32767}.`,
      `${'M'.repeat(32)}DCCLXVIII.`,
      `${'M'.repeat(32)}DCCLXVIII.1.`,
    ]);
    expect(elapsed).toBeLessThan(500);
  });

  test('a huge startOverride is capped the same way', () => {
    expect(
      parsedMarkers(numberingXml(1, { 2: { 0: HUGE } }), [
        [2, 0],
        [2, 0],
      ])
    ).toEqual(['32767.', '32768.']);
  });
});

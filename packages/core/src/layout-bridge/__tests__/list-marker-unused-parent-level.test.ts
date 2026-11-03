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

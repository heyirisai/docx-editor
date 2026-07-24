import { describe, expect, test } from 'bun:test';
import { Schema } from 'prosemirror-model';
import { EditorState } from 'prosemirror-state';
import { toFlowBlocks } from '../toFlowBlocks';
import type { ParagraphBlock } from '../../layout-engine/types';

const schema = new Schema({
  nodes: {
    doc: { content: 'paragraph+' },
    paragraph: {
      content: 'inline*',
      group: 'block',
      attrs: {
        paraId: { default: null },
      },
    },
    text: { group: 'inline' },
  },
});

function firstParagraph(blocks: unknown[]): ParagraphBlock {
  return blocks.find((b) => (b as ParagraphBlock).kind === 'paragraph') as ParagraphBlock;
}

describe('toFlowBlocks — paragraph identity', () => {
  test('carries stable PM paraId onto paragraph blocks', () => {
    const doc = schema.node('doc', null, [
      schema.node('paragraph', { paraId: '1A2B3C4D' }, [schema.text('Locate me')]),
    ]);

    const para = firstParagraph(toFlowBlocks(doc, {}));
    expect(para.paraId).toBe('1A2B3C4D');
  });

  test('keeps unchanged block IDs stable across layout passes', () => {
    const doc = schema.node('doc', null, [
      schema.node('paragraph', { paraId: 'FIRST' }, [schema.text('First paragraph')]),
      schema.node('paragraph', { paraId: 'SECOND' }, [schema.text('Second paragraph')]),
    ]);
    const state = EditorState.create({ doc });
    const before = toFlowBlocks(state.doc, {}) as ParagraphBlock[];

    const afterState = state.apply(state.tr.insertText('x', 2));
    const after = toFlowBlocks(afterState.doc, {}) as ParagraphBlock[];

    expect(after[0].id).not.toBe(before[0].id);
    expect(after[1].id).toBe(before[1].id);
  });
});

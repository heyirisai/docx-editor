/**
 * Switching a list to another one in suggesting mode clears the old list's
 * start values (they belong to its numId). Rejecting that switch must bring
 * them back with the old numId, or a list whose level 0 starts at 5 comes
 * back numbered from 1.
 */

import { describe, test, expect } from 'bun:test';
import { EditorState, TextSelection } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import { singletonManager } from '../../schema';
import { createSuggestionModePlugin } from '../../plugins/suggestionMode';
import { toggleNumberedList } from '../../commands/paragraph';
import { acceptChangeById, rejectChangeById } from '../../commands/comments';
import { toFlowBlocks } from '../../../layout-bridge/toFlowBlocks';

const schema = singletonManager.getSchema();

/** A list item of numId 5, whose level 0 starts at 5 (as parsed from w:start). */
function item(text: string): PMNode {
  return schema.node(
    'paragraph',
    {
      numPr: { numId: 5, ilvl: 0 },
      listIsBullet: false,
      listNumFmt: 'decimal',
      listLevelNumFmts: ['decimal'],
      listAbstractNumId: 5,
      listLevelStarts: [5, 0, 0, 0, 0, 0, 0, 0, 0],
      listLevelAbstractStarts: [5, 0, 0, 0, 0, 0, 0, 0, 0],
    },
    [schema.text(text)]
  );
}

function suggestingState(): EditorState {
  const doc = schema.node('doc', null, [item('first'), item('second')]);
  const state = EditorState.create({
    doc,
    schema,
    plugins: [...singletonManager.getPlugins(), createSuggestionModePlugin(true, 'Jane')],
  });
  return state.apply(
    state.tr.setSelection(TextSelection.create(state.doc, 1, state.doc.content.size - 1))
  );
}

function run(state: EditorState, command: typeof toggleNumberedList): EditorState {
  let next = state;
  command(state, (tr) => {
    next = state.apply(tr);
  });
  return next;
}

function markers(state: EditorState): string[] {
  const out: string[] = [];
  for (const block of toFlowBlocks(state.doc, {})) {
    if (block.kind === 'paragraph' && block.attrs?.listMarker) out.push(block.attrs.listMarker);
  }
  return out;
}

function revisionId(state: EditorState): number {
  const changes = state.doc.firstChild?.attrs.pPrChange as Array<{ info: { id: number } }>;
  return changes[0].info.id;
}

describe('toggleList — tracked switch to another list', () => {
  test('reject restores the old list and its start values', () => {
    const before = suggestingState();
    expect(markers(before)).toEqual(['5.', '6.']);

    const switched = run(before, toggleNumberedList);
    expect(switched.doc.firstChild?.attrs.numPr).toEqual({ numId: 2, ilvl: 0 });
    expect(switched.doc.firstChild?.attrs.listLevelStarts).toBeNull();
    expect(switched.doc.firstChild?.attrs.listLevelAbstractStarts).toBeNull();

    const rejected = run(switched, rejectChangeById(revisionId(switched)));
    expect(rejected.doc.firstChild?.attrs.numPr).toEqual({ numId: 5, ilvl: 0 });
    expect(rejected.doc.firstChild?.attrs.listLevelStarts).toEqual([5, 0, 0, 0, 0, 0, 0, 0, 0]);
    expect(rejected.doc.firstChild?.attrs.listLevelAbstractStarts).toEqual([
      5, 0, 0, 0, 0, 0, 0, 0, 0,
    ]);
    expect(markers(rejected)).toEqual(['5.', '6.']);
  });

  test('accept keeps the new list without the old start values', () => {
    const switched = run(suggestingState(), toggleNumberedList);
    const accepted = run(switched, acceptChangeById(revisionId(switched)));
    expect(accepted.doc.firstChild?.attrs.numPr).toEqual({ numId: 2, ilvl: 0 });
    expect(accepted.doc.firstChild?.attrs.listLevelStarts).toBeNull();
    expect(accepted.doc.firstChild?.attrs.listLevelAbstractStarts).toBeNull();
    expect(accepted.doc.firstChild?.attrs.pPrChange).toBeNull();
    expect(markers(accepted)).toEqual(['1.', '2.']);
  });
});

import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, beforeAll, describe, expect, test } from 'bun:test';

import { renderTextBoxFragment } from '../renderTextBox';
import type { TextBoxBlock, TextBoxFragment, TextBoxMeasure } from '../../layout-engine/types';
import type { RenderContext } from '../renderPage';

beforeAll(() => GlobalRegistrator.register());
afterAll(() => GlobalRegistrator.unregister());

/**
 * A decorative filled shape reaches the painter as a `renderOnly` text box
 * (see `isFilledShapeDrawing`). It carries an empty paragraph through
 * ProseMirror because the node schema demands one, and export drops it — so
 * anything typed into it would be lost. The painter therefore draws the fill
 * and nothing else: no pointer target, and a marker the interaction
 * hit-tests skip, matching `ImageBlock.renderOnly`.
 */
function paint(renderOnly: boolean): HTMLElement {
  const fragment = {
    kind: 'textBox',
    blockId: 'b1',
    x: 0,
    y: 0,
    width: 120,
    height: 60,
  } as TextBoxFragment;
  const block = {
    kind: 'textBox',
    id: 'b1',
    width: 120,
    height: 60,
    fillColor: '#304050',
    content: [],
    renderOnly: renderOnly || undefined,
  } as TextBoxBlock;
  const measure = {
    kind: 'textBox',
    innerMeasures: [],
    totalHeight: 60,
    width: 120,
    height: 60,
  } as TextBoxMeasure;
  return renderTextBoxFragment(fragment, block, measure, {} as RenderContext);
}

describe('render-only text box', () => {
  test('paints its fill but takes no clicks', () => {
    const el = paint(true);
    expect(el.style.backgroundColor).toBeTruthy();
    expect(el.dataset.renderOnly).toBe('1');
    expect(el.style.pointerEvents).toBe('none');
  });

  test('an ordinary text box stays interactive', () => {
    const el = paint(false);
    expect(el.dataset.renderOnly).toBeUndefined();
    expect(el.style.pointerEvents).toBe('');
  });
});

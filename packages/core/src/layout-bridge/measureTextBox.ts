/**
 * Shared text-box measurement helper.
 *
 * Both React's PagedEditor and Vue's useDocxEditor measured a text box by
 * mapping `measureParagraph` over its content. That silently produced a
 * zero-height measure for anything that was not a paragraph, and
 * `w:txbxContent` is EG_BlockLevelElts — a box may hold tables. This module
 * lives in core so the two adapters cannot drift; the per-block work is
 * delegated back through `measureBlock` because each adapter's block coverage
 * differs.
 */

import {
  DEFAULT_TEXTBOX_MARGINS,
  DEFAULT_TEXTBOX_WIDTH,
  type Measure,
  type ParagraphMeasure,
  type TableMeasure,
  type TextBoxBlock,
  type TextBoxMeasure,
} from '../layout-engine/types';

/** Measure one text box, including any tables inside it. */
export function measureTextBoxBlock(
  block: TextBoxBlock,
  measureBlock: (inner: TextBoxBlock['content'][number], width: number) => Measure
): TextBoxMeasure {
  const margins = block.margins ?? DEFAULT_TEXTBOX_MARGINS;
  const width = block.width ?? DEFAULT_TEXTBOX_WIDTH;
  const innerWidth = Math.max(1, width - margins.left - margins.right);

  const innerMeasures: Array<ParagraphMeasure | TableMeasure> = [];
  for (const inner of block.content) {
    const measure = measureBlock(inner, innerWidth);
    if (measure.kind === 'paragraph' || measure.kind === 'table') {
      innerMeasures.push(measure);
    }
  }

  const contentHeight = innerMeasures.reduce((sum, m) => sum + m.totalHeight, 0);
  return {
    kind: 'textBox',
    width,
    // A declared `cy` wins: Word clips a box to its authored height rather
    // than growing it to fit (see `renderTextBox`'s inset clamp).
    height: block.height ?? contentHeight + margins.top + margins.bottom,
    innerMeasures,
  };
}

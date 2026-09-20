/**
 * Text Box Renderer
 *
 * Renders text box fragments to DOM. Handles:
 * - Background fill color
 * - Border/outline
 * - Internal padding (margins)
 * - Paragraph content inside the box (using pre-measured data)
 */

import {
  DEFAULT_TEXTBOX_MARGINS,
  type TextBoxFragment,
  type TextBoxBlock,
  type TextBoxMeasure,
} from '../layout-engine/types';
import type { RenderContext } from './renderPage';
import { renderParagraphFragment } from './renderParagraph';
import { renderTableFragment } from './renderTable';

/** Word's default `roundRect` corner adjust — see {@link TextBoxBlock.cornerAdj}. */
const DEFAULT_ROUND_RECT_ADJ = 0.16667;

/**
 * CSS class names for text box elements
 */
export const TEXTBOX_CLASS_NAMES = {
  textBox: 'layout-textbox',
};

/**
 * Options for rendering a text box fragment
 */
export interface RenderTextBoxFragmentOptions {
  document?: Document;
}

/**
 * Render a text box fragment to DOM
 */
export function renderTextBoxFragment(
  fragment: TextBoxFragment,
  block: TextBoxBlock,
  measure: TextBoxMeasure,
  context: RenderContext,
  options: RenderTextBoxFragmentOptions = {}
): HTMLElement {
  const doc = options.document ?? document;

  const containerEl = doc.createElement('div');
  containerEl.className = TEXTBOX_CLASS_NAMES.textBox;

  // Basic styling
  containerEl.style.position = 'absolute';
  containerEl.style.width = `${fragment.width}px`;
  containerEl.style.height = `${fragment.height}px`;
  containerEl.style.overflow = 'hidden';
  containerEl.style.boxSizing = 'border-box';
  applyTextBoxStacking(containerEl, fragment);

  // Fill color
  if (block.fillColor) {
    containerEl.style.backgroundColor = block.fillColor;
  }

  // `a:prstGeom` rounding. Word draws the preset, not its bounding box, so an
  // `ellipse` badge and a `roundRect` pill both need a radius here — without
  // one they painted as hard-cornered rectangles.
  const radius = cornerRadiusCss(block, fragment);
  if (radius) containerEl.style.borderRadius = radius;

  // A canvas-only frame is a decoration, not a box anyone edits: let clicks
  // reach the body text painted over it, and mark it so the image/textbox
  // interaction hit-tests skip it (same contract as `ImageBlock.renderOnly`).
  if (block.renderOnly) {
    containerEl.dataset.renderOnly = '1';
    containerEl.style.pointerEvents = 'none';
  }

  // Store metadata
  containerEl.dataset.blockId = String(fragment.blockId);
  if (fragment.pmStart !== undefined) {
    containerEl.dataset.pmStart = String(fragment.pmStart);
  }
  if (fragment.pmEnd !== undefined) {
    containerEl.dataset.pmEnd = String(fragment.pmEnd);
  }

  // Border/outline. A connector strokes ONE line across the extent box
  // instead — see `renderConnectorStroke`.
  if (block.lineShape) {
    renderConnectorStroke(containerEl, fragment, block, doc);
    return containerEl;
  }
  if (block.outlineWidth && block.outlineWidth > 0) {
    const style = block.outlineStyle || 'solid';
    const color = block.outlineColor || '#000000';
    containerEl.style.border = `${block.outlineWidth}px ${style} ${color}`;
  }

  // Internal padding. `wps:bodyPr` insets never GROW the shape in Word — a
  // 0.35in-tall accent bar with a 0.2in top inset stays 0.35in tall and clips.
  // With `box-sizing: border-box` the browser floors an element at its own
  // padding, so insets taller than the declared box pushed it out to 57px and
  // it covered the panel title underneath.
  const margins = block.margins ?? DEFAULT_TEXTBOX_MARGINS;
  const padTop = Math.max(0, Math.min(margins.top, fragment.height));
  const padBottom = Math.max(0, Math.min(margins.bottom, fragment.height - padTop));
  containerEl.style.padding = `${padTop}px ${margins.right}px ${padBottom}px ${margins.left}px`;

  // Render inner content using pre-measured data
  const innerWidth = fragment.width - margins.left - margins.right;
  let yOffset = 0;

  for (let i = 0; i < block.content.length; i++) {
    const paraBlock = block.content[i];
    const paraMeasure = measure.innerMeasures[i];
    if (!paraMeasure) continue;

    // A table inside the box (`w:txbxContent` is EG_BlockLevelElts) paints
    // through the normal table renderer, stacked in the box's flow.
    if (paraBlock.kind === 'table') {
      if (paraMeasure.kind !== 'table') continue;
      const tableEl = renderTableFragment(
        {
          kind: 'table',
          blockId: paraBlock.id,
          x: 0,
          y: yOffset,
          width: innerWidth,
          height: paraMeasure.totalHeight,
          fromRow: 0,
          toRow: paraBlock.rows.length,
          pmStart: paraBlock.pmStart,
          pmEnd: paraBlock.pmEnd,
        },
        paraBlock,
        paraMeasure,
        { ...context, positioning: 'flow' },
        { document: doc }
      );
      containerEl.appendChild(tableEl);
      yOffset += paraMeasure.totalHeight;
      continue;
    }
    if (paraMeasure.kind !== 'paragraph') continue;

    const paraFragment = {
      kind: 'paragraph' as const,
      blockId: paraBlock.id,
      x: 0,
      y: yOffset,
      width: innerWidth,
      height: paraMeasure.totalHeight,
      pmStart: paraBlock.pmStart,
      pmEnd: paraBlock.pmEnd,
      fromLine: 0,
      toLine: paraMeasure.lines.length,
    };

    // Pass `positioning: 'flow'` so the renderer's outer position is
    // explicit. `renderParagraphFragment` already defaults to `position:
    // relative` (it needs to be a containing block for floating images),
    // so passing 'flow' here is documentation more than behavior change —
    // pre-PR the textbox caller re-set the same `position: relative; top:
    // 0; left: 0` after the renderer call (#379).
    const paraEl = renderParagraphFragment(
      paraFragment,
      paraBlock,
      paraMeasure,
      { ...context, positioning: 'flow' },
      { document: doc }
    );

    containerEl.appendChild(paraEl);
    yOffset += paraMeasure.totalHeight;
  }

  return containerEl;
}

/**
 * CSS `border-radius` for a preset geometry, or `undefined` for a plain
 * rectangle. A `roundRect`'s corner radius is a fraction of the SHORTER side
 * (§20.1.9.11), so a wide, short button rounds to a pill at `adj = 0.5`.
 */
function cornerRadiusCss(block: TextBoxBlock, fragment: TextBoxFragment): string | undefined {
  if (block.geometry === 'ellipse') return '50%';
  if (block.geometry !== 'roundRect') return undefined;
  const shortSide = Math.max(0, Math.min(fragment.width, fragment.height));
  const adj = block.cornerAdj ?? DEFAULT_ROUND_RECT_ADJ;
  return `${Math.round(shortSide * adj * 100) / 100}px`;
}

function applyTextBoxStacking(element: HTMLElement, fragment: TextBoxFragment): void {
  if (!fragment.isFloating && fragment.zIndex === undefined) return;
  element.style.zIndex = String(fragment.zIndex ?? 10);
}

/**
 * Paint a stroke-only connector (`Shape.lineShape`): a single line between
 * two corners of the extent box, rotated to the box's diagonal.
 *
 * The common case by far is a footer rule — `prst="line"` with `cy="0"` —
 * which degenerates to a plain horizontal border with no rotation. Drawing
 * the bounding box instead put a full-width rectangle on every page.
 */
function renderConnectorStroke(
  containerEl: HTMLElement,
  fragment: TextBoxFragment,
  block: TextBoxBlock,
  doc: Document
): void {
  // The container is just the anchor point; the stroke is its own element so
  // the rotation cannot disturb the anchored position.
  containerEl.style.overflow = 'visible';
  containerEl.style.padding = '0';
  containerEl.style.backgroundColor = 'transparent';

  const dx = fragment.width;
  const dy = fragment.height;
  const length = Math.sqrt(dx * dx + dy * dy);
  if (length <= 0) return;

  const line = doc.createElement('div');
  line.className = 'layout-textbox-stroke';
  line.style.position = 'absolute';
  line.style.left = '0';
  // `lineShape === 'up'` (a:xfrm/@flipV) runs bottom-left to top-right.
  line.style.top = block.lineShape === 'up' ? `${dy}px` : '0';
  line.style.width = `${length}px`;
  line.style.height = '0';
  line.style.borderTop = `${block.outlineWidth && block.outlineWidth > 0 ? block.outlineWidth : 1}px ${
    block.outlineStyle || 'solid'
  } ${block.outlineColor || '#000000'}`;
  const angle = (Math.atan2(block.lineShape === 'up' ? -dy : dy, dx) * 180) / Math.PI;
  if (angle !== 0) {
    line.style.transformOrigin = '0 0';
    line.style.transform = `rotate(${angle}deg)`;
  }
  containerEl.appendChild(line);
}

/**
 * Floating-image extraction for table cells.
 *
 * Pulls anchored/floating images out of a cell's paragraphs and computes their
 * positions relative to the cell content area. Split out of renderTable.ts so
 * that file stays focused on row/cell/fragment painting.
 */

import type {
  ImageRun,
  ParagraphBlock,
  ParagraphMeasure,
  TableCell,
  TableCellMeasure,
  TableMeasure,
} from '../layout-engine/types';
import type { RenderContext } from './renderPage';
import { emuToPixels } from '../utils/units';
import { renderFloatingImagesLayer } from './floatingImageLayer';
import {
  floatingImageIsBehindDoc,
  imageWrapTextFromCssFloat,
  isFloatingImageRun,
} from './floatingImageFlow';

/** Info about a floating image extracted from a cell paragraph */
export interface CellFloatingImage {
  /** Painted from a preserved group; marked so hit-testing skips it. */
  renderOnly?: boolean;
  src: string;
  assetId?: string;
  width: number;
  height: number;
  alt?: string;
  transform?: string;
  x: number;
  y: number;
  side: 'left' | 'right';
  distTop: number;
  distBottom: number;
  distLeft: number;
  distRight: number;
  /** OOXML wrapText: which side(s) TEXT flows on */
  wrapText?: 'bothSides' | 'left' | 'right' | 'largest';
  /** Wrap type (square, tight, through, behind, inFront) */
  wrapType?: string;
  /**
   * The picture's own appearance, forwarded to `applyImageVisualAttrs` by the
   * floating-image layer. A cell float is still a picture: COMET's bio
   * headshots are `srcRect`-cropped AND `a:prstGeom prst="ellipse"`, and
   * dropping these painted them as stretched squares.
   */
  cropTop?: number;
  cropRight?: number;
  cropBottom?: number;
  cropLeft?: number;
  opacity?: number;
  geometry?: 'ellipse' | 'roundRect';
  cornerAdj?: number;
  pmStart?: number;
  pmEnd?: number;
}

/**
 * Extract floating images from cell paragraphs and compute their positions
 * relative to the cell content area.
 *
 * NOTE: The horizontal/vertical position logic here mirrors
 * extractFloatingImagesFromParagraph() in renderPage.ts. Kept separate
 * because the coordinate systems differ (cell-relative vs page-relative).
 */
export function extractCellFloatingImages(
  cell: TableCell,
  cellMeasure: TableCellMeasure,
  contentWidth: number
): CellFloatingImage[] {
  const result: CellFloatingImage[] = [];
  let paragraphY = 0;

  for (let blockIndex = 0; blockIndex < cell.blocks.length; blockIndex++) {
    const block = cell.blocks[blockIndex];
    if (block?.kind !== 'paragraph') {
      // Use actual measured height for Y tracking
      const blockMeasure = cellMeasure.blocks[blockIndex];
      if (blockMeasure && blockMeasure.kind === 'table') {
        paragraphY += (blockMeasure as TableMeasure).totalHeight ?? 0;
      }
      continue;
    }
    const pBlock = block as ParagraphBlock;

    for (const run of pBlock.runs) {
      if (run.kind !== 'image') continue;
      const imgRun = run as ImageRun;
      if (!isFloatingImageRun(imgRun)) continue;

      const position = imgRun.position;
      const distTop = imgRun.distTop ?? 0;
      const distBottom = imgRun.distBottom ?? 0;
      const distLeft = imgRun.distLeft ?? 12;
      const distRight = imgRun.distRight ?? 12;

      // Horizontal position within cell
      let side: 'left' | 'right' = 'left';
      let x = 0;

      if (position?.horizontal) {
        const h = position.horizontal;
        if (h.align === 'right') {
          side = 'right';
          x = contentWidth - imgRun.width;
        } else if (h.align === 'left') {
          x = 0;
        } else if (h.align === 'center') {
          x = (contentWidth - imgRun.width) / 2;
        } else if (h.posOffset !== undefined) {
          x = emuToPixels(h.posOffset);
          side = x > contentWidth / 2 ? 'right' : 'left';
        }
      } else if (imgRun.cssFloat === 'right') {
        side = 'right';
        x = contentWidth - imgRun.width;
      }

      // Vertical position within cell
      let y = paragraphY;
      if (position?.vertical) {
        const v = position.vertical;
        if (v.posOffset !== undefined) {
          y = paragraphY + emuToPixels(v.posOffset);
        } else if (v.align === 'top') {
          y = 0;
        }
      }

      // NOT clamped into the cell. `layoutInCell="1"` says the anchor is
      // resolved against the CELL, not that the object is confined to it —
      // measured in Word, a picture anchored in a 1in cell paints at its full
      // 2.5in width across its neighbours and below the row. Clamping threw
      // away LCPS's header logo offset (`posOffset="-114300"`, -12px) and
      // pushed it 12px right, straight into the table's clip edge.

      result.push({
        src: imgRun.src,
        assetId: imgRun.assetId,
        width: imgRun.width,
        height: imgRun.height,
        alt: imgRun.alt,
        transform: imgRun.transform,
        x,
        y,
        side,
        distTop,
        distBottom,
        distLeft,
        distRight,
        wrapText: imageWrapTextFromCssFloat(imgRun.cssFloat),
        wrapType: imgRun.wrapType,
        cropTop: imgRun.cropTop,
        cropRight: imgRun.cropRight,
        cropBottom: imgRun.cropBottom,
        cropLeft: imgRun.cropLeft,
        opacity: imgRun.opacity,
        geometry: imgRun.geometry,
        cornerAdj: imgRun.cornerAdj,
        pmStart: imgRun.pmStart,
        pmEnd: imgRun.pmEnd,
        renderOnly: imgRun.renderOnly,
      });
    }

    // Use actual measured height for Y tracking
    const blockMeasure = cellMeasure.blocks[blockIndex];
    if (blockMeasure && blockMeasure.kind === 'paragraph') {
      paragraphY += (blockMeasure as ParagraphMeasure).totalHeight;
    }
  }

  return result;
}

/** Cell padding actually painted (Word's visual defaults when unset). */
export function cellPadding(cell: TableCell): {
  top: number;
  right: number;
  bottom: number;
  left: number;
} {
  return {
    top: cell.padding?.top ?? 1,
    right: cell.padding?.right ?? 7,
    bottom: cell.padding?.bottom ?? 1,
    left: cell.padding?.left ?? 7,
  };
}

/** Width of a cell's content box (its border box less horizontal padding). */
export function cellContentWidth(cell: TableCell, cellMeasure: TableCellMeasure): number {
  const pad = cellPadding(cell);
  return Math.max(0, cellMeasure.width - pad.left - pad.right);
}

/**
 * A cell's IN-FRONT anchored pictures, lifted out of the cell so nothing clips
 * them.
 *
 * Word resolves `layoutInCell="1"` against the cell but does not confine the
 * object to it — measured: a picture anchored in a 1in cell paints at its full
 * 2.5in width, across its neighbours and below the row. Our cell box carries
 * `overflow: hidden` (it must, so over-tall content and vertically-merged
 * continuation slices clip), and so does the table element (it is the
 * page-break row window). A layer painted inside the cell therefore loses
 * whatever pokes out — LCPS's header logo lost its final letter and the mark
 * above it. The ROW clips nothing, so the layer hangs off the row instead,
 * positioned at the cell's content origin.
 *
 * `behindDoc="1"` pictures deliberately stay INSIDE the cell
 * (`renderCellContent` paints them). Their whole point is to sit under the
 * cell's text, and among positioned siblings paint order is DOM order — a
 * hoisted layer placed before the cell would end up under the cell's own
 * BACKGROUND too, so a shaded cell would swallow the picture entirely. Keeping
 * them in the cell preserves "above the fill, below the text"; the cost is
 * that a behind-mode picture is still clipped to its cell.
 */
export interface CellFloatLayer {
  el: HTMLElement;
  /** Left within the ROW: the cell's x plus its left padding. */
  left: number;
  /** Top within the CELL box: its top padding plus any vertical-align shift. */
  top: number;
}

export function buildCellFloatLayers(
  cell: TableCell,
  cellMeasure: TableCellMeasure,
  cellX: number,
  verticalShift: number,
  context: RenderContext,
  doc: Document
): CellFloatLayer[] {
  const records = extractCellFloatingImages(
    cell,
    cellMeasure,
    cellContentWidth(cell, cellMeasure)
  ).filter((img) => !floatingImageIsBehindDoc(img));
  if (records.length === 0) return [];
  const pad = cellPadding(cell);
  return [
    {
      el: renderFloatingImagesLayer(records, doc, {
        layerClass: 'layout-cell-floating-images-layer',
        itemClass: 'layout-cell-floating-image',
        sizing: 'origin',
        layerMode: 'front',
        imageAssetLoader: context.imageAssetLoader,
      }),
      left: cellX + pad.left,
      top: pad.top + verticalShift,
    },
  ];
}

/**
 * Where the cell's content box starts vertically inside its border box. Mirrors
 * the `justify-content` the cell applies below, so a hoisted float tracks the
 * content it was measured against.
 */
export function cellContentShift(
  cell: TableCell,
  cellMeasure: TableCellMeasure,
  rowHeight: number,
  contentFillsBox: boolean
): number {
  if (!cell.verticalAlign || contentFillsBox) return 0;
  // `cellMeasure.height` ALREADY includes the cell's top and bottom padding
  // (measureTable adds them), so the slack is against the raw row height —
  // subtracting the padding again here shifted a centred picture up by half
  // the padding and a bottom-aligned one by all of it.
  const slack = Math.max(0, rowHeight - (cellMeasure.height ?? 0));
  if (cell.verticalAlign === 'center') return slack / 2;
  if (cell.verticalAlign === 'bottom') return slack;
  return 0;
}

/**
 * Place a cell's hoisted float layers next to the cell they belong to.
 * `cellTop` is the cell's own top in `parent` coordinates (0 inside a row).
 * They are appended AFTER the cell: these are in-front pictures, so painting
 * over the cell is what Word does.
 */
export function appendCellFloatLayers(
  parent: HTMLElement,
  layers: CellFloatLayer[],
  cellTop: number
): void {
  for (const layer of layers) {
    layer.el.style.left = `${layer.left}px`;
    layer.el.style.top = `${cellTop + layer.top}px`;
    parent.appendChild(layer.el);
  }
}

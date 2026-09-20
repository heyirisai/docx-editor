/**
 * Header / footer rendering for renderPage.
 *
 * Owns `renderHeaderFooterContent` — the mini-flow that lays paragraphs and
 * tables inside a header/footer container (separate from the body flow) —
 * plus the floating-image and floating-table positioning helpers used by
 * that flow. Coordinates returned by `resolveHeaderFooterFloatingTablePosition`
 * are relative to the HF container's flow origin (`layout.flowTop`/`flowLeft`)
 * so callers can drop them into `style.top`/`style.left`.
 */

import type {
  FlowBlock,
  ImageRun,
  Measure,
  ParagraphBlock,
  ParagraphFragment,
  TableBlock,
  TableFragment,
  ImageFragment,
  TextBoxFragment,
} from '../../layout-engine/types';
import { assertExhaustiveFlowBlock } from '../../layout-engine/types';
import { anchoredTopInHeaderFooterBand } from '../../layout-bridge/headerFooterLayout';
import { renderParagraphFragment } from '../renderParagraph';
import { renderTableFragment } from '../renderTable';
import {
  applyImageVisualAttrs,
  hasImageVisualAttrs,
  imageCornerRadiusCss,
  renderImageFragment,
} from '../renderImage';
import { renderTextBoxFragment } from '../renderTextBox';
import { emuToPixels } from '../../utils/units';
import { headerFooterFrontZIndex } from '../../layout-engine/zOrder';
import type { RenderContext, RenderPageOptions } from '../renderPage';
import { setImageAssetSource } from '../imageAssets';

/**
 * Header/footer content for rendering
 */
export interface HeaderFooterContent {
  /** Flow blocks for the header/footer content. */
  blocks: FlowBlock[];
  /** Measurements for the blocks. */
  measures: Measure[];
  /** Total height of the content (in-flow stack incl. floating blocks). */
  height: number;
  /**
   * In-flow band height: the height of strictly in-flow content
   * (paragraphs, tables, inline images/text boxes), EXCLUDING anchored /
   * floating objects. This is what grows the header/footer band and pushes
   * the body margin, mirroring Word: a page/margin-anchored shape (e.g. a
   * full-page letterhead in a header) is positioned independently and does
   * NOT push body text down. Use this — not `height`/`visualBottom` — for
   * margin extension. Falls back to `height` when undefined.
   */
  flowHeight?: number;
  /** Top-most visual extent relative to the nominal flow origin. */
  visualTop?: number;
  /** Bottom-most visual extent relative to the nominal flow origin. */
  visualBottom?: number;
}

/**
 * One section's resolved header and footer, as the painter needs them.
 *
 * Word resolves headers per section, not per document: a cover section that
 * declares none must not borrow the body's, and the body's must not be lost
 * because some other section was resolved instead.
 *
 * @public
 */
export interface SectionHeaderFooterContent {
  header?: HeaderFooterContent;
  footer?: HeaderFooterContent;
  /** Used on the section's first page when `titlePg` is set. */
  firstHeader?: HeaderFooterContent;
  firstFooter?: HeaderFooterContent;
  /** `w:titlePg` for this section. */
  titlePg?: boolean;
  /** `w:header` — distance from the page top to the header content. */
  headerDistance?: number;
  /** `w:footer` — distance from the page bottom to the footer content. */
  footerDistance?: number;
}

export interface HeaderFooterLayoutInfo {
  flowTop: number;
  flowLeft: number;
  contentWidth: number;
  pageWidth: number;
  pageHeight: number;
  margins: {
    top: number;
    right: number;
    bottom: number;
    left: number;
  };
}

function getPositionAlignment(
  position: { align?: string; alignment?: string } | undefined
): string | undefined {
  return position?.align ?? position?.alignment;
}

function resolveHeaderFooterFloatTop(
  floatImg: {
    height: number;
    paragraphY: number;
    position: {
      vertical?: { relativeTo?: string; posOffset?: number; align?: string; alignment?: string };
    };
  },
  layout: HeaderFooterLayoutInfo
): number {
  // One implementation, shared with the measurement pass — the two had
  // byte-identical copies of this and could drift apart.
  return anchoredTopInHeaderFooterBand(
    floatImg.position.vertical,
    floatImg.paragraphY,
    floatImg.height,
    { flowTop: layout.flowTop, pageHeight: layout.pageHeight, margins: layout.margins }
  );
}

/**
 * Resolve the CSS `left` (px) for an anchored object (image or text box) in a
 * header/footer, honoring `wp:positionH` (relativeTo page/margin, align
 * left/center/right, or posOffset). Shared by floating images and text boxes so
 * a page-centered text box in the header lands centered like Word, not pinned
 * to the left.
 */
export function resolveHeaderFooterFloatLeft(
  width: number,
  h: { relativeTo?: string; posOffset?: number; align?: string; alignment?: string } | undefined,
  layout: HeaderFooterLayoutInfo
): string {
  if (!h) return '0';
  const align = getPositionAlignment(h);

  if (h.relativeTo === 'page') {
    if (h.posOffset !== undefined) return `${emuToPixels(h.posOffset) - layout.flowLeft}px`;
    if (align === 'right') return `${layout.pageWidth - width - layout.flowLeft}px`;
    if (align === 'center') return `${(layout.pageWidth - width) / 2 - layout.flowLeft}px`;
    if (align === 'left') return `${-layout.flowLeft}px`;
  }

  // `relativeTo: margin` falls through here intentionally: the HF content width
  // IS the margin box, so the content-relative branch is already margin-correct.
  if (h.posOffset !== undefined) return `${emuToPixels(h.posOffset)}px`;
  if (align === 'right') return `${layout.contentWidth - width}px`;
  if (align === 'center') return `${(layout.contentWidth - width) / 2}px`;

  return '0';
}

function applyHeaderFooterFloatHorizontalPosition(
  img: HTMLImageElement,
  floatImg: {
    width: number;
    position: {
      horizontal?: { relativeTo?: string; posOffset?: number; align?: string; alignment?: string };
    };
  },
  layout: HeaderFooterLayoutInfo
): void {
  img.style.left = resolveHeaderFooterFloatLeft(
    floatImg.width,
    floatImg.position.horizontal,
    layout
  );
}

/**
 * Resolve the (left, top) position for a floating table inside a header/
 * footer container, per ECMA-376 §17.4.57. The table's `floating.tblpX/tblpY`
 * are already in pixels (parser converted from twips); `horzAnchor`/
 * `vertAnchor` decide whether the offset is relative to the page, the
 * margins, or the surrounding text/column. Coordinates returned are
 * relative to the HF container's flow origin (`layout.flowTop` /
 * `layout.flowLeft`) so the caller can drop them straight into
 * `style.top` / `style.left`.
 */
export function resolveHeaderFooterFloatingTablePosition(
  floating: NonNullable<TableBlock['floating']>,
  layout: HeaderFooterLayoutInfo
): { left: number; top: number } {
  // Vertical: tblpY relative to vertAnchor.
  let top = floating.tblpY ?? 0;
  if (floating.vertAnchor === 'page') {
    top -= layout.flowTop;
  } else if (floating.vertAnchor === 'margin') {
    top += layout.margins.top - layout.flowTop;
  }

  // Horizontal: tblpX relative to horzAnchor.
  let left = floating.tblpX ?? 0;
  if (floating.horzAnchor === 'page') {
    left -= layout.flowLeft;
  } else if (floating.horzAnchor === 'margin') {
    left += layout.margins.left - layout.flowLeft;
  }

  return { left, top };
}

/**
 * Render header or footer content
 */
export function renderHeaderFooterContent(
  content: HeaderFooterContent,
  context: RenderContext,
  options: RenderPageOptions,
  layout: HeaderFooterLayoutInfo
): HTMLElement {
  const doc = options.document ?? document;
  const containerEl = doc.createElement('div');
  containerEl.style.position = 'relative';

  // Use content width from context if available, otherwise default to reasonable width
  const contentWidth = context.contentWidth ?? 600;

  // Collect floating images to render separately, with their paragraph's Y
  // position. Keep the FULL ImageRun on the record: an earlier narrow
  // re-pack (src/width/height/alt/position only) silently dropped the
  // OOXML srcRect crop, opacity, and transform — a badge strip cropped out
  // of a full-page screenshot rendered as the whole screenshot squashed
  // into the display box.
  const floatingImages: Array<{
    run: ImageRun;
    src: string;
    width: number;
    height: number;
    alt?: string;
    paragraphY: number; // Y position of the containing paragraph
    position: {
      horizontal?: {
        relativeTo?: string;
        posOffset?: number;
        align?: string;
        alignment?: string;
      };
      vertical?: {
        relativeTo?: string;
        posOffset?: number;
        align?: string;
        alignment?: string;
      };
    };
  }> = [];

  let cursorY = 0;

  for (let i = 0; i < content.blocks.length; i++) {
    const block = content.blocks[i];
    const measure = content.measures[i];
    if (!block || !measure) continue;

    if (block.kind === 'paragraph') {
      if (measure.kind !== 'paragraph') continue;
      const paragraphBlock = block;
      const paragraphMeasure = measure;
      const paragraphSpacingBefore = paragraphBlock.attrs?.spacing?.before ?? 0;

      // Track the Y position where this paragraph starts
      const paragraphStartY = cursorY;

      // Extract floating images and filter them from runs
      const inlineRuns: typeof paragraphBlock.runs = [];
      for (const run of paragraphBlock.runs) {
        if (run.kind === 'image' && 'position' in run && run.position) {
          const imgRun = run as ImageRun & {
            position: NonNullable<ImageRun['position']>;
          };
          floatingImages.push({
            run: imgRun,
            src: imgRun.src,
            width: imgRun.width,
            height: imgRun.height,
            alt: imgRun.alt,
            paragraphY: paragraphStartY, // Store where this paragraph starts
            position: imgRun.position,
          });
        } else {
          // Keep non-floating runs for inline rendering
          inlineRuns.push(run);
        }
      }

      // Create a modified paragraph block without floating images
      const inlineBlock: ParagraphBlock = {
        ...paragraphBlock,
        runs: inlineRuns,
      };

      // Create a synthetic fragment for the paragraph. `pmStart` / `pmEnd`
      // are essential for HF caret resolution — without them the painter
      // emits no `data-pm-*` markers on this paragraph wrapper, and empty
      // paragraphs (or cursors at line boundaries) lose any anchor at all.
      // `computeHfCaretRectFromView`'s fallback chain depends on these.
      const syntheticFragment: ParagraphFragment = {
        kind: 'paragraph',
        blockId: paragraphBlock.id,
        x: 0,
        y: cursorY + paragraphSpacingBefore,
        width: contentWidth,
        height: paragraphMeasure.totalHeight,
        fromLine: 0,
        toLine: paragraphMeasure.lines.length,
        pmStart: paragraphBlock.pmStart,
        pmEnd: paragraphBlock.pmEnd,
      };

      // Render paragraph fragment (with floating images filtered out). The
      // HF context positions blocks absolutely within its own container,
      // stacking vertically via `cursorY` — `paragraphMeasure.totalHeight`
      // already includes `spaceBefore` / `spaceAfter`. Pass `positioning:
      // 'absolute'` so the renderer applies that mode itself instead of the
      // caller having to flip its inline style after the fact (#379).
      const fragEl = renderParagraphFragment(
        syntheticFragment,
        inlineBlock,
        paragraphMeasure,
        { ...context, positioning: 'absolute' },
        { document: doc }
      );

      fragEl.style.top = `${cursorY + paragraphSpacingBefore}px`;
      fragEl.style.left = '0';
      fragEl.style.width = `${contentWidth}px`;

      containerEl.appendChild(fragEl);
      cursorY += paragraphMeasure.totalHeight;
    } else if (block.kind === 'table') {
      if (measure.kind !== 'table') continue;
      // HF tables don't paginate, so the synthetic fragment covers all rows.
      const syntheticFragment: TableFragment = {
        kind: 'table',
        blockId: block.id,
        x: 0,
        y: cursorY,
        width: measure.totalWidth,
        height: measure.totalHeight,
        fromRow: 0,
        toRow: measure.rows.length,
        pmStart: block.pmStart,
        pmEnd: block.pmEnd,
      };
      const fragEl = renderTableFragment(
        syntheticFragment,
        block,
        measure,
        { ...context, positioning: 'absolute' },
        { document: doc }
      );

      // Floating tables (`<w:tblpPr>`) opt out of the cursorY flow. They
      // anchor at (tblpX, tblpY) relative to the page/margin/column per
      // ECMA-376 §17.4.57 and don't advance cursorY (#382). Inline tables
      // keep their cursorY-based stacking.
      if (block.floating) {
        const { left, top } = resolveHeaderFooterFloatingTablePosition(block.floating, layout);
        fragEl.style.top = `${top}px`;
        fragEl.style.left = `${left}px`;
        containerEl.appendChild(fragEl);
        // Floating tables do NOT advance cursorY — surrounding HF blocks
        // flow as if the table weren't there. Word renders text behind
        // floating tables when no wrap behavior is requested; we match.
      } else {
        // Inline placement: top/left stack within the HF container at cursorY.
        fragEl.style.top = `${cursorY}px`;
        fragEl.style.left = '0';
        containerEl.appendChild(fragEl);
        cursorY += measure.totalHeight;
      }
    } else if (block.kind === 'image') {
      if (measure.kind !== 'image') continue;
      // Block-level images stack in the HF flow like paragraphs/tables.
      const syntheticFragment: ImageFragment = {
        kind: 'image',
        blockId: block.id,
        x: 0,
        y: cursorY,
        width: measure.width,
        height: measure.height,
        pmStart: block.pmStart,
        pmEnd: block.pmEnd,
      };
      const fragEl = renderImageFragment(
        syntheticFragment,
        block,
        measure,
        { ...context, positioning: 'absolute' },
        { document: doc }
      );
      fragEl.style.top = `${cursorY}px`;
      fragEl.style.left = '0';
      containerEl.appendChild(fragEl);
      cursorY += measure.height;
    } else if (block.kind === 'textBox') {
      if (measure.kind !== 'textBox') continue;
      // Text boxes stack in the HF flow. headerFooterLayout already reserves
      // their height; without this branch they were measured but never
      // painted, so they showed in the inline editor but not the page view.
      // Reuse the image path's anchor resolution: a footer date is commonly
      // anchored with a negative positionV so it paints above the footer band.
      const boxTop = resolveHeaderFooterFloatTop(
        { height: measure.height, paragraphY: cursorY, position: block.position ?? {} },
        layout
      );
      const syntheticFragment: TextBoxFragment = {
        kind: 'textBox',
        blockId: block.id,
        x: 0,
        y: boxTop,
        width: measure.width,
        height: measure.height,
        pmStart: block.pmStart,
        pmEnd: block.pmEnd,
        // Same stacking the body gives its floats — without it an anchored HF
        // box paints under body artwork it is meant to sit on.
        isFloating: block.displayMode === 'float',
        // Only a FLOATING box stacks. An in-flow or topAndBottom box given a
        // front z-index paints over the body and, being pointer-interactive,
        // swallows clicks meant for the document text.
        zIndex:
          block.displayMode !== 'float'
            ? undefined
            : block.wrapType === 'behind'
              ? -1
              : headerFooterFrontZIndex(block.relativeHeight ?? 1),
      };
      const fragEl = renderTextBoxFragment(
        syntheticFragment,
        block,
        measure,
        { ...context, positioning: 'absolute' },
        { document: doc }
      );
      fragEl.style.top = `${boxTop}px`;
      // Honor the anchor's horizontal position (e.g. centered relative to the
      // page) instead of pinning the box to the left.
      fragEl.style.left = resolveHeaderFooterFloatLeft(
        measure.width,
        block.position?.horizontal,
        layout
      );
      containerEl.appendChild(fragEl);
      // Floating text boxes (square/tight/through/behind/inFront wrap) are
      // positioned and do NOT advance the flow — surrounding header content
      // flows as if the box weren't there, mirroring floating tables above and
      // matching Word: a centered banner sits beside the left/right header text
      // instead of pushing it down. This also keeps a tall page-anchored
      // letterhead from shoving the in-flow content past the header band
      // (#705). Advancing for a float made the in-flow content overflow the
      // band (which excludes floats from `flowHeight`) and overlap the body.
      // Inline and topAndBottom boxes still stack on the cursor (they reserve
      // in-flow vertical space).
      if (block.displayMode !== 'float') {
        cursorY += measure.height;
      }
    } else if (
      block.kind === 'sectionBreak' ||
      block.kind === 'pageBreak' ||
      block.kind === 'columnBreak'
    ) {
      // Section/page/column breaks carry no rendering in the header/footer
      // flow — headers and footers reflow per page, so a break has no meaning.
    } else {
      // Exhaustiveness guard: every FlowBlock variant must be handled above.
      // A new variant fails the typecheck here instead of silently vanishing
      // from the header/footer page view.
      assertExhaustiveFlowBlock(block, 'renderHeaderFooterContent');
    }
  }

  // Render floating images with absolute positioning
  let lastBehindEl: HTMLElement | null = null;
  for (const floatImg of floatingImages) {
    const top = resolveHeaderFooterFloatTop(floatImg, layout);

    // Word does not render a header/footer drawing that resolves entirely
    // outside the page — e.g. a paragraph-anchored image carrying a large
    // vertical offset that pushes it past the page bottom (some templates
    // accumulate such stale off-page anchors). The HF container intentionally
    // does NOT clip (images may bleed past the narrow text band, see below), so
    // an off-page float would otherwise spill visibly onto the adjacent page.
    // Cull it here to match Word. Must stay in lockstep with the matching guard
    // in `calculateHeaderFooterVisualBounds` (layout-bridge) so paint and
    // measurement agree on which floats exist.
    const absoluteTop = layout.flowTop + top;
    if (absoluteTop >= layout.pageHeight || absoluteTop + floatImg.height <= 0) {
      continue;
    }

    const img = doc.createElement('img');
    setImageAssetSource(
      img,
      {
        assetId: floatImg.run.assetId,
        src: floatImg.src,
        width: floatImg.width,
        height: floatImg.height,
      },
      context.imageAssetLoader
    );
    img.width = floatImg.width;
    img.height = floatImg.height;
    if (floatImg.alt) img.alt = floatImg.alt;

    img.style.position = 'absolute';
    img.style.display = 'block';
    // Header/footer images can intentionally extend beyond the text area.
    // Override global img resets (for example max-width: 100%) so the DOCX
    // anchor extent is honored instead of shrinking to the header/footer box.
    img.style.width = `${floatImg.width}px`;
    img.style.height = `${floatImg.height}px`;
    img.style.maxWidth = 'none';
    img.style.maxHeight = 'none';

    // OOXML srcRect crop / picture opacity / rotation — the header's
    // G2-badge-strip-cropped-from-a-screenshot case renders as a squashed
    // full screenshot without this. Unlike body images (object-fit: cover,
    // an approximation that leaks a sliver of the source when the display
    // aspect drifts from the cropped region's), header floats aren't
    // selectable/resizable, so we can afford the EXACT crop: a scaled
    // inner img inside an overflow-hidden box.
    const run = floatImg.run;
    const clamp01 = (v: number) => Math.min(1, Math.max(0, v));
    const cropL = clamp01(run.cropLeft ?? 0);
    const cropR = clamp01(run.cropRight ?? 0);
    const cropT = clamp01(run.cropTop ?? 0);
    const cropB = clamp01(run.cropBottom ?? 0);
    // Defense in depth on top of the parser's [0, 1] clamp: a crop pair
    // that consumes (nearly) the whole source axis would explode the
    // inner-img scale factor — render uncropped instead.
    const denomW = 1 - cropL - cropR;
    const denomH = 1 - cropT - cropB;
    let el: HTMLElement = img;
    if ((cropL || cropR || cropT || cropB) && denomW >= 0.01 && denomH >= 0.01) {
      const box = doc.createElement('div');
      box.style.position = 'absolute';
      box.style.width = `${floatImg.width}px`;
      box.style.height = `${floatImg.height}px`;
      box.style.overflow = 'hidden';
      const innerW = floatImg.width / denomW;
      const innerH = floatImg.height / denomH;
      img.style.position = 'absolute';
      img.style.width = `${innerW}px`;
      img.style.height = `${innerH}px`;
      img.style.left = `${-cropL * innerW}px`;
      img.style.top = `${-cropT * innerH}px`;
      box.appendChild(img);
      el = box;
      if (run.opacity != null && run.opacity < 1) {
        img.style.opacity = String(Math.max(0, run.opacity));
      }
    } else if (hasImageVisualAttrs(run)) {
      applyImageVisualAttrs(img, run);
    }
    // `a:prstGeom` rounding belongs on whichever element is the picture's
    // visible box — the overflow-hidden crop wrapper when there is one, the
    // `<img>` otherwise (where `applyImageVisualAttrs` has already set it).
    if (el !== img) {
      const radius = imageCornerRadiusCss({
        ...run,
        width: floatImg.width,
        height: floatImg.height,
      });
      if (radius) el.style.borderRadius = radius;
    }
    if (run.transform) {
      img.style.transform = run.transform;
      img.style.transformOrigin = 'center center';
    }

    applyHeaderFooterFloatHorizontalPosition(el as HTMLImageElement, floatImg, layout);
    el.style.top = `${top}px`;
    if (floatImg.run.renderOnly) el.dataset.renderOnly = '1';

    // Same band as an anchored HF TEXT BOX: `behindDoc="0"` means in front of
    // every story's text, the footer's included. A COMET divider page is a
    // full-bleed header picture whose white right half is what hides the
    // footer in Word — painted in DOM order instead, our footer rule and page
    // number sat on top of the artwork. Ordering by `relativeHeight` inside
    // the band keeps HF floats stacked against each other as authored.
    if (floatImg.run.wrapType !== 'behind') {
      el.style.zIndex = String(headerFooterFrontZIndex(floatImg.run.relativeHeight ?? 1));
    }

    // `behindDoc` floats paint under the flow text, and floats are appended
    // after it, so DOM order is what puts them behind (z-index cannot). Chain
    // them off the last one so they keep document order among themselves.
    if (floatImg.run.wrapType === 'behind') {
      containerEl.insertBefore(
        el,
        lastBehindEl ? lastBehindEl.nextSibling : containerEl.firstChild
      );
      lastBehindEl = el;
    } else {
      containerEl.appendChild(el);
    }
  }

  return containerEl;
}

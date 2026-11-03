/**
 * Body exclusion bands from wrapping header/footer floats.
 *
 * Word lets a floating object in a header or footer push the main body text
 * out of its way on every page that uses that header/footer — the familiar
 * "the picture in my header pushes my text down" behaviour — whenever the
 * object has a text-wrapping mode (`square`, `tight`, `through`,
 * `topAndBottom`). `wrapNone` floats (in front of / behind text) displace
 * nothing.
 *
 * This module turns those floats into full-width {@link BodyExclusionBand}s
 * for the paginator (`LayoutOptions.pageBodyExclusions`). Only floats that
 * leave no usable width beside them inside the text column become bands — a
 * page-wide banner, or a `topAndBottom` object. A narrower float that text
 * could wrap beside does not narrow body lines yet (body lines are measured
 * before pagination decides which lines start a page), so it stays out of the
 * body's way only through z-order: header/footer content paints behind the
 * body.
 *
 * Positions come from the painter's own resolvers
 * (`resolveHeaderFooterFloatTop` / `resolveHeaderFooterFloatLeftPx`) and the
 * same flow walk and off-page culling as `renderHeaderFooterContent`, so the
 * band sits exactly where the float is painted.
 */

import type { ImageRun, Page, TextBoxBlock } from '../layout-engine/types';
import type { BodyExclusionBand, PageBodyExclusions } from '../layout-engine/bodyExclusions';
import { wrapsAroundText } from '../docx/wrapTypes';
import { imageWrapTextFromCssFloat } from '../layout-painter/floatingImageFlow';
import {
  resolveHeaderFooterFloatLeftPx,
  resolveHeaderFooterFloatTop,
  type HeaderFooterContent,
  type HeaderFooterLayoutInfo,
} from '../layout-painter/renderPage/headerFooter';
import { MIN_WRAP_SEGMENT_WIDTH } from './measuring/floatingZones';

/** Word's default `w:header` / `w:footer` distance (0.5in = 48px). */
const DEFAULT_HF_DISTANCE_PX = 48;

type PageGeometry = Readonly<Pick<Page, 'size' | 'margins'>>;

/** Which side(s) of a float text may use (OOXML `wp:wrapSquare@wrapText`). */
type WrapSide = 'bothSides' | 'left' | 'right' | 'largest';

/** A text-wrapping float of an HF story, resolved to page coordinates. */
interface PlacedFloat {
  x: number;
  y: number;
  width: number;
  height: number;
  distTop: number;
  distBottom: number;
  distLeft: number;
  distRight: number;
  wrapType: string;
  wrapSide: WrapSide;
}

/** Wrap modes that displace body text: text wraps around, or above and below. */
function displacesBody(wrapType: string | undefined): wrapType is string {
  return wrapsAroundText(wrapType) || wrapType === 'topAndBottom';
}

/**
 * True when text could not use either side of the float inside the body's
 * text column — the float then blocks the column's full width over its height.
 * Same `MIN_WRAP_SEGMENT_WIDTH` floor the body uses to decide a floating table
 * is block-like.
 */
function leavesNoRoomBeside(float: PlacedFloat, page: PageGeometry): boolean {
  if (float.wrapType === 'topAndBottom') return true;
  const columnLeft = page.margins.left;
  const columnRight = page.size.w - page.margins.right;
  const roomLeft = float.x - float.distLeft - columnLeft >= MIN_WRAP_SEGMENT_WIDTH;
  const roomRight =
    columnRight - (float.x + float.width + float.distRight) >= MIN_WRAP_SEGMENT_WIDTH;
  switch (float.wrapSide) {
    case 'left':
      return !roomLeft;
    case 'right':
      return !roomRight;
    default:
      return !roomLeft && !roomRight;
  }
}

/**
 * Walk an HF story the way `renderHeaderFooterContent` does and collect its
 * text-wrapping floats (anchored picture runs and floating text boxes) in page
 * coordinates.
 */
function placeWrappingFloats(content: HeaderFooterContent, layout: HeaderFooterLayoutInfo) {
  const floats: PlacedFloat[] = [];
  const place = (
    width: number,
    height: number,
    paragraphY: number,
    position: NonNullable<ImageRun['position']>,
    rest: Omit<PlacedFloat, 'x' | 'y' | 'width' | 'height'>
  ): void => {
    const top = resolveHeaderFooterFloatTop({ height, paragraphY, position }, layout);
    const y = layout.flowTop + top;
    // The painter culls floats lying wholly off the page; so do we.
    if (y >= layout.pageHeight || y + height <= 0) return;
    const left = resolveHeaderFooterFloatLeftPx(width, position.horizontal, layout) ?? 0;
    floats.push({ x: layout.flowLeft + left, y, width, height, ...rest });
  };

  // Mirrors the cursor of `renderHeaderFooterContent`: paragraphs, inline
  // tables, block images and non-floating text boxes advance it; floating
  // tables and floating text boxes do not.
  let cursorY = 0;
  for (let i = 0; i < content.blocks.length; i++) {
    const block = content.blocks[i];
    const measure = content.measures[i];
    if (!block || !measure) continue;

    if (block.kind === 'paragraph') {
      if (measure.kind !== 'paragraph') continue;
      for (const run of block.runs) {
        if (run.kind !== 'image' || !run.position) continue;
        if (!displacesBody(run.wrapType)) continue;
        place(run.width, run.height, cursorY, run.position, {
          distTop: run.distTop ?? 0,
          distBottom: run.distBottom ?? 0,
          distLeft: run.distLeft ?? 12,
          distRight: run.distRight ?? 12,
          wrapType: run.wrapType,
          wrapSide: imageWrapTextFromCssFloat(run.cssFloat),
        });
      }
      cursorY += measure.totalHeight;
    } else if (block.kind === 'table') {
      if (measure.kind !== 'table') continue;
      if (!block.floating) cursorY += measure.totalHeight;
    } else if (block.kind === 'image') {
      if (measure.kind !== 'image') continue;
      cursorY += measure.height;
    } else if (block.kind === 'textBox') {
      if (measure.kind !== 'textBox') continue;
      const box: TextBoxBlock = block;
      if (displacesBody(box.wrapType)) {
        place(measure.width, measure.height, cursorY, box.position ?? {}, {
          distTop: box.distTop ?? 0,
          distBottom: box.distBottom ?? 0,
          distLeft: box.distLeft ?? 12,
          distRight: box.distRight ?? 12,
          wrapType: box.wrapType,
          wrapSide: box.wrapText ?? 'bothSides',
        });
      }
      if (box.displayMode !== 'float') cursorY += measure.height;
    }
  }
  return floats;
}

/**
 * Full-width body exclusion bands one header or footer creates on a page with
 * the given geometry. `distance` is the `w:header` / `w:footer` distance the
 * painter uses (defaults to the page margins' value, then 0.5in).
 *
 * @public
 */
export function headerFooterBodyExclusions(
  content: HeaderFooterContent | undefined,
  section: 'header' | 'footer',
  page: PageGeometry,
  distance?: number
): BodyExclusionBand[] {
  if (!content || content.blocks.length === 0) return [];
  const { size, margins } = page;
  // Flow origin exactly as renderPage places it: the header flow starts at the
  // header distance; the footer band's BOTTOM sits at the footer distance and
  // its in-flow content grows upward.
  const flowTop =
    section === 'header'
      ? (distance ?? margins.header ?? DEFAULT_HF_DISTANCE_PX)
      : size.h -
        (distance ?? margins.footer ?? DEFAULT_HF_DISTANCE_PX) -
        (content.flowHeight ?? content.height ?? 0);
  const layout: HeaderFooterLayoutInfo = {
    flowTop,
    flowLeft: margins.left,
    contentWidth: size.w - margins.left - margins.right,
    pageWidth: size.w,
    pageHeight: size.h,
    margins,
  };

  const bands: BodyExclusionBand[] = [];
  for (const float of placeWrappingFloats(content, layout)) {
    if (!leavesNoRoomBeside(float, page)) continue;
    bands.push({
      top: float.y - float.distTop,
      bottom: float.y + float.height + float.distBottom,
    });
  }
  return bands;
}

/** Header/footer stories in play for a layout, as the painter receives them. */
export interface HeaderFooterExclusionSources {
  header?: HeaderFooterContent;
  footer?: HeaderFooterContent;
  /** Used on page 1 instead of `header` when `titlePg` is set. */
  firstPageHeader?: HeaderFooterContent;
  /** Used on page 1 instead of `footer` when `titlePg` is set. */
  firstPageFooter?: HeaderFooterContent;
  titlePg?: boolean;
  /** `w:header` distance in px, when the section states one. */
  headerDistance?: number;
  /** `w:footer` distance in px, when the section states one. */
  footerDistance?: number;
}

/**
 * Build the `LayoutOptions.pageBodyExclusions` callback for a layout: each
 * page gets the bands of the header and footer the painter will draw on it
 * (the first-page pair on page 1 under `titlePg`, the default pair elsewhere).
 * Returns `undefined` when no header/footer float displaces the body, so
 * pagination takes its usual path.
 *
 * @public
 */
export function createHeaderFooterBodyExclusions(
  sources: HeaderFooterExclusionSources
): PageBodyExclusions | undefined {
  const stories = [
    sources.header,
    sources.footer,
    sources.titlePg ? sources.firstPageHeader : undefined,
    sources.titlePg ? sources.firstPageFooter : undefined,
  ];
  const mayDisplace = stories.some((story) =>
    story?.blocks.some(
      (block) =>
        (block.kind === 'paragraph' &&
          block.runs.some(
            (run) => run.kind === 'image' && !!run.position && displacesBody(run.wrapType)
          )) ||
        (block.kind === 'textBox' && displacesBody(block.wrapType))
    )
  );
  if (!mayDisplace) return undefined;

  // Pages share a handful of geometries; resolve each (story, geometry) once.
  const cache = new Map<string, BodyExclusionBand[]>();
  return (page) => {
    const firstPage = sources.titlePg === true && page.number === 1;
    const header = firstPage ? sources.firstPageHeader : sources.header;
    const footer = firstPage ? sources.firstPageFooter : sources.footer;
    const { size, margins } = page;
    const key = [
      firstPage ? 'first' : 'default',
      size.w,
      size.h,
      margins.top,
      margins.right,
      margins.bottom,
      margins.left,
      margins.header,
      margins.footer,
    ].join('|');
    let bands = cache.get(key);
    if (!bands) {
      bands = [
        ...headerFooterBodyExclusions(header, 'header', page, sources.headerDistance),
        ...headerFooterBodyExclusions(footer, 'footer', page, sources.footerDistance),
      ];
      cache.set(key, bands);
    }
    return bands;
  };
}

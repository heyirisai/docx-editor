/**
 * `w:pgBorders` — the rule Word draws around the page, painted as an absolutely
 * positioned overlay so it does not disturb the content flow.
 */

import type { Page } from '../../layout-engine/types';
import type { BorderSpec, Theme } from '../../types/document';

import { PAGE_OVERLAY_Z } from '../../layout-engine/zOrder';
import { borderToStyle } from '../../utils/formatToStyle';
import { pointsToPixels } from '../../utils/units';
import type { RenderPageOptions } from '../renderPage';

function pageBorderShouldRender(
  pageNumber: number,
  display?: 'allPages' | 'firstPage' | 'notFirstPage'
): boolean {
  switch (display ?? 'allPages') {
    case 'firstPage':
      return pageNumber === 1;
    case 'notFirstPage':
      return pageNumber !== 1;
    case 'allPages':
    default:
      return true;
  }
}

function pageBorderSpacePx(border: BorderSpec | undefined): number {
  return border?.space !== undefined ? pointsToPixels(border.space) : 0;
}

function applyPageBorderSide(
  element: HTMLElement,
  border: BorderSpec | undefined,
  side: 'Top' | 'Bottom' | 'Left' | 'Right',
  theme?: Theme | null
): void {
  if (!border || border.style === 'none' || border.style === 'nil') return;

  const styles = borderToStyle(border, side, theme);
  for (const [key, value] of Object.entries(styles)) {
    (element.style as unknown as Record<string, string>)[key] = String(value);
  }

  const styleKey = `border${side}Style`;
  const widthKey = `border${side}Width`;
  const styleValue = (element.style as unknown as Record<string, string>)[styleKey];
  if (styleValue === 'double') {
    const widthValue = parseFloat((element.style as unknown as Record<string, string>)[widthKey]);
    if (!Number.isFinite(widthValue) || widthValue < 3) {
      (element.style as unknown as Record<string, string>)[widthKey] = '3px';
    }
  }
}

export function renderPageBorderOverlay(
  page: Page,
  options: RenderPageOptions,
  doc: Document
): HTMLElement | null {
  const pb = options.pageBorders;
  if (!pb || !pageBorderShouldRender(page.number, pb.display)) return null;

  const hasBorder = [pb.top, pb.bottom, pb.left, pb.right].some(
    (border) => border && border.style !== 'none' && border.style !== 'nil'
  );
  if (!hasBorder) return null;

  const offsetFrom = pb.offsetFrom ?? 'text';
  const topOffset = pageBorderSpacePx(pb.top);
  const rightOffset = pageBorderSpacePx(pb.right);
  const bottomOffset = pageBorderSpacePx(pb.bottom);
  const leftOffset = pageBorderSpacePx(pb.left);

  const overlay = doc.createElement('div');
  overlay.className = 'layout-page-border';
  overlay.style.position = 'absolute';
  overlay.style.pointerEvents = 'none';
  overlay.style.boxSizing = 'border-box';
  overlay.style.zIndex = pb.zOrder === 'back' ? '0' : String(PAGE_OVERLAY_Z);

  if (offsetFrom === 'page') {
    overlay.style.top = `${topOffset}px`;
    overlay.style.right = `${rightOffset}px`;
    overlay.style.bottom = `${bottomOffset}px`;
    overlay.style.left = `${leftOffset}px`;
  } else {
    overlay.style.top = `${Math.max(0, page.margins.top - topOffset)}px`;
    overlay.style.right = `${Math.max(0, page.margins.right - rightOffset)}px`;
    overlay.style.bottom = `${Math.max(0, page.margins.bottom - bottomOffset)}px`;
    overlay.style.left = `${Math.max(0, page.margins.left - leftOffset)}px`;
  }

  applyPageBorderSide(overlay, pb.top, 'Top', options.theme);
  applyPageBorderSide(overlay, pb.bottom, 'Bottom', options.theme);
  applyPageBorderSide(overlay, pb.left, 'Left', options.theme);
  applyPageBorderSide(overlay, pb.right, 'Right', options.theme);

  return overlay;
}

/**
 * Coordinate mapping for grouped drawings (`wpg:wgp`).
 *
 * A group declares its own child coordinate space (`a:chOff` / `a:chExt`) and
 * the page-space box it is drawn into (`a:off` / `a:ext`). Every descendant's
 * `a:xfrm` is expressed in that child space, and a nested `wpg:grpSp`
 * introduces another one. Both consumers of a group — the picture preview in
 * `groupPreview.ts` and the shape/text-box preview in `textBoxParser.ts` —
 * need the same mapping, so it lives here rather than in either of them.
 */

import { getChildElements, getLocalName, getAttribute, type XmlElement } from './xmlParser';

/**
 * Depth bound for the walks below. The markup is attacker-controlled (a `.docx`
 * is a zip of XML), and groups nest arbitrarily, so an unbounded walk overflows
 * the stack on a hostile file. Real drawings nest a handful of levels.
 */
export const MAX_GROUP_DEPTH = 64;

export function findDeepByLocalName(
  root: XmlElement,
  localName: string,
  depth = 0
): XmlElement | null {
  if (depth >= MAX_GROUP_DEPTH) return null;
  for (const child of getChildElements(root)) {
    if (getLocalName(child.name ?? '') === localName) return child;
    const nested = findDeepByLocalName(child, localName, depth + 1);
    if (nested) return nested;
  }
  return null;
}

export function numOrNull(value: string | null | undefined): number | null {
  if (value === null || value === undefined || value === '') return null;
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}

/** `a:off` / `a:ext` / `a:chOff` / `a:chExt` of an `a:xfrm`, in EMU. */
export function readXfrm(xfrm: XmlElement | null) {
  if (!xfrm) return null;
  const read = (name: string, xAttr: string, yAttr: string) => {
    const el = getChildElements(xfrm).find((c) => getLocalName(c.name ?? '') === name);
    if (!el) return null;
    const x = numOrNull(getAttribute(el, null, xAttr));
    const y = numOrNull(getAttribute(el, null, yAttr));
    return x === null || y === null ? null : { x, y };
  };
  return {
    off: read('off', 'x', 'y'),
    ext: read('ext', 'cx', 'cy'),
    chOff: read('chOff', 'x', 'y'),
    chExt: read('chExt', 'cx', 'cy'),
  };
}

/** Maps a coordinate in some group's child space onto the page, in EMU. */
export interface GroupFrame {
  scaleX: number;
  scaleY: number;
  chOffX: number;
  chOffY: number;
  baseX: number;
  baseY: number;
}

export const mapX = (f: GroupFrame, x: number): number => f.baseX + (x - f.chOffX) * f.scaleX;
export const mapY = (f: GroupFrame, y: number): number => f.baseY + (y - f.chOffY) * f.scaleY;

/** The outermost frame of a `wpg:wgp`, or null when it declares no geometry. */
export function rootGroupFrame(group: XmlElement): GroupFrame | null {
  const grpSpPr = getChildElements(group).find((c) => getLocalName(c.name ?? '') === 'grpSpPr');
  const xfrm = readXfrm(grpSpPr ? findDeepByLocalName(grpSpPr, 'xfrm') : null);
  if (!xfrm?.ext || !xfrm.chExt) return null;
  const chOff = xfrm.chOff ?? { x: 0, y: 0 };
  return {
    scaleX: xfrm.chExt.x === 0 ? 1 : xfrm.ext.x / xfrm.chExt.x,
    scaleY: xfrm.chExt.y === 0 ? 1 : xfrm.ext.y / xfrm.chExt.y,
    chOffX: chOff.x,
    chOffY: chOff.y,
    baseX: 0,
    baseY: 0,
  };
}

/**
 * Visit every descendant with local name `localName`, paired with the frame of
 * the group that actually contains it. A nested `wpg:grpSp` introduces its own
 * child coordinate space, so applying the outer group's scale to its children
 * puts them in the wrong place and at the wrong size.
 */
export function collectGroupLeaves(
  container: XmlElement,
  frame: GroupFrame,
  localName: string,
  out: Array<{ element: XmlElement; frame: GroupFrame }>,
  depth = 0
): void {
  if (depth >= MAX_GROUP_DEPTH) return;
  for (const child of getChildElements(container)) {
    const name = getLocalName(child.name ?? '');
    if (name === localName) {
      out.push({ element: child, frame });
      continue;
    }
    if (name === 'grpSp') {
      const xfrm = readXfrm(findDeepByLocalName(child, 'xfrm'));
      let inner = frame;
      if (xfrm?.off && xfrm.ext) {
        const chExt = xfrm.chExt ?? xfrm.ext;
        const chOff = xfrm.chOff ?? { x: 0, y: 0 };
        inner = {
          scaleX: (chExt.x === 0 ? 1 : xfrm.ext.x / chExt.x) * frame.scaleX,
          scaleY: (chExt.y === 0 ? 1 : xfrm.ext.y / chExt.y) * frame.scaleY,
          chOffX: chOff.x,
          chOffY: chOff.y,
          baseX: mapX(frame, xfrm.off.x),
          baseY: mapY(frame, xfrm.off.y),
        };
      }
      collectGroupLeaves(child, inner, localName, out, depth + 1);
      continue;
    }
    collectGroupLeaves(child, frame, localName, out, depth + 1);
  }
}

/**
 * `wp:positionH` / `wp:positionV` off a `wp:anchor`. An anchor carries EITHER
 * `wp:posOffset` OR `wp:align`; treating a missing offset as 0 pins a centred
 * or right-aligned group to the top-left.
 */
export function readGroupAnchorPosition(anchor: XmlElement, axis: 'positionH' | 'positionV') {
  const pos = getChildElements(anchor).find((c) => getLocalName(c.name ?? '') === axis);
  if (!pos) return null;
  const relativeTo = getAttribute(pos, null, 'relativeFrom') ?? undefined;
  const offsetEl = getChildElements(pos).find((c) => getLocalName(c.name ?? '') === 'posOffset');
  const offset = offsetEl ? numOrNull((offsetEl.elements?.[0]?.text as string) ?? null) : null;
  const alignEl = getChildElements(pos).find((c) => getLocalName(c.name ?? '') === 'align');
  const alignment = (alignEl?.elements?.[0]?.text as string | undefined)?.trim() || undefined;
  return { relativeTo, offset: offset ?? 0, alignment, hasOffset: offset !== null };
}

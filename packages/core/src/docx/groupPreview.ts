/**
 * Paint-only view of a grouped drawing (`wpg:wgp`). The group is preserved
 * verbatim as `rawXml`; these derived `renderOnly` images just let it paint.
 */

import type { DrawingContent } from '../types/content';
import type { Image } from '../types/content/image';
import type { MediaFile, RelationshipMap } from '../types/document';
import { resolveImageData } from './imageParser';
import { getChildElements, getLocalName, getAttribute, type XmlElement } from './xmlParser';

/**
 * Depth bound for the walks below. The markup is attacker-controlled (a `.docx`
 * is a zip of XML), and groups nest arbitrarily, so an unbounded walk overflows
 * the stack on a hostile file. Real drawings nest a handful of levels.
 */
const MAX_DEPTH = 64;

function findDeepByLocalName(root: XmlElement, localName: string, depth = 0): XmlElement | null {
  if (depth >= MAX_DEPTH) return null;
  for (const child of getChildElements(root)) {
    if (getLocalName(child.name ?? '') === localName) return child;
    const nested = findDeepByLocalName(child, localName, depth + 1);
    if (nested) return nested;
  }
  return null;
}

function findAllDeepByLocalName(root: XmlElement, localName: string, depth = 0): XmlElement[] {
  if (depth >= MAX_DEPTH) return [];
  const out: XmlElement[] = [];
  for (const child of getChildElements(root)) {
    if (getLocalName(child.name ?? '') === localName) out.push(child);
    out.push(...findAllDeepByLocalName(child, localName, depth + 1));
  }
  return out;
}

/**
 * `r:embed` off an `a:blip`. The relationships namespace is conventionally
 * bound to `r`, but nothing requires it, so fall back to matching on the
 * attribute's local name rather than assuming the prefix.
 */
function blipRelationshipId(blip: XmlElement | null): string | null {
  const direct = getAttribute(blip, 'r', 'embed');
  if (direct) return direct;
  for (const [key, value] of Object.entries(blip?.attributes ?? {})) {
    if (getLocalName(key) === 'embed') return String(value);
  }
  return null;
}

function num(value: string | null | undefined): number | null {
  if (value === null || value === undefined || value === '') return null;
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}

/** `a:off` / `a:ext` pair from an `a:xfrm`, in EMU. */
function readXfrm(xfrm: XmlElement | null) {
  if (!xfrm) return null;
  const read = (name: string, xAttr: string, yAttr: string) => {
    const el = getChildElements(xfrm).find((c) => getLocalName(c.name ?? '') === name);
    if (!el) return null;
    const x = num(getAttribute(el, null, xAttr));
    const y = num(getAttribute(el, null, yAttr));
    return x === null || y === null ? null : { x, y };
  };
  return {
    off: read('off', 'x', 'y'),
    ext: read('ext', 'cx', 'cy'),
    chOff: read('chOff', 'x', 'y'),
    chExt: read('chExt', 'cx', 'cy'),
  };
}

function readAnchorPosition(anchor: XmlElement, axis: 'positionH' | 'positionV') {
  const pos = getChildElements(anchor).find((c) => getLocalName(c.name ?? '') === axis);
  if (!pos) return null;
  const relativeTo = getAttribute(pos, null, 'relativeFrom') ?? undefined;
  const offsetEl = getChildElements(pos).find((c) => getLocalName(c.name ?? '') === 'posOffset');
  const offset = offsetEl ? num((offsetEl.elements?.[0]?.text as string) ?? null) : null;
  // An anchor carries EITHER `wp:posOffset` OR `wp:align`. Treating a missing
  // offset as 0 pins a centred or right-aligned group to the top-left.
  const alignEl = getChildElements(pos).find((c) => getLocalName(c.name ?? '') === 'align');
  const alignment = (alignEl?.elements?.[0]?.text as string | undefined)?.trim() || undefined;
  return { relativeTo, offset: offset ?? 0, alignment, hasOffset: offset !== null };
}

/**
 * Anchored images for each picture inside a grouped drawing, or `[]` when the
 * element holds no group (a lone shape has nothing to show).
 */
export function deriveGroupPreviewImages(
  alternateContentEl: XmlElement,
  rels: RelationshipMap | null | undefined,
  media: Map<string, MediaFile> | null | undefined
): DrawingContent[] {
  const anchor = findDeepByLocalName(alternateContentEl, 'anchor');
  if (!anchor) return [];
  const group = findDeepByLocalName(anchor, 'wgp');
  if (!group) return [];

  const grpSpPr = getChildElements(group).find((c) => getLocalName(c.name ?? '') === 'grpSpPr');
  const groupXfrm = readXfrm(grpSpPr ? findDeepByLocalName(grpSpPr, 'xfrm') : null);
  if (!groupXfrm?.ext || !groupXfrm.chExt) return [];

  // Child coordinates live in the group's own space; map them onto the page.
  const scaleX = groupXfrm.chExt.x === 0 ? 1 : groupXfrm.ext.x / groupXfrm.chExt.x;
  const scaleY = groupXfrm.chExt.y === 0 ? 1 : groupXfrm.ext.y / groupXfrm.chExt.y;
  const chOff = groupXfrm.chOff ?? { x: 0, y: 0 };

  const posH = readAnchorPosition(anchor, 'positionH');
  const posV = readAnchorPosition(anchor, 'positionV');
  const behindDoc = getAttribute(anchor, null, 'behindDoc') === '1';
  const relativeHeight = num(getAttribute(anchor, null, 'relativeHeight')) ?? undefined;

  const out: DrawingContent[] = [];
  for (const pic of findAllDeepByLocalName(group, 'pic')) {
    const rId = blipRelationshipId(findDeepByLocalName(pic, 'blip'));
    if (!rId) continue;
    const resolved = resolveImageData(rId, rels ?? undefined, media ?? undefined);
    // `externalMedia` mode resolves to an assetId with no inline src — the
    // painter loads those lazily, so only a picture with neither is unpaintable.
    if (!resolved.src && !resolved.assetId) continue;

    const childXfrm = readXfrm(findDeepByLocalName(pic, 'xfrm'));
    if (!childXfrm?.off || !childXfrm.ext) continue;

    const image: Image = {
      type: 'image',
      rId,
      src: resolved.src,
      assetId: resolved.assetId,
      mimeType: resolved.mimeType,
      filename: resolved.filename,
      size: {
        width: Math.round(childXfrm.ext.x * scaleX),
        height: Math.round(childXfrm.ext.y * scaleY),
      },
      wrap: { type: behindDoc ? 'behind' : 'inFront' },
      position: {
        horizontal: {
          relativeTo: (posH?.relativeTo ?? 'column') as NonNullable<
            Image['position']
          >['horizontal']['relativeTo'],
          // An aligned group carries `alignment` and NO `posOffset`: the
          // painters resolve alignment only when no offset is present, so
          // emitting both would silently pin the group back to the origin.
          // The child delta is dropped with it — a lone picture (the common
          // case, a centred logo) then lands correctly, and the pictures of a
          // multi-picture aligned group share the group's anchor instead of
          // every one of them collapsing to the left edge.
          ...(posH?.hasOffset === false && posH.alignment
            ? {
                alignment: posH.alignment as NonNullable<
                  Image['position']
                >['horizontal']['alignment'],
              }
            : {
                posOffset: Math.round((posH?.offset ?? 0) + (childXfrm.off.x - chOff.x) * scaleX),
              }),
        },
        vertical: {
          relativeTo: (posV?.relativeTo ?? 'paragraph') as NonNullable<
            Image['position']
          >['vertical']['relativeTo'],
          ...(posV?.hasOffset === false && posV.alignment
            ? {
                alignment: posV.alignment as NonNullable<
                  Image['position']
                >['vertical']['alignment'],
              }
            : {
                posOffset: Math.round((posV?.offset ?? 0) + (childXfrm.off.y - chOff.y) * scaleY),
              }),
        },
      },
      relativeHeight,
      renderOnly: true,
    };
    out.push({ type: 'drawing', image });
  }
  return out;
}

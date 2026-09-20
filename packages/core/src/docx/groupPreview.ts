/**
 * Paint-only view of a grouped drawing (`wpg:wgp`). The group is preserved
 * verbatim as `rawXml`; these derived `renderOnly` images just let it paint.
 */

import type { DrawingContent } from '../types/content';
import type { Image } from '../types/content/image';
import type { MediaFile, RelationshipMap } from '../types/document';
import { resolveImageData } from './imageParser';
import { getLocalName, getAttribute, type XmlElement } from './xmlParser';
import {
  collectGroupLeaves,
  findDeepByLocalName,
  mapX,
  mapY,
  numOrNull as num,
  readGroupAnchorPosition as readAnchorPosition,
  readXfrm,
  rootGroupFrame,
  type GroupFrame,
} from './groupFrame';

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

  // Child coordinates live in the group's own space; map them onto the page.
  const rootFrame: GroupFrame | null = rootGroupFrame(group);
  if (!rootFrame) return [];

  const posH = readAnchorPosition(anchor, 'positionH');
  const posV = readAnchorPosition(anchor, 'positionV');
  const behindDoc = getAttribute(anchor, null, 'behindDoc') === '1';
  const relativeHeight = num(getAttribute(anchor, null, 'relativeHeight')) ?? undefined;

  const out: DrawingContent[] = [];
  const pictures: Array<{ element: XmlElement; frame: GroupFrame }> = [];
  collectGroupLeaves(group, rootFrame, 'pic', pictures);
  for (const { element: pic, frame } of pictures) {
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
        width: Math.round(childXfrm.ext.x * frame.scaleX),
        height: Math.round(childXfrm.ext.y * frame.scaleY),
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
                posOffset: Math.round((posH?.offset ?? 0) + mapX(frame, childXfrm.off.x)),
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
                posOffset: Math.round((posV?.offset ?? 0) + mapY(frame, childXfrm.off.y)),
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

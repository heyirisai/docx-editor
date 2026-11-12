/**
 * Paint-only previews for drawings the document model does not represent:
 * DrawingML shapes without text (`wps:wsp`), groups (`wpg:wgp`), charts,
 * diagrams, canvases and legacy VML shapes.
 *
 * The source markup is preserved verbatim as `rawXml` and remains the only
 * thing written back on save. What this module adds is a `renderOnly` image
 * that carries the drawing's real frame — its extent, position, alignment,
 * wrapping and z-order, read from `wp:anchor` / `wp:inline` exactly as for a
 * picture — so the page lays out around it the way Word does. Its pixels are
 * an SVG drawn from the shape's geometry, fill and outline where those are
 * supported. Anything that is not (an unknown preset, a text box inside a
 * group, a chart) is drawn as a labelled, dashed placeholder occupying its
 * exact box, rather than dropped or guessed at.
 *
 * Every value here comes from an untrusted file: numbers are formatted, colours
 * are validated hex, embedded pictures must be `data:image/*` URLs, labels are
 * escaped, and walks are depth- and count-bounded. The SVG is loaded through an
 * `<img>`, so it cannot run script or fetch anything.
 */

import type { DrawingContent } from '../../types/content';
import type { Image, ImageTransform } from '../../types/content/image';
import type { MediaFile, RelationshipMap, Theme } from '../../types/document';
import { parseDrawing, resolveImageData } from '../imageParser';
import { deriveGroupPreviewImages } from '../groupPreview';
import { isWatermarkShape, parseStyleAttr } from '../vmlWatermarkParser';
import { getAttribute, getChildElements, getLocalName, type XmlElement } from '../xmlParser';
import { geometryPaths } from './geometry';
import { resolveDrawingColor, type ResolvedColor } from './color';

/** Labels painted on (and used as the alt text of) placeholders. */
export const UNSUPPORTED_DRAWING_LABELS = {
  shape: 'Unsupported shape',
  group: 'Unsupported shape group',
  textbox: 'Unsupported text box',
  picture: 'Unsupported picture',
  chart: 'Unsupported chart',
  diagram: 'Unsupported diagram',
  canvas: 'Unsupported drawing canvas',
  graphic: 'Unsupported graphic',
} as const;

export type UnsupportedDrawingKind = keyof typeof UNSUPPORTED_DRAWING_LABELS;

export interface DrawingPreviewContext {
  rels: RelationshipMap | null | undefined;
  media: Map<string, MediaFile> | null | undefined;
  theme: Theme | null | undefined;
}

const EMU_PER_PX = 9525;
const MAX_DEPTH = 32;
const MAX_ITEMS = 4000;

const GRAPHIC_URI_KIND: ReadonlyMap<
  string,
  'picture' | 'shape' | 'group' | UnsupportedDrawingKind
> = new Map([
  ['http://schemas.openxmlformats.org/drawingml/2006/picture', 'picture'],
  ['http://schemas.microsoft.com/office/word/2010/wordprocessingShape', 'shape'],
  ['http://schemas.microsoft.com/office/word/2010/wordprocessingGroup', 'group'],
  ['http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas', 'canvas'],
  ['http://schemas.openxmlformats.org/drawingml/2006/chart', 'chart'],
  ['http://schemas.microsoft.com/office/drawing/2014/chartex', 'chart'],
  ['http://schemas.openxmlformats.org/drawingml/2006/diagram', 'diagram'],
]);

// ============================================================================
// XML HELPERS
// ============================================================================

const local = (el: XmlElement | null | undefined): string => getLocalName(el?.name ?? '');
const child = (el: XmlElement | null | undefined, name: string): XmlElement | null =>
  getChildElements(el).find((c) => local(c) === name) ?? null;

function num(raw: string | null | undefined): number | null {
  if (raw === null || raw === undefined || raw.trim() === '') return null;
  const n = Number(raw);
  return Number.isFinite(n) ? n : null;
}

const fmt = (n: number): string => {
  if (!Number.isFinite(n)) return '0';
  const r = Math.round(n * 100) / 100;
  return Object.is(r, -0) ? '0' : String(r);
};

function escapeXmlText(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** The container (`wp:anchor` / `wp:inline`) and its `a:graphicData`. */
function drawingParts(drawingEl: XmlElement) {
  const container =
    getChildElements(drawingEl).find((c) => local(c) === 'anchor' || local(c) === 'inline') ?? null;
  const graphicData = child(child(container, 'graphic'), 'graphicData');
  return { container, graphicData };
}

/** What a `w:drawing` holds, from its `a:graphicData@uri`. */
export function drawingGraphicKind(
  drawingEl: XmlElement
): 'picture' | 'shape' | 'group' | UnsupportedDrawingKind | null {
  const { container, graphicData } = drawingParts(drawingEl);
  if (!container) return null;
  if (!graphicData) return 'graphic';
  const uri = getAttribute(graphicData, null, 'uri') ?? '';
  const kind = GRAPHIC_URI_KIND.get(uri);
  if (kind) return kind;
  // Some producers omit or mangle the uri; trust the content instead.
  const first = local(getChildElements(graphicData)[0]);
  if (first === 'pic') return 'picture';
  if (first === 'wsp') return 'shape';
  if (first === 'wgp') return 'group';
  return 'graphic';
}

// ============================================================================
// GEOMETRY / TRANSFORM
// ============================================================================

interface Box {
  x: number;
  y: number;
  w: number;
  h: number;
}

/** Child space → root space: X = tx + x * sx. */
interface Frame {
  sx: number;
  sy: number;
  tx: number;
  ty: number;
}

interface Xfrm {
  off: { x: number; y: number } | null;
  ext: { x: number; y: number } | null;
  chOff: { x: number; y: number } | null;
  chExt: { x: number; y: number } | null;
  rot: number;
  flipH: boolean;
  flipV: boolean;
}

function readXfrm(xfrm: XmlElement | null): Xfrm | null {
  if (!xfrm) return null;
  const pair = (name: string, a: string, b: string) => {
    const el = child(xfrm, name);
    const x = num(getAttribute(el, null, a));
    const y = num(getAttribute(el, null, b));
    return x === null || y === null ? null : { x, y };
  };
  return {
    off: pair('off', 'x', 'y'),
    ext: pair('ext', 'cx', 'cy'),
    chOff: pair('chOff', 'x', 'y'),
    chExt: pair('chExt', 'cx', 'cy'),
    rot: (num(getAttribute(xfrm, null, 'rot')) ?? 0) / 60000,
    flipH:
      getAttribute(xfrm, null, 'flipH') === '1' || getAttribute(xfrm, null, 'flipH') === 'true',
    flipV:
      getAttribute(xfrm, null, 'flipV') === '1' || getAttribute(xfrm, null, 'flipV') === 'true',
  };
}

function mapBox(f: Frame, x: Xfrm): Box | null {
  if (!x.off || !x.ext) return null;
  return {
    x: f.tx + x.off.x * f.sx,
    y: f.ty + x.off.y * f.sy,
    w: x.ext.x * f.sx,
    h: x.ext.y * f.sy,
  };
}

/** Flip, then rotate, about the box centre (DrawingML order). */
function boxTransform(b: Box, rot: number, flipH: boolean, flipV: boolean): string {
  const cx = fmt(b.x + b.w / 2);
  const cy = fmt(b.y + b.h / 2);
  const parts: string[] = [];
  if (rot) parts.push(`rotate(${fmt(rot)} ${cx} ${cy})`);
  if (flipH || flipV) {
    parts.push(
      `translate(${cx} ${cy}) scale(${flipH ? -1 : 1} ${flipV ? -1 : 1}) translate(-${cx} -${cy})`
    );
  }
  return parts.join(' ');
}

// ============================================================================
// SVG ASSEMBLY
// ============================================================================

interface Painter {
  ctx: DrawingPreviewContext;
  defs: string[];
  body: string[];
  /** Placeholder labels, in CSS-pixel coordinates (see `placeholder`). */
  labels: string[];
  nextId: number;
  items: number;
  /** EMU in root space per CSS pixel of the painted image. */
  emuPerPx: number;
  /** Labels of the parts drawn as placeholders. */
  unsupported: Set<string>;
  /** Pictures that resolved only to an external asset (not embeddable). */
  externalPictures: boolean;
  /** Something real was drawn (not only placeholders). */
  drew: boolean;
}

type Paint =
  | { kind: 'none' }
  | { kind: 'color'; color: ResolvedColor }
  | { kind: 'ref'; id: string }
  | { kind: 'picture'; href: string };

function paintAttrs(p: Paint, prefix: 'fill' | 'stroke'): string {
  switch (p.kind) {
    case 'none':
      return `${prefix}="none"`;
    case 'color':
      return p.color.alpha < 1
        ? `${prefix}="${p.color.hex}" ${prefix}-opacity="${fmt(p.color.alpha)}"`
        : `${prefix}="${p.color.hex}"`;
    case 'ref':
      return `${prefix}="url(#${p.id})"`;
    case 'picture':
      return `${prefix}="none"`;
  }
}

/** A `data:image/*` URL the browser can decode inside an SVG, or null. */
function embeddablePicture(src: string | undefined, mimeType: string | undefined): string | null {
  if (
    !src ||
    !/^data:image\/(png|jpe?g|gif|webp|bmp|svg\+xml);base64,[A-Za-z0-9+/=\s]*$/i.test(src)
  ) {
    return null;
  }
  if (mimeType && /(emf|wmf|tiff)/i.test(mimeType)) return null;
  return src.replace(/\s+/g, '');
}

function blipRelId(blip: XmlElement | null): string | null {
  for (const [key, value] of Object.entries(blip?.attributes ?? {})) {
    if (getLocalName(key) === 'embed') return String(value);
  }
  return null;
}

/** `a:gradFill` → an SVG gradient in `defs`. */
function gradientPaint(p: Painter, grad: XmlElement, placeholder?: ResolvedColor): Paint {
  const stops: string[] = [];
  const gs = getChildElements(child(grad, 'gsLst'))
    .filter((g) => local(g) === 'gs')
    .map((g) => ({
      pos: num(getAttribute(g, null, 'pos')) ?? 0,
      color: resolveDrawingColor(g, p.ctx.theme, placeholder),
    }))
    .filter((s): s is { pos: number; color: ResolvedColor } => s.color !== null)
    .sort((a, b) => a.pos - b.pos)
    .slice(0, 64);
  if (gs.length === 0) return { kind: 'none' };
  if (gs.length === 1) return { kind: 'color', color: gs[0].color };
  for (const s of gs) {
    const op = s.color.alpha < 1 ? ` stop-opacity="${fmt(s.color.alpha)}"` : '';
    stops.push(
      `<stop offset="${fmt(Math.max(0, Math.min(100, s.pos / 1000)))}%" stop-color="${s.color.hex}"${op}/>`
    );
  }
  const id = `g${p.nextId++}`;
  const path = child(grad, 'path');
  if (path) {
    const ftr = child(path, 'fillToRect');
    const l = (num(getAttribute(ftr, null, 'l')) ?? 50000) / 100000;
    const t = (num(getAttribute(ftr, null, 't')) ?? 50000) / 100000;
    const r = (num(getAttribute(ftr, null, 'r')) ?? 50000) / 100000;
    const b = (num(getAttribute(ftr, null, 'b')) ?? 50000) / 100000;
    const cx = (l + (1 - r)) / 2;
    const cy = (t + (1 - b)) / 2;
    p.defs.push(
      `<radialGradient id="${id}" cx="${fmt(cx)}" cy="${fmt(cy)}" r="0.71">${stops.join('')}</radialGradient>`
    );
  } else {
    const ang =
      ((num(getAttribute(child(grad, 'lin'), null, 'ang')) ?? 0) / 60000) * (Math.PI / 180);
    const dx = Math.cos(ang) / 2;
    const dy = Math.sin(ang) / 2;
    p.defs.push(
      `<linearGradient id="${id}" x1="${fmt(0.5 - dx)}" y1="${fmt(0.5 - dy)}" x2="${fmt(0.5 + dx)}" y2="${fmt(0.5 + dy)}">${stops.join('')}</linearGradient>`
    );
  }
  return { kind: 'ref', id };
}

/** `a:pattFill` → foreground hatching over the background colour. */
function patternPaint(p: Painter, patt: XmlElement): Paint {
  const fg = resolveDrawingColor(child(patt, 'fgClr'), p.ctx.theme);
  const bg = resolveDrawingColor(child(patt, 'bgClr'), p.ctx.theme);
  if (!fg && !bg) return { kind: 'none' };
  const id = `g${p.nextId++}`;
  const size = fmt(6 * p.emuPerPx);
  const bgRect = bg
    ? `<rect width="${size}" height="${size}" ${paintAttrs({ kind: 'color', color: bg }, 'fill')}/>`
    : '';
  const fgLine = fg
    ? `<path d="M0 ${size}L${size} 0" ${paintAttrs({ kind: 'color', color: fg }, 'stroke')} stroke-width="1" vector-effect="non-scaling-stroke"/>`
    : '';
  p.defs.push(
    `<pattern id="${id}" patternUnits="userSpaceOnUse" width="${size}" height="${size}">${bgRect}${fgLine}</pattern>`
  );
  return { kind: 'ref', id };
}

/**
 * The fill of a shape: its own `spPr` fill, else the group's for `a:grpFill`,
 * else the style's `a:fillRef`. Null means a fill this module cannot draw.
 */
function fillPaint(
  p: Painter,
  spPr: XmlElement | null,
  style: XmlElement | null,
  groupFill: Paint
): Paint | null {
  const fillRef = child(style, 'fillRef');
  const refColor = resolveDrawingColor(fillRef, p.ctx.theme) ?? undefined;
  for (const el of getChildElements(spPr)) {
    switch (local(el)) {
      case 'noFill':
        return { kind: 'none' };
      case 'solidFill': {
        const c = resolveDrawingColor(el, p.ctx.theme, refColor);
        return c ? { kind: 'color', color: c } : { kind: 'none' };
      }
      case 'gradFill':
        return gradientPaint(p, el, refColor);
      case 'pattFill':
        return patternPaint(p, el);
      case 'grpFill':
        return groupFill;
      case 'blipFill': {
        const rId = blipRelId(child(el, 'blip'));
        const res = rId
          ? resolveImageData(rId, p.ctx.rels ?? undefined, p.ctx.media ?? undefined)
          : null;
        const href = embeddablePicture(res?.src, res?.mimeType);
        return href ? { kind: 'picture', href } : null;
      }
      default:
        break;
    }
  }
  // No explicit fill: the shape style decides (idx 0 = no fill).
  if (fillRef && getAttribute(fillRef, null, 'idx') !== '0' && refColor) {
    return { kind: 'color', color: refColor };
  }
  return { kind: 'none' };
}

const DASHES: Record<string, number[]> = {
  dash: [4, 3],
  sysDash: [3, 1],
  dot: [1, 3],
  sysDot: [1, 1],
  lgDash: [8, 3],
  dashDot: [4, 3, 1, 3],
  sysDashDot: [3, 1, 1, 1],
  lgDashDot: [8, 3, 1, 3],
  lgDashDotDot: [8, 3, 1, 3, 1, 3],
  sysDashDotDot: [3, 1, 1, 1, 1, 1],
};

/** Theme line widths for `a:lnRef@idx` 1..3 in the stock Office theme. */
const THEME_LINE_WIDTHS = [6350, 12700, 19050];

/** SVG stroke attributes for a shape outline, or `stroke="none"`. */
function strokeAttrs(p: Painter, spPr: XmlElement | null, style: XmlElement | null): string {
  const lnRef = child(style, 'lnRef');
  const refIdx = num(getAttribute(lnRef, null, 'idx')) ?? 0;
  const refColor = resolveDrawingColor(lnRef, p.ctx.theme) ?? undefined;
  const ln = child(spPr, 'ln');

  let paint: Paint = { kind: 'none' };
  let widthEmu = THEME_LINE_WIDTHS[Math.min(Math.max(refIdx, 1), 3) - 1];
  if (ln) {
    const w = num(getAttribute(ln, null, 'w'));
    if (w !== null) widthEmu = w;
    const fillEl = getChildElements(ln).find((c) =>
      ['noFill', 'solidFill', 'gradFill', 'pattFill'].includes(local(c))
    );
    if (fillEl && local(fillEl) === 'solidFill') {
      const c = resolveDrawingColor(fillEl, p.ctx.theme, refColor);
      if (c) paint = { kind: 'color', color: c };
    } else if (fillEl && local(fillEl) === 'gradFill') {
      const firstStop = getChildElements(child(fillEl, 'gsLst'))[0];
      const c = resolveDrawingColor(firstStop, p.ctx.theme, refColor);
      if (c) paint = { kind: 'color', color: c };
    } else if (!fillEl && refIdx > 0 && refColor) {
      paint = { kind: 'color', color: refColor };
    }
  } else if (refIdx > 0 && refColor) {
    paint = { kind: 'color', color: refColor };
  }
  if (paint.kind === 'none') return 'stroke="none"';

  const px = Math.max(0.5, widthEmu / EMU_PER_PX);
  let out = `${paintAttrs(paint, 'stroke')} stroke-width="${fmt(px)}" vector-effect="non-scaling-stroke"`;
  const dash = DASHES[getAttribute(child(ln, 'prstDash'), null, 'val') ?? ''];
  if (dash) out += ` stroke-dasharray="${dash.map((d) => fmt(d * Math.max(1, px))).join(' ')}"`;
  if (child(ln, 'round')) out += ' stroke-linejoin="round"';
  else if (child(ln, 'bevel')) out += ' stroke-linejoin="bevel"';
  const cap = getAttribute(ln, null, 'cap');
  if (cap === 'rnd') out += ' stroke-linecap="round"';
  else if (cap === 'sq') out += ' stroke-linecap="square"';
  return out;
}

/** A dashed, labelled box marking content that is reserved but not drawn. */
function placeholder(
  p: Painter,
  b: Box,
  label: string,
  xf: { rot?: number; flipH?: boolean; flipV?: boolean } = {},
  tint = true
): void {
  p.unsupported.add(label);
  const transform = boxTransform(b, xf.rot ?? 0, !!xf.flipH, !!xf.flipV);
  const t = transform ? ` transform="${transform}"` : '';
  const fill = tint ? 'fill="#9ca3af" fill-opacity="0.15"' : 'fill="none"';
  p.body.push(
    `<g${t}><rect x="${fmt(b.x)}" y="${fmt(b.y)}" width="${fmt(Math.max(0, b.w))}" height="${fmt(Math.max(0, b.h))}" ${fill} ` +
      `stroke="#6b7280" stroke-width="1" stroke-dasharray="4 3" vector-effect="non-scaling-stroke"/></g>`
  );
  // Label only where it fits legibly; a sliver of a box stays just dashed.
  // Labels live in a pixel-space layer: browsers clamp font sizes, so an
  // 11 px label written in EMU (104775 user units) would render at ~1 px.
  const px: Box = {
    x: b.x / p.emuPerPx,
    y: b.y / p.emuPerPx,
    w: b.w / p.emuPerPx,
    h: b.h / p.emuPerPx,
  };
  const fontPx = Math.min(11, px.h * 0.5, (px.w * 0.92) / (label.length * 0.56));
  if (fontPx < 6) return;
  const lt = boxTransform(px, xf.rot ?? 0, false, false);
  p.labels.push(
    `<text${lt ? ` transform="${lt}"` : ''} x="${fmt(px.x + px.w / 2)}" y="${fmt(px.y + px.h / 2)}" font-size="${fmt(fontPx)}" ` +
      `font-family="Arial, Helvetica, sans-serif" fill="#4b5563" text-anchor="middle" dominant-baseline="central">` +
      `${escapeXmlText(label)}</text>`
  );
}

/** `wps:wsp` → its geometry painted in box `b` (root space). */
function paintShape(p: Painter, wsp: XmlElement, b: Box, groupFill: Paint, x: Xfrm | null): void {
  const spPr = child(wsp, 'spPr');
  const style = child(wsp, 'style');
  const rot = x?.rot ?? 0;
  const transform = boxTransform(b, rot, !!x?.flipH, !!x?.flipV);
  const hasText = child(wsp, 'txbx') !== null || child(wsp, 'linkedTxbx') !== null;

  const paths = geometryPaths(spPr, b.w, b.h);
  const fill = paths ? fillPaint(p, spPr, style, groupFill) : null;
  if (!paths || !fill) {
    placeholder(
      p,
      b,
      hasText ? UNSUPPORTED_DRAWING_LABELS.textbox : UNSUPPORTED_DRAWING_LABELS.shape,
      x ?? {}
    );
    return;
  }

  const stroke = strokeAttrs(p, spPr, style);
  const inner = `translate(${fmt(b.x)} ${fmt(b.y)})`;
  const g = transform ? `${transform} ${inner}` : inner;
  let out = `<g transform="${g}">`;
  for (const path of paths) {
    if (fill.kind === 'picture' && path.fill) {
      const clip = `c${p.nextId++}`;
      p.defs.push(`<clipPath id="${clip}"><path d="${path.d}"/></clipPath>`);
      out += `<image href="${fill.href}" x="0" y="0" width="${fmt(b.w)}" height="${fmt(b.h)}" preserveAspectRatio="none" clip-path="url(#${clip})"/>`;
    }
    const f = path.fill && fill.kind !== 'picture' ? paintAttrs(fill, 'fill') : 'fill="none"';
    const s = path.stroke ? stroke : 'stroke="none"';
    if (f === 'fill="none"' && s === 'stroke="none"') continue;
    out += `<path d="${path.d}" ${f} ${s}/>`;
  }
  p.body.push(`${out}</g>`);
  p.drew = true;
  // A text box nested in a group has no editable model: mark its text area.
  if (hasText) placeholder(p, b, UNSUPPORTED_DRAWING_LABELS.textbox, x ?? {}, false);
}

/** `pic:pic` inside a group → the picture, cropped by `a:srcRect`. */
function paintPicture(p: Painter, pic: XmlElement, b: Box, x: Xfrm | null): void {
  const transform = boxTransform(b, x?.rot ?? 0, !!x?.flipH, !!x?.flipV);
  const blipFill = child(pic, 'blipFill');
  const rId = blipRelId(child(blipFill, 'blip'));
  const res = rId ? resolveImageData(rId, p.ctx.rels ?? undefined, p.ctx.media ?? undefined) : null;
  const href = embeddablePicture(res?.src, res?.mimeType);
  if (!href) {
    // Externally loaded media cannot be embedded in the SVG; the caller paints
    // those pictures as separate images instead.
    if (res?.assetId) p.externalPictures = true;
    else placeholder(p, b, UNSUPPORTED_DRAWING_LABELS.picture, x ?? {});
    return;
  }
  const src = child(blipFill, 'srcRect');
  const crop = (a: string) =>
    Math.max(0, Math.min(1, (num(getAttribute(src, null, a)) ?? 0) / 100000));
  const [l, t, r, bt] = [crop('l'), crop('t'), crop('r'), crop('b')];
  const vw = Math.max(0.0001, 1 - l - r);
  const vh = Math.max(0.0001, 1 - t - bt);
  const tAttr = transform ? ` transform="${transform}"` : '';
  p.body.push(
    `<g${tAttr}><svg x="${fmt(b.x)}" y="${fmt(b.y)}" width="${fmt(b.w)}" height="${fmt(b.h)}" ` +
      `viewBox="${l} ${t} ${vw} ${vh}" preserveAspectRatio="none">` +
      `<image href="${href}" x="0" y="0" width="1" height="1" preserveAspectRatio="none"/></svg></g>`
  );
  p.drew = true;
}

/** The children of a group container, painted in document (z) order. */
function paintGroupChildren(
  p: Painter,
  container: XmlElement,
  frame: Frame,
  groupFill: Paint,
  depth: number
): void {
  if (depth >= MAX_DEPTH) return;
  for (const c of getChildElements(container)) {
    if (++p.items > MAX_ITEMS) return;
    const name = local(c);
    if (name === 'wsp') {
      const x = readXfrm(child(child(c, 'spPr'), 'xfrm'));
      const b = x && mapBox(frame, x);
      if (b) paintShape(p, c, b, groupFill, x);
    } else if (name === 'pic') {
      const x = readXfrm(child(child(c, 'spPr'), 'xfrm'));
      const b = x && mapBox(frame, x);
      if (b) paintPicture(p, c, b, x);
    } else if (name === 'grpSp') {
      const grpSpPr = child(c, 'grpSpPr');
      const x = readXfrm(child(grpSpPr, 'xfrm'));
      const b = x && mapBox(frame, x);
      if (!x || !b) continue;
      const chExt = x.chExt ?? x.ext!;
      const chOff = x.chOff ?? { x: 0, y: 0 };
      const sx = chExt.x === 0 ? frame.sx : b.w / chExt.x;
      const sy = chExt.y === 0 ? frame.sy : b.h / chExt.y;
      const inner: Frame = { sx, sy, tx: b.x - chOff.x * sx, ty: b.y - chOff.y * sy };
      const fill = fillPaint(p, grpSpPr, null, groupFill) ?? groupFill;
      const transform = boxTransform(b, x.rot, x.flipH, x.flipV);
      if (transform) p.body.push(`<g transform="${transform}">`);
      paintGroupChildren(p, c, inner, fill, depth + 1);
      if (transform) p.body.push('</g>');
    } else if (name === 'graphicFrame') {
      // A chart / diagram / picture frame nested in the group.
      const x = readXfrm(child(c, 'xfrm'));
      const b = x && mapBox(frame, x);
      const uri = getAttribute(child(child(c, 'graphic'), 'graphicData'), null, 'uri') ?? '';
      const kind = GRAPHIC_URI_KIND.get(uri);
      const label =
        kind && kind in UNSUPPORTED_DRAWING_LABELS
          ? UNSUPPORTED_DRAWING_LABELS[kind as UnsupportedDrawingKind]
          : UNSUPPORTED_DRAWING_LABELS.graphic;
      if (b) placeholder(p, b, label, x!);
    }
  }
}

function svgDataUri(p: Painter, wEmu: number, hEmu: number): string {
  const wPx = Math.max(1, wEmu / p.emuPerPx);
  const hPx = Math.max(1, hEmu / p.emuPerPx);
  const defs = p.defs.length ? `<defs>${p.defs.join('')}</defs>` : '';
  const labelLayer = p.labels.length
    ? `<svg x="0" y="0" width="${fmt(Math.max(1, wEmu))}" height="${fmt(Math.max(1, hEmu))}" ` +
      `viewBox="0 0 ${fmt(wPx)} ${fmt(hPx)}" preserveAspectRatio="none" overflow="visible">${p.labels.join('')}</svg>`
    : '';
  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" width="${fmt(wPx)}" height="${fmt(hPx)}" ` +
    `viewBox="0 0 ${fmt(Math.max(1, wEmu))} ${fmt(Math.max(1, hEmu))}" preserveAspectRatio="none">` +
    `${defs}${p.body.join('')}${labelLayer}</svg>`;
  return `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
}

function newPainter(ctx: DrawingPreviewContext): Painter {
  return {
    ctx,
    defs: [],
    body: [],
    labels: [],
    nextId: 0,
    items: 0,
    emuPerPx: EMU_PER_PX,
    unsupported: new Set(),
    externalPictures: false,
    drew: false,
  };
}

/** The anchor/inline frame of a drawing, as an image with no picture of its own. */
function frameImage(drawingEl: XmlElement, ctx: DrawingPreviewContext): Image | null {
  const parsed = parseDrawing(drawingEl, ctx.rels ?? undefined, ctx.media ?? undefined);
  if (!parsed || !(parsed.size.width > 0) || !(parsed.size.height > 0)) return null;
  const frame: Image = {
    type: 'image',
    rId: '',
    size: parsed.size,
    wrap: parsed.wrap,
    renderOnly: true,
  };
  if (parsed.id) frame.id = parsed.id;
  if (parsed.alt) frame.alt = parsed.alt;
  if (parsed.title) frame.title = parsed.title;
  if (parsed.decorative) frame.decorative = true;
  if (parsed.position) frame.position = parsed.position;
  if (parsed.padding) frame.padding = parsed.padding;
  if (parsed.layoutInCell !== undefined) frame.layoutInCell = parsed.layoutInCell;
  if (parsed.allowOverlap !== undefined) frame.allowOverlap = parsed.allowOverlap;
  if (parsed.relativeHeight !== undefined) frame.relativeHeight = parsed.relativeHeight;
  return frame;
}

function rootTransform(x: Xfrm | null): ImageTransform | undefined {
  if (!x || (!x.rot && !x.flipH && !x.flipV)) return undefined;
  const t: ImageTransform = {};
  if (x.rot) t.rotation = x.rot;
  if (x.flipH) t.flipH = true;
  if (x.flipV) t.flipV = true;
  return t;
}

function finish(p: Painter, frame: Image, kindLabel: string): DrawingContent {
  frame.src = svgDataUri(p, frame.size.width, frame.size.height);
  frame.mimeType = 'image/svg+xml';
  // Mark the drawing as a placeholder for hosts and assistive tech; a partly
  // drawn one keeps its author-supplied description.
  if (!p.drew) frame.alt = p.unsupported.size === 1 ? [...p.unsupported][0] : kindLabel;
  return { type: 'drawing', image: frame };
}

/**
 * Render-only preview for a `w:drawing` the model does not represent (any
 * graphic but a picture or a text box). The first item carries the drawing's
 * frame; a group whose pictures are externally loaded is followed by those
 * pictures as separate images.
 */
export function deriveDrawingPreview(
  drawingEl: XmlElement,
  ctx: DrawingPreviewContext
): DrawingContent[] {
  const kind = drawingGraphicKind(drawingEl);
  if (!kind || kind === 'picture') return [];
  const frame = frameImage(drawingEl, ctx);
  if (!frame) return [];
  const { graphicData } = drawingParts(drawingEl);
  const p = newPainter(ctx);
  const whole: Box = { x: 0, y: 0, w: frame.size.width, h: frame.size.height };

  if (kind === 'shape') {
    const wsp = child(graphicData, 'wsp');
    const x = readXfrm(child(child(wsp, 'spPr'), 'xfrm'));
    if (wsp) paintShape(p, wsp, whole, { kind: 'none' }, null);
    const t = rootTransform(x);
    if (t) frame.transform = t;
    return [finish(p, frame, UNSUPPORTED_DRAWING_LABELS.shape)];
  }

  if (kind === 'group') {
    const wgp = child(graphicData, 'wgp');
    const grpSpPr = child(wgp, 'grpSpPr');
    const x = readXfrm(child(grpSpPr, 'xfrm'));
    // The anchor's extent is authoritative for the group's size on the page;
    // children map from the group's child space onto it.
    const chExt = x?.chExt ?? x?.ext ?? { x: whole.w, y: whole.h };
    const chOff = x?.chOff ?? { x: 0, y: 0 };
    const sx = chExt.x === 0 ? 1 : whole.w / chExt.x;
    const sy = chExt.y === 0 ? 1 : whole.h / chExt.y;
    const groupFill = fillPaint(p, grpSpPr, null, { kind: 'none' }) ?? { kind: 'none' };
    if (wgp)
      paintGroupChildren(p, wgp, { sx, sy, tx: -chOff.x * sx, ty: -chOff.y * sy }, groupFill, 0);
    const t = rootTransform(x);
    if (t) frame.transform = t;
    const out = [finish(p, frame, UNSUPPORTED_DRAWING_LABELS.group)];
    if (p.externalPictures) out.push(...deriveGroupPreviewImages(drawingEl, ctx.rels, ctx.media));
    return out;
  }

  placeholder(p, whole, UNSUPPORTED_DRAWING_LABELS[kind]);
  return [finish(p, frame, UNSUPPORTED_DRAWING_LABELS[kind])];
}

// ============================================================================
// VML
// ============================================================================

const VML_SHAPES = new Set([
  'shape',
  'rect',
  'roundrect',
  'oval',
  'line',
  'polyline',
  'arc',
  'curve',
  'group',
  'image',
]);

/** CSS length (pt/in/cm/mm/pc/px, default px) → EMU. */
function cssLengthToEmu(raw: string | undefined): number | null {
  const m = /^(-?[\d.]+)\s*(pt|in|px|cm|mm|pc)?$/.exec((raw ?? '').trim());
  if (!m) return null;
  const v = parseFloat(m[1]);
  if (!Number.isFinite(v)) return null;
  const perUnit: Record<string, number> = {
    pt: 12700,
    in: 914400,
    cm: 360000,
    mm: 36000,
    pc: 152400,
    px: 9525,
  };
  return Math.round(v * perUnit[m[2] ?? 'px']);
}

/**
 * Render-only placeholder for a legacy VML drawing (`w:pict` / `w:object`)
 * with no picture: its box, at its position when absolutely positioned. VML
 * shapes are not drawn; a watermark is left to the watermark parser.
 */
export function deriveVmlPreview(pictEl: XmlElement): DrawingContent[] {
  const shape = getChildElements(pictEl).find((c) => VML_SHAPES.has(local(c)));
  if (!shape) return [];
  if (isWatermarkShape(shape, (getAttribute(shape, null, 'id') ?? '').toLowerCase())) return [];

  const style = parseStyleAttr(getAttribute(shape, null, 'style'));
  const w = cssLengthToEmu(style['width']);
  const h = cssLengthToEmu(style['height']);
  if (!w || !h || w <= 0 || h <= 0) return [];

  const frame: Image = {
    type: 'image',
    rId: '',
    size: { width: w, height: h },
    wrap: { type: 'inline' },
    renderOnly: true,
  };
  if (style['position'] === 'absolute') {
    const wrapEl = getChildElements(shape).find((c) => local(c) === 'wrap');
    const wrapType = getAttribute(wrapEl, null, 'type');
    const z = num(style['z-index']) ?? 0;
    frame.wrap =
      wrapType === 'topAndBottom'
        ? { type: 'topAndBottom' }
        : wrapType === 'square' || wrapType === 'tight' || wrapType === 'through'
          ? { type: wrapType }
          : { type: z < 0 ? 'behind' : 'inFront' };
    const hRel = style['mso-position-horizontal-relative'];
    const vRel = style['mso-position-vertical-relative'];
    const hAlign = style['mso-position-horizontal'];
    const vAlign = style['mso-position-vertical'];
    const H_REL: Record<string, 'page' | 'margin' | 'column' | 'character'> = {
      page: 'page',
      margin: 'margin',
      text: 'column',
      char: 'character',
    };
    const V_REL: Record<string, 'page' | 'margin' | 'paragraph' | 'line'> = {
      page: 'page',
      margin: 'margin',
      text: 'paragraph',
      line: 'line',
    };
    const H_ALIGN = new Set(['left', 'center', 'right', 'inside', 'outside']);
    const V_ALIGN = new Set(['top', 'center', 'bottom', 'inside', 'outside']);
    frame.position = {
      horizontal: {
        relativeTo: H_REL[hRel ?? ''] ?? 'column',
        ...(hAlign && H_ALIGN.has(hAlign)
          ? { alignment: hAlign as 'left' }
          : {
              posOffset: cssLengthToEmu(style['margin-left']) ?? cssLengthToEmu(style['left']) ?? 0,
            }),
      },
      vertical: {
        relativeTo: V_REL[vRel ?? ''] ?? 'paragraph',
        ...(vAlign && V_ALIGN.has(vAlign)
          ? { alignment: vAlign as 'top' }
          : {
              posOffset: cssLengthToEmu(style['margin-top']) ?? cssLengthToEmu(style['top']) ?? 0,
            }),
      },
    };
  }

  const isTextBox = getChildElements(shape).some((c) => local(c) === 'textbox');
  const label =
    local(shape) === 'group'
      ? UNSUPPORTED_DRAWING_LABELS.group
      : isTextBox
        ? UNSUPPORTED_DRAWING_LABELS.textbox
        : UNSUPPORTED_DRAWING_LABELS.shape;
  const p = newPainter({ rels: null, media: null, theme: null });
  placeholder(p, { x: 0, y: 0, w, h }, label);
  return [finish(p, frame, label)];
}

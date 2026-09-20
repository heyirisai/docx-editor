/**
 * DrawingML serialization — inline and floating images, shapes, and text
 * boxes. Dispatched from runSerializer.ts for `drawing` and `shape` runs.
 *
 * Auto-incrementing IDs for pasted images/shapes are scoped to this
 * module; reset with `resetAutoIdCounter` before each serialization pass.
 */

import type {
  ColorValue,
  DrawingContent,
  Image,
  ImagePosition,
  ImageWrap,
  Shape,
  ShapeContent,
  ShapeFill,
  ShapeOutline,
  ShapeTextBody,
  ShapeBlockContent,
} from '../../../types/document';
import { serializeParagraph } from '../paragraphSerializer';
import { serializeTable } from '../tableSerializer';
import { escapeXml, intAttr } from '../xmlUtils';
import { preservedBodyPrXml, preservedSpPrExtra } from '../../preservedShapeXml';

/**
 * Auto-incrementing counter for generating unique image/shape IDs.
 * Used as a fallback when `image.id` or `shape.id` is undefined (e.g., pasted images).
 * Starts high (100000) to avoid collisions with IDs parsed from existing DOCX content.
 */
let nextAutoId = 100000;

/**
 * Reset the auto-incrementing ID counter. Call before each serialization pass
 * to keep IDs deterministic across saves.
 */
export function resetAutoIdCounter(): void {
  nextAutoId = 100000;
}

/** Get a unique positive integer ID, using the provided value or generating one */
function getUniqueId(id: string | number | undefined): string {
  if (id !== undefined && id !== null && id !== '' && id !== 0) {
    return String(id);
  }
  return String(nextAutoId++);
}

/** Serialize a color value to DrawingML a:srgbClr or a:schemeClr */
function serializeDrawingColor(color: ColorValue | undefined): string {
  if (!color) return '';
  if (color.rgb) {
    return `<a:srgbClr val="${color.rgb.replace('#', '')}"/>`;
  }
  if (color.themeColor) {
    let clr = `<a:schemeClr val="${color.themeColor}"`;
    if (color.themeTint) {
      clr += `><a:tint val="${color.themeTint}"/></a:schemeClr>`;
    } else if (color.themeShade) {
      clr += `><a:shade val="${color.themeShade}"/></a:schemeClr>`;
    } else {
      clr += `/>`;
    }
    return clr;
  }
  return '';
}

/** Serialize shape fill to DrawingML */
function serializeFill(fill: ShapeFill | undefined): string {
  if (!fill || fill.type === 'none') return '<a:noFill/>';
  if (fill.type === 'solid' && fill.color) {
    return `<a:solidFill>${serializeDrawingColor(fill.color)}</a:solidFill>`;
  }
  if (fill.type === 'gradient' && fill.gradient) {
    const g = fill.gradient;
    const stops = g.stops
      .map((s) => `<a:gs pos="${s.position}">${serializeDrawingColor(s.color)}</a:gs>`)
      .join('');
    const direction =
      g.type === 'linear' ? `<a:lin ang="${(g.angle || 0) * 60000}" scaled="1"/>` : '';
    return `<a:gradFill><a:gsLst>${stops}</a:gsLst>${direction}</a:gradFill>`;
  }
  return '';
}

/** Serialize shape outline to DrawingML a:ln */
/**
 * `anchor` back to its schema token. The model names the positions; DrawingML
 * spells them `t`/`ctr`/`b`/`dist`/`just`, and anything else makes Word reject
 * the shape — the inverse of the map `parseBodyProperties` reads them with.
 */
const BODY_PR_ANCHOR: Record<NonNullable<ShapeTextBody['anchor']>, string> = {
  top: 't',
  middle: 'ctr',
  bottom: 'b',
  distributed: 'dist',
  justified: 'just',
};

/**
 * The `wps:bodyPr` attributes the model parses, as they would be written. Every
 * other attribute on the element belongs to the source.
 */
function modelledBodyPrAttrs(tb: ShapeTextBody): Record<string, string> {
  const attrs: Record<string, string> = {};
  if (tb.anchor) attrs.anchor = BODY_PR_ANCHOR[tb.anchor];
  if (tb.anchorCenter) attrs.anchorCtr = '1';
  const insets: [string, number | undefined][] = [
    ['lIns', tb.margins?.left],
    ['tIns', tb.margins?.top],
    ['rIns', tb.margins?.right],
    ['bIns', tb.margins?.bottom],
  ];
  for (const [name, value] of insets) {
    if (value != null) attrs[name] = String(intAttr(value));
  }
  return attrs;
}

/** `wps:spPr` children below the fill: the line, then the effects. */
/**
 * `a:prstGeom`. A modelled geometry wins: a `roundRect` pill, an `ellipse`
 * badge or a circular headshot re-serialized as the default `rect` came back
 * with square corners. Shared by `wps:spPr` and `pic:spPr` — the preset means
 * the same thing on a picture as on a shape.
 */
function presetGeometryXml(
  geometry: 'ellipse' | 'roundRect' | undefined,
  cornerAdj: number | undefined,
  fallbackPrst: string
): string {
  if (geometry === 'ellipse') return '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>';
  if (geometry === 'roundRect') {
    const avLst =
      cornerAdj === undefined
        ? '<a:avLst/>'
        : `<a:avLst><a:gd name="adj" fmla="val ${Math.round(cornerAdj * 100000)}"/></a:avLst>`;
    return `<a:prstGeom prst="roundRect">${avLst}</a:prstGeom>`;
  }
  return `<a:prstGeom prst="${fallbackPrst}"><a:avLst/></a:prstGeom>`;
}

/**
 * A shape's `a:prstGeom`. `shapeType` 'textBox' is spelled `rect` in the file.
 */
function presetGeometry(shape: Shape): string {
  return presetGeometryXml(
    shape.geometry,
    shape.cornerAdj,
    shape.shapeType === 'textBox' ? 'rect' : shape.shapeType
  );
}

function spPrExtras(shape: Shape): string {
  const preserved = preservedSpPrExtra(shape.spPrExtraXml);
  const ln = shape.outline ? serializeOutline(shape.outline) : (preserved.ln ?? '');
  return `${ln}${preserved.effectLst ?? ''}`;
}

function serializeOutline(outline: ShapeOutline | undefined): string {
  if (!outline) return '';
  const attrs: string[] = [];
  if (outline.width != null) attrs.push(`w="${outline.width}"`);
  if (outline.cap) attrs.push(`cap="${outline.cap}"`);

  const parts: string[] = [];
  if (outline.color) {
    parts.push(`<a:solidFill>${serializeDrawingColor(outline.color)}</a:solidFill>`);
  }
  if (outline.style && outline.style !== 'solid') {
    parts.push(`<a:prstDash val="${outline.style}"/>`);
  }
  if (outline.headEnd) {
    parts.push(
      `<a:headEnd type="${outline.headEnd.type}"${outline.headEnd.width ? ` w="${outline.headEnd.width}"` : ''}${outline.headEnd.length ? ` len="${outline.headEnd.length}"` : ''}/>`
    );
  }
  if (outline.tailEnd) {
    parts.push(
      `<a:tailEnd type="${outline.tailEnd.type}"${outline.tailEnd.width ? ` w="${outline.tailEnd.width}"` : ''}${outline.tailEnd.length ? ` len="${outline.tailEnd.length}"` : ''}/>`
    );
  }

  if (parts.length === 0 && attrs.length === 0) return '';
  return `<a:ln${attrs.length ? ' ' + attrs.join(' ') : ''}>${parts.join('')}</a:ln>`;
}

/** Build wp:positionH and wp:positionV for floating drawings */
function serializePosition(pos: ImagePosition): string {
  const parts: string[] = [];

  // Horizontal
  const h = pos.horizontal;
  parts.push(`<wp:positionH relativeFrom="${h.relativeTo}">`);
  if (h.alignment) {
    parts.push(`<wp:align>${h.alignment}</wp:align>`);
  } else {
    parts.push(`<wp:posOffset>${intAttr(h.posOffset)}</wp:posOffset>`);
  }
  parts.push('</wp:positionH>');

  // Vertical
  const v = pos.vertical;
  parts.push(`<wp:positionV relativeFrom="${v.relativeTo}">`);
  if (v.alignment) {
    parts.push(`<wp:align>${v.alignment}</wp:align>`);
  } else {
    parts.push(`<wp:posOffset>${intAttr(v.posOffset)}</wp:posOffset>`);
  }
  parts.push('</wp:positionV>');

  return parts.join('');
}

/** Serialize wrap type to wp:wrap* element */
function serializeWrap(wrap: ImageWrap): string {
  const wrapText = wrap.wrapText ? ` wrapText="${wrap.wrapText}"` : ' wrapText="bothSides"';
  switch (wrap.type) {
    case 'square':
      return `<wp:wrapSquare${wrapText}/>`;
    case 'tight':
      return `<wp:wrapTight${wrapText}><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon></wp:wrapTight>`;
    case 'through':
      return `<wp:wrapThrough${wrapText}><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon></wp:wrapThrough>`;
    case 'topAndBottom':
      return '<wp:wrapTopAndBottom/>';
    case 'behind':
    case 'inFront':
      return '<wp:wrapNone/>';
    default:
      return '<wp:wrapNone/>';
  }
}

/** Build the common a:graphic > pic:pic element for images */
function serializePicGraphic(image: Image, sharedId: string): string {
  const cx = image.size.width;
  const cy = image.size.height;
  const rId = image.rId || 'rId1';
  const id = sharedId;
  const name = image.filename || `image${id}`;

  let xfrmAttrs = '';
  if (image.transform?.rotation) {
    xfrmAttrs += ` rot="${Math.round(image.transform.rotation * 60000)}"`;
  }
  if (image.transform?.flipH) xfrmAttrs += ' flipH="1"';
  if (image.transform?.flipV) xfrmAttrs += ' flipV="1"';

  // Build <a:blip> with optional <a:alphaModFix> child for transparency.
  const alphaChild =
    image.opacity !== undefined && image.opacity < 1
      ? `<a:alphaModFix amt="${Math.round(Math.max(0, Math.min(1, image.opacity)) * 100000)}"/>`
      : '';
  const blipEl = alphaChild
    ? `<a:blip r:embed="${rId}">${alphaChild}</a:blip>`
    : `<a:blip r:embed="${rId}"/>`;

  // Build optional <a:srcRect/> for wp:srcRect crop. Each side is a
  // fraction in [0, 1]; OOXML expects 1/100000 units.
  const cropAttrs: string[] = [];
  if (image.crop?.left) cropAttrs.push(`l="${Math.round(image.crop.left * 100000)}"`);
  if (image.crop?.top) cropAttrs.push(`t="${Math.round(image.crop.top * 100000)}"`);
  if (image.crop?.right) cropAttrs.push(`r="${Math.round(image.crop.right * 100000)}"`);
  if (image.crop?.bottom) cropAttrs.push(`b="${Math.round(image.crop.bottom * 100000)}"`);
  const srcRectEl = cropAttrs.length > 0 ? `<a:srcRect ${cropAttrs.join(' ')}/>` : '';

  return [
    '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">',
    '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">',
    '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">',
    '<pic:nvPicPr>',
    `<pic:cNvPr id="${id}" name="${escapeXml(name)}"${image.alt ? ` descr="${escapeXml(image.alt)}"` : ''}/>`,
    '<pic:cNvPicPr/>',
    '</pic:nvPicPr>',
    '<pic:blipFill>',
    blipEl,
    srcRectEl,
    '<a:stretch><a:fillRect/></a:stretch>',
    '</pic:blipFill>',
    '<pic:spPr>',
    `<a:xfrm${xfrmAttrs}>`,
    '<a:off x="0" y="0"/>',
    `<a:ext cx="${intAttr(cx)}" cy="${intAttr(cy)}"/>`,
    '</a:xfrm>',
    presetGeometryXml(image.geometry, image.cornerAdj, 'rect'),
    image.outline ? serializeOutline(image.outline) : '',
    '</pic:spPr>',
    '</pic:pic>',
    '</a:graphicData>',
    '</a:graphic>',
  ].join('');
}

/**
 * Serialize drawing/image content (w:drawing) to full DrawingML XML
 */
export function serializeDrawingContent(content: DrawingContent): string {
  const image = content.image;
  const isFloating = image.wrap.type !== 'inline';
  const cx = image.size.width;
  const cy = image.size.height;
  // dist* on wp:inline / wp:anchor are text-wrap distances; effectExtent is
  // a separate element. Don't conflate `image.padding` (effectExtent) into
  // the wrap distance attributes.
  const distT = image.wrap.distT ?? 0;
  const distB = image.wrap.distB ?? 0;
  const distL = image.wrap.distL ?? 0;
  const distR = image.wrap.distR ?? 0;
  const effL = image.padding?.left ?? 0;
  const effT = image.padding?.top ?? 0;
  const effR = image.padding?.right ?? 0;
  const effB = image.padding?.bottom ?? 0;
  const effectExtentEl = `<wp:effectExtent l="${intAttr(effL)}" t="${intAttr(effT)}" r="${intAttr(effR)}" b="${intAttr(effB)}"/>`;
  const docPrId = getUniqueId(image.id);
  const docPrName = image.title || image.filename || `Picture ${docPrId}`;

  const graphic = serializePicGraphic(image, docPrId);

  if (!isFloating) {
    // Inline image
    return [
      '<w:drawing>',
      `<wp:inline distT="${intAttr(distT)}" distB="${intAttr(distB)}" distL="${intAttr(distL)}" distR="${intAttr(distR)}">`,
      `<wp:extent cx="${intAttr(cx)}" cy="${intAttr(cy)}"/>`,
      effectExtentEl,
      `<wp:docPr id="${docPrId}" name="${escapeXml(docPrName)}"${image.alt ? ` descr="${escapeXml(image.alt)}"` : ''}${image.decorative ? ' hidden="1"' : ''}/>`,
      '<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr>',
      graphic,
      '</wp:inline>',
      '</w:drawing>',
    ].join('');
  }

  // Floating (anchored) image
  const behindDoc = image.wrap.type === 'behind' ? '1' : '0';
  const position = image.position
    ? serializePosition(image.position)
    : '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>';
  const wrap = serializeWrap(image.wrap);

  return [
    '<w:drawing>',
    `<wp:anchor distT="${intAttr(distT)}" distB="${intAttr(distB)}" distL="${intAttr(distL)}" distR="${intAttr(distR)}" simplePos="0" relativeHeight="${intAttr(image.relativeHeight ?? 251658240)}" behindDoc="${behindDoc}" locked="0" layoutInCell="${image.layoutInCell === false ? '0' : '1'}" allowOverlap="${image.allowOverlap === false ? '0' : '1'}">`,
    '<wp:simplePos x="0" y="0"/>',
    position,
    `<wp:extent cx="${intAttr(cx)}" cy="${intAttr(cy)}"/>`,
    effectExtentEl,
    wrap,
    `<wp:docPr id="${docPrId}" name="${escapeXml(docPrName)}"${image.alt ? ` descr="${escapeXml(image.alt)}"` : ''}/>`,
    '<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr>',
    graphic,
    '</wp:anchor>',
    '</w:drawing>',
  ].join('');
}

/** Serialize text body content for shapes/textboxes */
function serializeShapeTextBody(blocks: ShapeBlockContent[]): string {
  // `w:txbxContent` is EG_BlockLevelElts, so a text box round-trips tables
  // as well as paragraphs.
  return blocks
    .map((b) => (b.type === 'table' ? serializeTable(b) : serializeParagraph(b)))
    .join('');
}

/**
 * Serialize shape content to full DrawingML XML (wps:wsp inside w:drawing)
 */
export function serializeShapeContent(content: ShapeContent): string {
  const shape = content.shape;
  const cx = shape.size.width;
  const cy = shape.size.height;
  const isTextBox = shape.shapeType === 'textBox';
  const isFloating = shape.wrap && shape.wrap.type !== 'inline';
  const distT = shape.wrap?.distT ?? 0;
  const distB = shape.wrap?.distB ?? 0;
  const distL = shape.wrap?.distL ?? 0;
  const distR = shape.wrap?.distR ?? 0;
  const docPrId = getUniqueId(shape.id);
  const docPrName = shape.name || (isTextBox ? `TextBox ${docPrId}` : `Shape ${docPrId}`);

  // Build xfrm
  let xfrmAttrs = '';
  if (shape.transform?.rotation) {
    xfrmAttrs += ` rot="${Math.round(shape.transform.rotation * 60000)}"`;
  }
  if (shape.transform?.flipH) xfrmAttrs += ' flipH="1"';
  if (shape.transform?.flipV) xfrmAttrs += ' flipV="1"';

  // Build wps:spPr
  const spPr = [
    '<wps:spPr>',
    `<a:xfrm${xfrmAttrs}>`,
    '<a:off x="0" y="0"/>',
    `<a:ext cx="${intAttr(cx)}" cy="${intAttr(cy)}"/>`,
    '</a:xfrm>',
    presetGeometry(shape),
    serializeFill(shape.fill),
    // A model outline wins over the source `<a:ln>` — the user may have set one
    // through the UI. With nothing modelled, replay the source's: an explicit
    // "no outline" parses to no outline at all, so rebuilding from the model
    // alone put the DEFAULT outline back on the shape. `<a:effectLst>` is
    // unrelated to either and is replayed whichever line wins.
    spPrExtras(shape),
    '</wps:spPr>',
  ].join('');

  // Build text body if present
  let textBody = '';
  if (shape.textBody) {
    const tb = shape.textBody;
    const modelled = modelledBodyPrAttrs(tb);

    // The source element carries a dozen attributes plus an autofit child and
    // the model holds five of them, so replay it when we have it, with the
    // modelled attributes written back over it. Nothing in `bodyPr` depends on
    // the shape's size or its text, so the rest stays correct across an edit.
    // See ShapeTextBody.bodyPrXml.
    const bodyPr =
      preservedBodyPrXml(tb.bodyPrXml, { overrides: modelled }) ??
      `<wps:bodyPr rot="0" vert="horz"${Object.entries(modelled)
        .map(([name, value]) => ` ${name}="${value}"`)
        .join('')}/>`;

    if (isTextBox) {
      textBody = [
        '<wps:txbx><w:txbxContent>',
        serializeShapeTextBody(tb.content),
        '</w:txbxContent></wps:txbx>',
        bodyPr,
      ].join('');
    } else {
      textBody = bodyPr;
    }
  }

  // Build wps:wsp
  const wsp = [
    '<wps:wsp>',
    `<wps:cNvSpPr${isTextBox ? ' txBox="1"' : ''}/>`,
    spPr,
    textBody,
    '</wps:wsp>',
  ].join('');

  // Wrap in a:graphic
  const graphic = [
    '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">',
    '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">',
    wsp,
    '</a:graphicData>',
    '</a:graphic>',
  ].join('');

  if (!isFloating) {
    return [
      '<w:drawing>',
      `<wp:inline distT="${intAttr(distT)}" distB="${intAttr(distB)}" distL="${intAttr(distL)}" distR="${intAttr(distR)}">`,
      `<wp:extent cx="${intAttr(cx)}" cy="${intAttr(cy)}"/>`,
      '<wp:effectExtent l="0" t="0" r="0" b="0"/>',
      `<wp:docPr id="${docPrId}" name="${escapeXml(docPrName)}"/>`,
      '<wp:cNvGraphicFramePr/>',
      graphic,
      '</wp:inline>',
      '</w:drawing>',
    ].join('');
  }

  // Floating shape
  const behindDoc = shape.wrap?.type === 'behind' ? '1' : '0';
  const position = shape.position
    ? serializePosition(shape.position)
    : '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>';
  const wrap = serializeWrap(shape.wrap!);

  return [
    '<w:drawing>',
    `<wp:anchor distT="${intAttr(distT)}" distB="${intAttr(distB)}" distL="${intAttr(distL)}" distR="${intAttr(distR)}" simplePos="0" relativeHeight="${intAttr(shape.relativeHeight ?? 251658240)}" behindDoc="${behindDoc}" locked="0" layoutInCell="1" allowOverlap="1">`,
    '<wp:simplePos x="0" y="0"/>',
    position,
    `<wp:extent cx="${intAttr(cx)}" cy="${intAttr(cy)}"/>`,
    '<wp:effectExtent l="0" t="0" r="0" b="0"/>',
    wrap,
    `<wp:docPr id="${docPrId}" name="${escapeXml(docPrName)}"/>`,
    '<wp:cNvGraphicFramePr/>',
    graphic,
    '</wp:anchor>',
    '</w:drawing>',
  ].join('');
}

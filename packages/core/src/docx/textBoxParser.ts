/**
 * Text Box Parser - Parse floating text box containers
 *
 * Text boxes in DOCX are implemented as shapes (wps:wsp) with text body content (wps:txbx).
 * The text body contains w:txbxContent which holds paragraphs and tables like the main document.
 *
 * OOXML Structure:
 * w:drawing
 *   └── wp:inline or wp:anchor
 *       └── a:graphic
 *           └── a:graphicData
 *               └── wps:wsp (shape)
 *                   ├── wps:cNvSpPr (non-visual properties)
 *                   ├── wps:spPr (shape properties)
 *                   │   ├── a:xfrm (transform: position, size)
 *                   │   ├── a:prstGeom (preset geometry - typically "rect" for text boxes)
 *                   │   ├── a:solidFill / a:noFill (fill)
 *                   │   └── a:ln (outline)
 *                   ├── wps:txbx (text box container)
 *                   │   └── w:txbxContent (text content)
 *                   │       ├── w:p (paragraphs)
 *                   │       └── w:tbl (tables)
 *                   └── wps:bodyPr (body properties - margins, text direction, etc.)
 *
 * EMU (English Metric Units): 914400 EMU = 1 inch
 */

import type {
  TextBox,
  Paragraph,
  Table,
  ShapeBlockContent,
  ImageSize,
  ImagePosition,
  ImageWrap,
  Theme,
  RelationshipMap,
  MediaFile,
} from '../types/document';
import type { StyleMap } from './styleParser';
import type { NumberingMap } from './numberingParser';
import {
  getChildElements,
  elementToSelfContainedXml,
  getAttribute,
  getLocalName,
  parseNumericAttribute,
  findByFullName,
  findChildrenByLocalName,
  type XmlElement,
} from './xmlParser';
import {
  collectGroupLeaves,
  mapX,
  mapY,
  readGroupAnchorPosition,
  readXfrm,
  rootGroupFrame,
  type GroupFrame,
} from './groupFrame';
import {
  parseFill,
  parseOutline,
  parseAnchorPosition,
  parseAnchorWrap,
  parsePresetGeometry,
  resolveColorValueToHex,
} from './drawingUtils';
import { emuToPixels } from '../utils/units';

// Re-export emuToPixels for backwards compatibility
export { emuToPixels } from '../utils/units';

// ============================================================================
// CONSTANTS
// ============================================================================

/** Default text box margins in EMUs (0.1 inch) */
const DEFAULT_MARGIN_EMU = 91440;

// ============================================================================
// BODY PROPERTIES PARSING
// ============================================================================

/**
 * Parse text body properties from wps:bodyPr
 * Returns margins/insets for the text box
 */
function parseBodyProperties(bodyPr: XmlElement | null): {
  margins?: TextBox['margins'];
} {
  if (!bodyPr) {
    return {};
  }

  const result: { margins?: TextBox['margins'] } = {};

  // Margins (insets) in EMUs
  const lIns = parseNumericAttribute(bodyPr, null, 'lIns');
  const rIns = parseNumericAttribute(bodyPr, null, 'rIns');
  const tIns = parseNumericAttribute(bodyPr, null, 'tIns');
  const bIns = parseNumericAttribute(bodyPr, null, 'bIns');

  if (lIns !== undefined || rIns !== undefined || tIns !== undefined || bIns !== undefined) {
    result.margins = {
      left: lIns,
      right: rIns,
      top: tIns,
      bottom: bIns,
    };
  }

  return result;
}

// ============================================================================
// CONTENT EXTRACTION
// ============================================================================

/**
 * Extract raw paragraph elements from w:txbxContent
 * Actual parsing happens via document parser to avoid circular dependencies
 */
export function extractTextBoxContentElements(txbxContent: XmlElement | null): {
  paragraphElements: XmlElement[];
  tableElements: XmlElement[];
} {
  if (!txbxContent) {
    return { paragraphElements: [], tableElements: [] };
  }

  const paragraphElements = findChildrenByLocalName(txbxContent, 'p');
  const tableElements = findChildrenByLocalName(txbxContent, 'tbl');

  return { paragraphElements, tableElements };
}

/**
 * Type for the paragraph parser function to avoid circular imports
 */
export type ParagraphParserFn = (
  node: XmlElement,
  styles: StyleMap | null,
  theme: Theme | null,
  numbering: NumberingMap | null,
  rels?: RelationshipMap | null
) => Paragraph;

/**
 * Type for the table parser function to avoid circular imports
 */
export type TableParserFn = (
  node: XmlElement,
  styles: StyleMap | null,
  theme: Theme | null,
  numbering: NumberingMap | null,
  rels?: RelationshipMap | null,
  media?: Map<string, MediaFile>
) => Table;

/**
 * Parse text box content with provided parser functions
 * This avoids circular dependencies by accepting parser functions as parameters
 */
export function parseTextBoxContent(
  txbxContent: XmlElement | null,
  parseParagraph: ParagraphParserFn,
  parseTable: TableParserFn | null,
  styles: StyleMap | null,
  theme: Theme | null,
  numbering: NumberingMap | null,
  rels?: RelationshipMap | null,
  _media?: Map<string, MediaFile>
): ShapeBlockContent[] {
  if (!txbxContent) {
    return [];
  }

  const blocks: ShapeBlockContent[] = [];
  const children = getChildElements(txbxContent);

  for (const child of children) {
    const name = child.name || '';
    const colonIdx = name.indexOf(':');
    const localName = colonIdx >= 0 ? name.substring(colonIdx + 1) : name;

    if (localName === 'p') {
      blocks.push(parseParagraph(child, styles, theme, numbering, rels));
    } else if (localName === 'tbl' && parseTable) {
      // `w:txbxContent` is EG_BlockLevelElts. Skipping tables here painted an
      // empty frame wherever a template put a laid-out panel inside a box.
      blocks.push(parseTable(child, styles, theme, numbering, rels, _media));
    }
  }

  return blocks;
}

// ============================================================================
// TEXT BOX DETECTION
// ============================================================================

/**
 * Check if a drawing element contains a text box
 * Text boxes are shapes with wps:txbx content
 */
export function isTextBoxDrawing(drawingEl: XmlElement): boolean {
  const children = getChildElements(drawingEl);
  const container = children.find((el) => el.name === 'wp:inline' || el.name === 'wp:anchor');

  if (!container) return false;

  const graphic = findByFullName(container, 'a:graphic');
  if (!graphic) return false;

  const graphicData = findByFullName(graphic, 'a:graphicData');
  if (!graphicData) return false;

  // Check for wps:wsp (shape) with text box content
  const wsp = findByFullName(graphicData, 'wps:wsp');
  if (!wsp) return false;

  // Check for text box element
  const txbx = findByFullName(wsp, 'wps:txbx');
  return txbx !== null;
}

/**
 * Check if a wps:wsp element is a text box
 */
export function isShapeTextBox(wsp: XmlElement): boolean {
  const txbx = findByFullName(wsp, 'wps:txbx');
  return txbx !== null;
}

// ============================================================================
// MAIN PARSING FUNCTIONS
// ============================================================================

/**
 * Parse a text box from a w:drawing element
 *
 * This creates a TextBox object with placeholder content.
 * The actual content parsing requires paragraph/table parsers which
 * creates a circular dependency. The document parser should call
 * parseTextBoxContent() separately with the required parsers.
 *
 * @param drawingEl - The w:drawing XML element
 * @returns TextBox object with placeholder content, or null if not a text box
 */
export function parseTextBox(drawingEl: XmlElement): TextBox | null {
  return parseWspDrawing(drawingEl, /* requireTextBox */ true);
}

/**
 * `a:prstGeom` presets Word draws as a single stroke rather than a closed
 * outline. `wps:cNvCnPr` (a connection shape) implies the same thing, so
 * either signal is enough.
 */
const LINE_GEOMETRIES = new Set([
  'line',
  'straightConnector1',
  'bentConnector2',
  'bentConnector3',
  'bentConnector4',
  'bentConnector5',
  'curvedConnector2',
  'curvedConnector3',
  'curvedConnector4',
  'curvedConnector5',
]);

/**
 * Direction of a stroke-only connector, or undefined when the shape is a
 * normal closed geometry. See {@link Shape.lineShape}.
 */
function parseLineShape(wsp: XmlElement, spPr: XmlElement | undefined): 'down' | 'up' | undefined {
  const isConnector = getChildElements(wsp).some((el) => el.name === 'wps:cNvCnPr');
  const prst = getAttribute(findByFullName(spPr ?? null, 'a:prstGeom'), null, 'prst');
  if (!isConnector && !(prst && LINE_GEOMETRIES.has(prst))) return undefined;
  const flipV = getAttribute(findByFullName(spPr ?? null, 'a:xfrm'), null, 'flipV');
  return flipV === '1' || flipV === 'true' ? 'up' : 'down';
}

/**
 * Parse a `wps:wsp` drawing into the {@link TextBox} model.
 *
 * `requireTextBox` false accepts a shape with no `wps:txbx` — a decorative
 * filled rectangle. Those carry no text, so the only thing worth modelling is
 * the frame: size, anchor position, wrap, z-order, fill and outline. See
 * {@link parseFilledShapeAsTextBox}.
 */
function parseWspDrawing(drawingEl: XmlElement, requireTextBox: boolean): TextBox | null {
  const children = getChildElements(drawingEl);

  // Find wp:inline or wp:anchor
  const container = children.find((el) => el.name === 'wp:inline' || el.name === 'wp:anchor');

  if (!container) return null;

  const isAnchor = container.name === 'wp:anchor';

  // Navigate to graphic data
  const graphic = findByFullName(container, 'a:graphic');
  if (!graphic) return null;

  const graphicData = findByFullName(graphic, 'a:graphicData');
  if (!graphicData) return null;

  // Check for wps:wsp (shape)
  const wsp = findByFullName(graphicData, 'wps:wsp');
  if (!wsp) return null;

  // Check for text box
  const txbx = findByFullName(wsp, 'wps:txbx');
  if (!txbx && requireTextBox) return null;

  const wspChildren = getChildElements(wsp);

  // Get shape properties
  const spPr = wspChildren.find((el) => el.name === 'wps:spPr');

  // Get body properties
  const bodyPr = wspChildren.find((el) => el.name === 'wps:bodyPr');

  // Parse size from extent
  const extent = findByFullName(container, 'wp:extent');
  const cx = parseNumericAttribute(extent, null, 'cx') ?? 0;
  const cy = parseNumericAttribute(extent, null, 'cy') ?? 0;
  const size: ImageSize = { width: cx, height: cy };

  // Get document properties
  const docPr = findByFullName(container, 'wp:docPr');
  const id = docPr ? (getAttribute(docPr, null, 'id') ?? undefined) : undefined;

  // Parse fill
  const fill = parseFill(spPr ?? null);

  // Parse outline
  const outline = parseOutline(spPr ?? null);

  // A connector is a stroke between two corners of the extent box, not a
  // rectangle: a footer rule (`prst="line"`, `cy="0"`) outlined as a box
  // paints a full-width border where the file asked for a hairline.
  const lineShape = parseLineShape(wsp, spPr);
  const preset = parsePresetGeometry(spPr);

  // Parse body properties (margins)
  const bodyProps = parseBodyProperties(bodyPr ?? null);

  // Build text box object with placeholder content
  const textBox: TextBox = {
    type: 'textBox',
    size,
    content: [], // Placeholder - will be filled by document parser
  };

  // Add optional properties
  if (id) textBox.id = id;
  if (fill) textBox.fill = fill;
  if (outline) textBox.outline = outline;
  if (lineShape) textBox.lineShape = lineShape;
  if (preset) {
    textBox.geometry = preset.geometry;
    if (preset.cornerAdj !== undefined) textBox.cornerAdj = preset.cornerAdj;
  }
  if (bodyProps.margins) textBox.margins = bodyProps.margins;
  // `wps:bodyPr` carries a dozen attributes and an autofit child that the
  // model has no field for; rebuilding it from margins alone dropped
  // `<a:spAutoFit/>` and the overflow/wrap settings. Same for the `<a:ln>`
  // that says "no outline" — it parses to no `outline`, so the default one
  // came back in its place. See ShapeTextBody.bodyPrXml.
  if (bodyPr) textBox.bodyPrXml = elementToSelfContainedXml(bodyPr);
  if (spPr) {
    const extras = getChildElements(spPr)
      .filter((el) => el.name === 'a:ln' || el.name === 'a:effectLst')
      .map((el) => elementToSelfContainedXml(el))
      .join('');
    if (extras) textBox.spPrExtraXml = extras;
  }

  // Parse position for anchored text boxes
  if (isAnchor) {
    const position = parseAnchorPosition(container);
    if (position) {
      textBox.position = position;
    }

    const wrap = parseAnchorWrap(container);
    if (wrap) {
      textBox.wrap = wrap;
    }

    // Z-order among overlapping anchored objects — Word stacks by this
    // value (cover title text boxes paint over the banner image).
    const relativeHeight = parseNumericAttribute(container, null, 'relativeHeight');
    if (relativeHeight !== null && relativeHeight !== undefined) {
      textBox.relativeHeight = relativeHeight;
    }
  }

  return textBox;
}

/**
 * Does this `w:drawing` hold a decorative filled shape (no text)?
 *
 * Modern templates paint their colour blocks — the navy band behind a
 * section, the accent bar beside a heading — as anchored `wps:wsp` shapes
 * with a solid fill and no `wps:txbx`. Nothing in the pipeline modelled them:
 * `parseImage` returns null for a shape, `isTextBoxDrawing` requires a text
 * box, and the layout engine has no shape block — so they vanished from the
 * page while their text (laid over them in the body) kept its dark-on-dark
 * colour. Treat them as empty text boxes so the existing anchored-frame path
 * paints the fill.
 *
 * A shape with neither fill nor outline draws nothing, so it is not lifted.
 */
export function isFilledShapeDrawing(drawingEl: XmlElement): boolean {
  const children = getChildElements(drawingEl);
  const container = children.find((el) => el.name === 'wp:inline' || el.name === 'wp:anchor');
  if (!container) return false;
  const graphicData = findByFullName(findByFullName(container, 'a:graphic'), 'a:graphicData');
  if (!graphicData) return false;
  const wsp = findByFullName(graphicData, 'wps:wsp');
  if (!wsp) return false;
  // A shape that owns text is a text box; `enrichParagraphTextBoxes` has it.
  if (findByFullName(wsp, 'wps:txbx')) return false;
  const spPr = getChildElements(wsp).find((el) => el.name === 'wps:spPr');
  if (!spPr) return false;
  return !!parseFill(spPr ?? null) || !!parseOutline(spPr ?? null);
}

/**
 * Parse a decorative filled shape (see {@link isFilledShapeDrawing}) as a
 * text box with no content. The caller marks the resulting shape
 * `renderOnly` — the original markup is preserved verbatim for the round
 * trip, so this model exists only to paint.
 */
export function parseFilledShapeAsTextBox(drawingEl: XmlElement): TextBox | null {
  return parseWspDrawing(drawingEl, /* requireTextBox */ false);
}

/**
 * Text boxes and shapes inside a grouped drawing (`wpg:wgp`), each already
 * mapped from the group's child coordinate space onto the page.
 *
 * Word groups a callout panel — a filled rectangle, an accent bar and the
 * text box over them — into one `wpg:wgp`. `isTextBoxDrawing` and
 * `isFilledShapeDrawing` only look for a `wps:wsp` directly under
 * `a:graphicData`, so the whole panel fell out of the model and the page
 * showed nothing where the file drew a bordered box (and the body text that
 * should wrap beside it ran full width instead).
 *
 * The group itself round-trips verbatim as preserved `rawXml`, so every box
 * returned here is paint-only; the caller marks them `renderOnly`. The
 * `wsp` element rides along so the caller can parse `w:txbxContent` with the
 * paragraph parser it owns.
 *
 * @returns One entry per shape, in document order, or `[]` when the drawing
 *   holds no group.
 */
export function parseGroupShapesAsTextBoxes(
  drawingEl: XmlElement
): Array<{ textBox: TextBox; wsp: XmlElement }> {
  const children = getChildElements(drawingEl);
  const container = children.find((el) => el.name === 'wp:inline' || el.name === 'wp:anchor');
  if (!container) return [];
  const graphicData = findByFullName(findByFullName(container, 'a:graphic'), 'a:graphicData');
  if (!graphicData) return [];
  const group = findByFullName(graphicData, 'wpg:wgp');
  if (!group) return [];
  const rootFrame = rootGroupFrame(group);
  if (!rootFrame) return [];

  const isAnchor = container.name === 'wp:anchor';
  const posH = isAnchor ? readGroupAnchorPosition(container, 'positionH') : null;
  const posV = isAnchor ? readGroupAnchorPosition(container, 'positionV') : null;
  const wrap = isAnchor ? parseAnchorWrap(container) : undefined;
  const groupRelativeHeight = isAnchor
    ? parseNumericAttribute(container, null, 'relativeHeight')
    : null;

  const leaves: Array<{ element: XmlElement; frame: GroupFrame }> = [];
  collectGroupLeaves(group, rootFrame, 'wsp', leaves);

  const out: Array<{ textBox: TextBox; wsp: XmlElement }> = [];
  leaves.forEach(({ element: wsp, frame }, index) => {
    const spPr = getChildElements(wsp).find((el) => getLocalName(el.name ?? '') === 'spPr');
    const bodyPr = getChildElements(wsp).find((el) => getLocalName(el.name ?? '') === 'bodyPr');
    const xfrm = readXfrm(findByFullName(spPr ?? null, 'a:xfrm'));
    if (!xfrm?.off || !xfrm.ext) return;

    const fill = parseFill(spPr ?? null);
    const outline = parseOutline(spPr ?? null);
    const hasText = !!findByFullName(wsp, 'wps:txbx');
    // Nothing to paint and nothing to read — skip it rather than stack an
    // invisible click-through frame over the body.
    if (!fill && !outline && !hasText) return;

    const textBox: TextBox = {
      type: 'textBox',
      size: {
        width: Math.round(xfrm.ext.x * frame.scaleX),
        height: Math.round(xfrm.ext.y * frame.scaleY),
      },
      content: [],
    };
    if (fill) textBox.fill = fill;
    if (outline) textBox.outline = outline;
    const lineShape = parseLineShape(wsp, spPr);
    if (lineShape) textBox.lineShape = lineShape;
    const preset = parsePresetGeometry(spPr);
    if (preset) {
      textBox.geometry = preset.geometry;
      if (preset.cornerAdj !== undefined) textBox.cornerAdj = preset.cornerAdj;
    }
    const bodyProps = parseBodyProperties(bodyPr ?? null);
    if (bodyProps.margins) textBox.margins = bodyProps.margins;
    if (bodyPr) textBox.bodyPrXml = elementToSelfContainedXml(bodyPr);

    if (isAnchor) {
      textBox.position = {
        horizontal: {
          relativeTo: (posH?.relativeTo ?? 'column') as ImagePosition['horizontal']['relativeTo'],
          // An aligned group carries `alignment` and NO `posOffset`; emitting
          // both silently pins it back to the origin. Same rule as
          // `deriveGroupPreviewImages`.
          //
          // KNOWN LIMIT: the child's own mapped offset is dropped with it, so
          // every child of an ALIGNED group aligns independently. Correct
          // when the children share the group's box (a panel and the bar
          // across its top — the common callout), wrong for a group whose
          // children sit at different offsets. Fixing it needs the anchor to
          // carry an alignment AND a delta, which the position model has no
          // field for.
          ...(posH?.hasOffset === false && posH.alignment
            ? { alignment: posH.alignment as ImagePosition['horizontal']['alignment'] }
            : { posOffset: Math.round((posH?.offset ?? 0) + mapX(frame, xfrm.off.x)) }),
        },
        vertical: {
          relativeTo: (posV?.relativeTo ?? 'paragraph') as ImagePosition['vertical']['relativeTo'],
          ...(posV?.hasOffset === false && posV.alignment
            ? { alignment: posV.alignment as ImagePosition['vertical']['alignment'] }
            : { posOffset: Math.round((posV?.offset ?? 0) + mapY(frame, xfrm.off.y)) }),
        },
      };
      if (wrap) textBox.wrap = wrap;
      // Children stack in document order within the group; keep them above
      // one another but all at the group's own z-level.
      if (groupRelativeHeight !== null && groupRelativeHeight !== undefined) {
        textBox.relativeHeight = groupRelativeHeight + index;
      }
    }

    out.push({ textBox, wsp });
  });

  return out;
}

/**
 * Parse text box content XML element
 * @param wsp - The wps:wsp element containing the text box
 * @returns The w:txbxContent element or null
 */
export function getTextBoxContentElement(wsp: XmlElement): XmlElement | null {
  const txbx = findByFullName(wsp, 'wps:txbx');
  if (!txbx) return null;

  return findByFullName(txbx, 'w:txbxContent');
}

/**
 * Parse text box from a wps:wsp element directly
 * Useful when you already have the shape element
 */
export function parseTextBoxFromShape(
  wsp: XmlElement,
  size: ImageSize,
  position?: ImagePosition,
  wrap?: ImageWrap
): TextBox | null {
  const txbx = findByFullName(wsp, 'wps:txbx');
  if (!txbx) return null;

  const wspChildren = getChildElements(wsp);

  // Get shape properties
  const spPr = wspChildren.find((el) => el.name === 'wps:spPr');

  // Get body properties
  const bodyPr = wspChildren.find((el) => el.name === 'wps:bodyPr');

  // Get non-visual properties for ID
  const cNvPr = wspChildren.find((el) => el.name === 'wps:cNvPr');
  const id = cNvPr ? (getAttribute(cNvPr, null, 'id') ?? undefined) : undefined;

  // Parse fill
  const fill = parseFill(spPr ?? null);

  // Parse outline
  const outline = parseOutline(spPr ?? null);

  // Parse body properties (margins)
  const bodyProps = parseBodyProperties(bodyPr ?? null);

  // Build text box object
  const textBox: TextBox = {
    type: 'textBox',
    size,
    content: [], // Placeholder
  };

  if (id) textBox.id = id;
  if (fill) textBox.fill = fill;
  if (outline) textBox.outline = outline;
  if (bodyProps.margins) textBox.margins = bodyProps.margins;
  if (position) textBox.position = position;
  if (wrap) textBox.wrap = wrap;

  return textBox;
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Get text box width in pixels
 */
export function getTextBoxWidthPx(textBox: TextBox): number {
  return emuToPixels(textBox.size.width);
}

/**
 * Get text box height in pixels
 */
export function getTextBoxHeightPx(textBox: TextBox): number {
  return emuToPixels(textBox.size.height);
}

/**
 * Get text box dimensions in pixels
 */
export function getTextBoxDimensionsPx(textBox: TextBox): { width: number; height: number } {
  return {
    width: emuToPixels(textBox.size.width),
    height: emuToPixels(textBox.size.height),
  };
}

/**
 * Get text box margins in pixels
 */
export function getTextBoxMarginsPx(textBox: TextBox): {
  top: number;
  bottom: number;
  left: number;
  right: number;
} {
  const margins = textBox.margins;
  return {
    top: emuToPixels(margins?.top ?? DEFAULT_MARGIN_EMU),
    bottom: emuToPixels(margins?.bottom ?? DEFAULT_MARGIN_EMU),
    left: emuToPixels(margins?.left ?? DEFAULT_MARGIN_EMU),
    right: emuToPixels(margins?.right ?? DEFAULT_MARGIN_EMU),
  };
}

/**
 * Check if text box is floating (anchored)
 */
export function isFloatingTextBox(textBox: TextBox): boolean {
  return textBox.position !== undefined || textBox.wrap !== undefined;
}

/**
 * Check if text box has fill
 */
export function hasTextBoxFill(textBox: TextBox): boolean {
  return textBox.fill !== undefined && textBox.fill.type !== 'none';
}

/**
 * Check if text box has outline
 */
export function hasTextBoxOutline(textBox: TextBox): boolean {
  return textBox.outline !== undefined;
}

/**
 * Check if text box has content
 */
export function hasTextBoxContent(textBox: TextBox): boolean {
  return textBox.content.length > 0;
}

/**
 * Get plain text from text box (helper for search/indexing)
 */
export function getTextBoxText(textBox: TextBox): string {
  // This would require getParagraphText utility
  // For now, just join paragraph content
  const parts: string[] = [];

  for (const paragraph of textBox.content) {
    if (paragraph.type !== 'paragraph') continue;
    const runTexts: string[] = [];
    for (const item of paragraph.content) {
      if (item.type === 'run') {
        for (const content of item.content) {
          if (content.type === 'text') {
            runTexts.push(content.text);
          }
        }
      }
    }
    parts.push(runTexts.join(''));
  }

  return parts.join('\n');
}

/**
 * Resolve fill color to CSS color string
 */
export function resolveTextBoxFillColor(textBox: TextBox): string | undefined {
  if (!textBox.fill || textBox.fill.type !== 'solid') return undefined;
  return resolveColorValueToHex(textBox.fill.color);
}

/**
 * Resolve outline color to CSS color string
 */
export function resolveTextBoxOutlineColor(textBox: TextBox): string | undefined {
  if (!textBox.outline?.color) return undefined;
  return resolveColorValueToHex(textBox.outline.color);
}

/**
 * Get outline width in pixels
 */
export function getTextBoxOutlineWidthPx(textBox: TextBox): number {
  if (!textBox.outline?.width) return 0;
  return emuToPixels(textBox.outline.width);
}

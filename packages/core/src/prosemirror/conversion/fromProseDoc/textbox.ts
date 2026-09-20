/**
 * PM textBox → Document Paragraph/Run conversion.
 *
 * Text boxes round-trip as a Shape with a textBody (the inner paragraphs
 * become the shape's content). Two entry points exist: `convertPMTextBoxRun`
 * returns the run for the anchored-inside-paragraph path (Word commonly
 * places anchored shapes inside the following paragraph), and
 * `convertPMTextBox` wraps that run in its own paragraph for the standalone
 * path.
 */

import type { Node as PMNode } from 'prosemirror-model';
import { pixelsToEmu } from '../../../docx/imageParser';
import type {
  Paragraph,
  Run,
  Shape,
  ShapeBlockContent,
  ShapeContent,
} from '../../../types/document';
import { textBoxPositionFromAttrs, textBoxWrapFromAttrs } from '../textBoxAnchors';
import { convertPMParagraph } from './paragraph';
import { convertPMTable } from './tables';

/**
 * Convert a ProseMirror textBox node back to a Paragraph wrapping a ShapeContent run.
 * The text box content becomes a Shape with textBody.
 */
export function convertPMTextBoxRun(node: PMNode): Run {
  const attrs = node.attrs as import('../../extensions/nodes/TextBoxExtension').TextBoxAttrs;

  // Extract the text box's block content. `w:txbxContent` is
  // EG_BlockLevelElts, so a table inside a box round-trips as a table.
  const childBlocks: ShapeBlockContent[] = [];
  node.forEach((child) => {
    if (child.type.name === 'paragraph') {
      childBlocks.push(convertPMParagraph(child));
    } else if (child.type.name === 'table') {
      childBlocks.push(convertPMTable(child));
    }
  });

  // Build shape with text body. shapeType MUST be 'textBox' (not 'rect') so the
  // serializer emits the `<wps:txbx><w:txbxContent>` text body — otherwise a
  // full-repack export writes a plain rect shape with no text, dropping the
  // text box's content (e.g. the cover title/date). The geometry is still a
  // rectangle: the serializer maps shapeType 'textBox' -> prstGeom 'rect'.
  const shape: Shape = {
    type: 'shape',
    shapeType: 'textBox',
    id: attrs.textBoxId || undefined,
    size: {
      width: attrs.width ? pixelsToEmu(attrs.width) : 0,
      height: attrs.height ? pixelsToEmu(attrs.height) : 0,
    },
    textBody: {
      content: childBlocks.length > 0 ? childBlocks : [{ type: 'paragraph', content: [] }],
      margins: {
        top: attrs.marginTop != null ? pixelsToEmu(attrs.marginTop) : undefined,
        bottom: attrs.marginBottom != null ? pixelsToEmu(attrs.marginBottom) : undefined,
        left: attrs.marginLeft != null ? pixelsToEmu(attrs.marginLeft) : undefined,
        right: attrs.marginRight != null ? pixelsToEmu(attrs.marginRight) : undefined,
      },
      bodyPrXml: attrs.bodyPrXml ?? undefined,
    },
    spPrExtraXml: attrs.spPrExtraXml ?? undefined,
  };
  // Canvas-only frame: the source markup is written back verbatim, so the run
  // serializer drops this shape rather than emitting a second copy of it.
  if (attrs.renderOnly) shape.renderOnly = true;
  if (attrs.lineShape === 'down' || attrs.lineShape === 'up') shape.lineShape = attrs.lineShape;
  if (attrs.geometry === 'ellipse' || attrs.geometry === 'roundRect') {
    shape.geometry = attrs.geometry;
    if (typeof attrs.cornerAdj === 'number') shape.cornerAdj = attrs.cornerAdj;
  }

  const position = textBoxPositionFromAttrs(attrs);
  if (position) {
    shape.position = position;
  }

  const wrap = textBoxWrapFromAttrs(attrs);
  if (wrap) {
    shape.wrap = wrap;
  }

  if (attrs.relativeHeight != null) {
    shape.relativeHeight = attrs.relativeHeight;
  }

  // Convert fill color back
  if (attrs.fillColor) {
    shape.fill = {
      type: 'solid',
      color: { rgb: attrs.fillColor.replace('#', '') },
    };
  }

  // Convert outline back
  if (attrs.outlineWidth && attrs.outlineWidth > 0) {
    const cssToOoxmlOutline: Record<string, string> = {
      solid: 'solid',
      dotted: 'dot',
      dashed: 'dash',
    };
    shape.outline = {
      width: pixelsToEmu(attrs.outlineWidth),
      color: attrs.outlineColor ? { rgb: attrs.outlineColor.replace('#', '') } : undefined,
      style: attrs.outlineStyle
        ? (cssToOoxmlOutline[
            attrs.outlineStyle
          ] as import('../../../types/content').ShapeOutline['style']) || 'solid'
        : 'solid',
    };
  }

  // Wrap the shape in a paragraph with a run containing ShapeContent
  const shapeContent: ShapeContent = { type: 'shape', shape };
  return { type: 'run', content: [shapeContent] };
}

export function convertPMTextBox(node: PMNode): Paragraph {
  const hostParaId = (node.attrs as { hostParaId?: string | null }).hostParaId;
  const paragraph: Paragraph = {
    type: 'paragraph',
    content: [convertPMTextBoxRun(node)],
  };
  if (hostParaId) {
    paragraph.paraId = hostParaId;
  }
  return paragraph;
}

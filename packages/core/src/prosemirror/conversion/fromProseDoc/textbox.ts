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
import type { Paragraph, Run, Shape, ShapeContent } from '../../../types/document';
import { textBoxPositionFromAttrs, textBoxWrapFromAttrs } from '../textBoxAnchors';
import { convertPMParagraph } from './paragraph';

/**
 * Convert a ProseMirror textBox node back to a Paragraph wrapping a ShapeContent run.
 * The text box content becomes a Shape with textBody.
 */
export function convertPMTextBoxRun(node: PMNode): Run {
  const attrs = node.attrs as import('../../extensions/nodes/TextBoxExtension').TextBoxAttrs;

  // Extract child paragraphs from the text box content
  const childParagraphs: Paragraph[] = [];
  node.forEach((child) => {
    if (child.type.name === 'paragraph') {
      childParagraphs.push(convertPMParagraph(child));
    }
    // Tables inside text boxes are currently not round-tripped
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
      content: childParagraphs.length > 0 ? childParagraphs : [{ type: 'paragraph', content: [] }],
      margins: {
        top: attrs.marginTop != null ? pixelsToEmu(attrs.marginTop) : undefined,
        bottom: attrs.marginBottom != null ? pixelsToEmu(attrs.marginBottom) : undefined,
        left: attrs.marginLeft != null ? pixelsToEmu(attrs.marginLeft) : undefined,
        right: attrs.marginRight != null ? pixelsToEmu(attrs.marginRight) : undefined,
      },
    },
  };

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
  return {
    type: 'paragraph',
    content: [convertPMTextBoxRun(node)],
  };
}

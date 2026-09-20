/**
 * Document TextBox/anchored shape → PM textBox node + sibling-paragraph
 * extraction (Document → ProseMirror direction).
 *
 * Word stores text boxes as `<w:r><mc:AlternateContent><...><w:txbxContent>`
 * embedded in a host paragraph. This module pulls them out into sibling PM
 * `textBox` nodes (anchored before the host paragraph, in-flow ones after)
 * and converts the shape body into a paragraph-bearing PM node. The host
 * paragraph is dropped if extracting the text boxes leaves it empty.
 */

import type { Node as PMNode } from 'prosemirror-model';
import { schema } from '../../schema';
import type { Paragraph, TextBox, Shape, Theme } from '../../../types/document';
import { emuToPixels } from '../../../docx/imageParser';
import { resolveColor } from '../../../utils/colorResolver';
import type { StyleResolver } from '../../styles';
import { isAnchoredDocxTextBox, textBoxAnchorAttrsFromDocx } from '../textBoxAnchors';
import { convertParagraph } from './paragraph';
import { convertTable } from './tables';

/**
 * Convert a paragraph block to PM nodes, extracting text boxes as sibling nodes.
 * Skips ghost empty paragraphs that only contained text box drawings.
 */
export function convertParagraphWithTextBoxes(
  block: Paragraph,
  styleResolver: StyleResolver | null,
  theme?: Theme | null
): PMNode[] {
  const textBoxes = extractTextBoxesFromParagraph(block);
  const pmParagraph = convertParagraph(block, styleResolver);
  const nodes: PMNode[] = [];
  const { anchored, inFlow } = partitionTextBoxesByAnchor(textBoxes);

  // Whether the host paragraph survives turns on the ANCHORING, not on whether
  // extraction emptied it:
  //
  //   - An ANCHORED box is out of flow. The `w:p` holding it still occupies a
  //     line in Word — that stray empty paragraph under a floating object is
  //     why you cannot delete one without deleting the other. Dropping it
  //     pulled every later block up by a line plus the host's spacing, which on
  //     a cover page built from empty spacer paragraphs walked the artwork up
  //     off the bottom of the page.
  //   - An IN-FLOW box IS the paragraph's content, so it replaces an emptied
  //     host and carries its id.
  //
  // Keeping the host also preserves its `w:pPr`: the export path rebuilds a
  // BARE paragraph from `hostParaId`, losing the style, spacing and alignment.
  // With the host present the box is queued into the following paragraph — this
  // one — by `shouldExportTextBoxInsideFollowingParagraph`, which keeps both,
  // and lets several boxes that shared one host land back on that same host
  // instead of drifting onto different ones.
  const hostEmptiedByExtraction = textBoxes.length > 0 && pmParagraph.content.size === 0;
  const inFlowBoxReplacesHost = hostEmptiedByExtraction && anchored.length === 0;

  for (const tb of anchored) {
    nodes.push(convertTextBox(tb, styleResolver, theme, null));
  }

  if (!inFlowBoxReplacesHost) {
    nodes.push(pmParagraph);
  }

  let hostParaIdTaken = false;
  for (const tb of inFlow) {
    // Only the first box may adopt the id of a host it replaced; a second one
    // would otherwise duplicate that paraId across two paragraphs.
    const hostParaId = inFlowBoxReplacesHost && !hostParaIdTaken ? (block.paraId ?? null) : null;
    if (hostParaId) hostParaIdTaken = true;
    nodes.push(convertTextBox(tb, styleResolver, theme, hostParaId));
  }
  return nodes;
}

function partitionTextBoxesByAnchor(textBoxes: TextBox[]): {
  anchored: TextBox[];
  inFlow: TextBox[];
} {
  const anchored: TextBox[] = [];
  const inFlow: TextBox[] = [];

  for (const textBox of textBoxes) {
    if (isAnchoredDocxTextBox(textBox)) {
      anchored.push(textBox);
    } else {
      inFlow.push(textBox);
    }
  }

  return { anchored, inFlow };
}

/**
 * Extract text boxes from paragraph runs.
 * Text boxes appear as ShapeContent where the shape has textBody,
 * or as DrawingContent that contains a text box instead of an image.
 */
function extractTextBoxesFromParagraph(paragraph: Paragraph): TextBox[] {
  const textBoxes: TextBox[] = [];
  for (const content of paragraph.content) {
    if (content.type === 'run') {
      for (const rc of content.content) {
        if (rc.type === 'shape' && 'shape' in rc) {
          const shape = rc.shape as Shape;
          // A `renderOnly` shape is a decorative filled rectangle: it has a
          // body but no paragraphs, and still has to reach the painter.
          if (shape.textBody && (shape.textBody.content.length > 0 || shape.renderOnly)) {
            // Convert shape with text body to TextBox
            textBoxes.push({
              type: 'textBox',
              id: shape.id,
              size: shape.size,
              position: shape.position,
              wrap: shape.wrap,
              relativeHeight: shape.relativeHeight,
              fill: shape.fill,
              outline: shape.outline,
              content: shape.textBody.content,
              margins: shape.textBody.margins,
              bodyPrXml: shape.textBody.bodyPrXml,
              spPrExtraXml: shape.spPrExtraXml,
              lineShape: shape.lineShape,
              geometry: shape.geometry,
              cornerAdj: shape.cornerAdj,
              renderOnly: shape.renderOnly,
            });
          }
        }
      }
    }
  }
  return textBoxes;
}

/**
 * Convert a TextBox to a ProseMirror textBox node
 */
function convertTextBox(
  textBox: TextBox,
  styleResolver: StyleResolver | null,
  theme?: Theme | null,
  hostParaId?: string | null
): PMNode {
  // A connector declares its real extent, and a rule's is zero on the axis it
  // does not span: `cy="0"` for a horizontal one, `cx="0"` for a vertical one.
  // Falling back on either axis drew the rule as a 200x26 diagonal instead.
  const widthPx = textBox.size?.width
    ? emuToPixels(textBox.size.width)
    : textBox.lineShape
      ? 0
      : 200;
  const heightPx = textBox.size?.height
    ? emuToPixels(textBox.size.height)
    : textBox.lineShape
      ? 0
      : undefined;

  // Convert fill color. Theme-referenced fills (a:schemeClr — e.g. a
  // full-page cover background rectangle filled with tx1/dk1) must
  // resolve through the DOCUMENT theme: dropping them painted the box
  // transparent, leaving its light-colored text invisible on the page.
  let fillColor: string | undefined;
  if (textBox.fill?.color?.rgb) {
    fillColor = `#${textBox.fill.color.rgb}`;
  } else if (textBox.fill?.color?.themeColor) {
    fillColor = resolveColor(textBox.fill.color, theme ?? null);
  }

  // Convert outline
  let outlineWidth: number | undefined;
  let outlineColor: string | undefined;
  let outlineStyle: string | undefined;
  if (textBox.outline && textBox.outline.width) {
    outlineWidth = Math.round((textBox.outline.width / 914400) * 96 * 100) / 100;
    if (textBox.outline.color?.rgb) {
      outlineColor = `#${textBox.outline.color.rgb}`;
    } else if (textBox.outline.color?.themeColor) {
      // Same reason as the fill above: a stroke declared as `a:schemeClr`
      // fell through to the painter's black default.
      outlineColor = resolveColor(textBox.outline.color, theme ?? null);
    }
    outlineStyle = textBox.outline.style || 'solid';
  }

  // Convert margins from EMU to pixels. An inset the source did not declare
  // stays null so a save does not invent one — the painter and `toDOM` supply
  // the visual default (DEFAULT_TEXTBOX_MARGINS).
  const marginTop = textBox.margins?.top != null ? emuToPixels(textBox.margins.top) : null;
  const marginBottom = textBox.margins?.bottom != null ? emuToPixels(textBox.margins.bottom) : null;
  const marginLeft = textBox.margins?.left != null ? emuToPixels(textBox.margins.left) : null;
  const marginRight = textBox.margins?.right != null ? emuToPixels(textBox.margins.right) : null;

  // Convert text box content to PM nodes. `w:txbxContent` is
  // EG_BlockLevelElts, so it can hold tables too — the Iris proposal
  // template's "PROOF POINT" panel is a two-column table inside a box, and
  // dropping it painted an empty frame where the content should be.
  const contentNodes: PMNode[] = [];
  for (const child of textBox.content) {
    contentNodes.push(
      child.type === 'table'
        ? convertTable(child, styleResolver, theme ?? null)
        : convertParagraph(child, styleResolver)
    );
  }

  // Ensure at least one paragraph
  if (contentNodes.length === 0) {
    contentNodes.push(schema.node('paragraph', {}, []));
  }

  return schema.node(
    'textBox',
    {
      width: widthPx,
      height: heightPx,
      textBoxId: textBox.id,
      fillColor,
      outlineWidth,
      outlineColor,
      outlineStyle,
      marginTop,
      marginBottom,
      marginLeft,
      marginRight,
      hostParaId: hostParaId ?? null,
      bodyPrXml: textBox.bodyPrXml ?? null,
      spPrExtraXml: textBox.spPrExtraXml ?? null,
      lineShape: textBox.lineShape ?? null,
      geometry: textBox.geometry ?? null,
      cornerAdj: textBox.cornerAdj ?? null,
      renderOnly: textBox.renderOnly ?? null,
      ...textBoxAnchorAttrsFromDocx(textBox),
    },
    contentNodes
  );
}

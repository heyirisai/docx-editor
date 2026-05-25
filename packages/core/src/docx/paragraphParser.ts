/**
 * Paragraph Parser - Parse paragraphs (w:p) with complete formatting
 *
 * A paragraph is the fundamental block-level element containing text runs,
 * hyperlinks, bookmarks, and fields.
 *
 * OOXML Reference:
 * - Paragraph: w:p
 * - Paragraph properties: w:pPr
 * - Content: runs, hyperlinks, bookmarks, fields
 *
 * This file owns `parseParagraph` (the orchestrator) and re-exports the
 * other public symbols. Property parsing lives in ./paragraphParser/
 * properties.ts, inline-content parsing in ./content.ts, and read-only
 * predicates/text extraction in ./utilities.ts.
 */

import type { Paragraph, Theme, RelationshipMap, MediaFile, NumberFormat } from '../types/document';
import type { StyleMap } from './styleParser';
import type { NumberingMap } from './numberingParser';
import { findChild, getAttribute, type XmlElement } from './xmlParser';
import { parseSectionProperties } from './sectionParser';
import { consolidateParagraphContent } from './runConsolidator';

import { parseParagraphProperties } from './paragraphParser/properties';
import {
  paragraphStartsWithRenderedPageBreak,
  parseParagraphContents,
  parseParagraphPropertyChanges,
} from './paragraphParser/content';

// Public re-exports (preserve historical import surface).
export { parseParagraphProperties } from './paragraphParser/properties';
export {
  getParagraphText,
  isEmptyParagraph,
  isListItem,
  getListLevel,
  hasStyle,
  getTemplateVariable,
} from './paragraphParser/queries';

/**
 * Parse a paragraph element (w:p)
 *
 * @param node - The w:p XML element
 * @param styles - Style map for resolving style references
 * @param theme - Theme for resolving theme colors/fonts
 * @param numbering - Numbering definitions for list info
 * @param rels - Relationship map for resolving hyperlink URLs
 * @param media - Media files map for image data
 * @param options - `inHeaderFooter` skips `<w:lastRenderedPageBreak/>`
 *   detection since headers and footers reflow per page.
 * @returns Parsed Paragraph object
 */
export function parseParagraph(
  node: XmlElement,
  styles: StyleMap | null,
  theme: Theme | null,
  numbering: NumberingMap | null,
  rels: RelationshipMap | null = null,
  media: Map<string, MediaFile> | null = null,
  options?: { inHeaderFooter?: boolean }
): Paragraph {
  const paragraph: Paragraph = {
    type: 'paragraph',
    content: [],
  };

  // Get paragraph ID attributes (Word 2010+ uses these for collaboration)
  const paraId = getAttribute(node, 'w14', 'paraId') ?? getAttribute(node, 'w', 'paraId');
  if (paraId) {
    paragraph.paraId = paraId;
  }

  const textId = getAttribute(node, 'w14', 'textId') ?? getAttribute(node, 'w', 'textId');
  if (textId) {
    paragraph.textId = textId;
  }

  // `<w:lastRenderedPageBreak/>` only makes sense in body flow; headers and
  // footers reflow per page, so detection is skipped there.
  if (!options?.inHeaderFooter && paragraphStartsWithRenderedPageBreak(node)) {
    paragraph.renderedPageBreakBefore = true;
  }

  // Parse paragraph properties (w:pPr)
  const pPr = findChild(node, 'w', 'pPr');
  if (pPr) {
    paragraph.formatting = parseParagraphProperties(pPr, theme, styles ?? undefined);
    paragraph.propertyChanges = parseParagraphPropertyChanges(
      pPr,
      theme,
      styles,
      paragraph.formatting
    );

    // Check for section properties within paragraph (marks end of a section)
    const sectPr = findChild(pPr, 'w', 'sectPr');
    if (sectPr) {
      paragraph.sectionProperties = parseSectionProperties(sectPr, rels);
    }
  }

  // Parse paragraph contents (runs, hyperlinks, bookmarks, fields)
  const rawContent = parseParagraphContents(node, styles, theme, numbering, rels, media);

  // Consolidate consecutive runs with identical formatting
  // This reduces fragmentation (e.g., 252 tiny runs → a few larger runs)
  paragraph.content = consolidateParagraphContent(rawContent);

  // Compute list rendering if this is a list item.
  // numPr can come from inline pPr or from the referenced paragraph style.
  let effectiveNumPr = paragraph.formatting?.numPr;
  if (!effectiveNumPr && paragraph.formatting?.styleId && styles) {
    const style = styles.get(paragraph.formatting.styleId);
    if (style?.pPr?.numPr) {
      effectiveNumPr = style.pPr.numPr;
      // Store it on the paragraph formatting so downstream code sees it
      if (!paragraph.formatting) paragraph.formatting = {};
      paragraph.formatting.numPr = effectiveNumPr;
    }
  }

  if (effectiveNumPr && numbering) {
    const { numId, ilvl = 0 } = effectiveNumPr;
    if (numId !== undefined && numId !== 0) {
      const level = numbering.getLevel(numId, ilvl);
      if (level) {
        // Collect numFmts for levels 0..ilvl so multi-level templates like
        // "%1.%2." can resolve each %N with its own format (e.g., upperRoman
        // parent + decimal child).
        const levelNumFmts: NumberFormat[] = [];
        for (let i = 0; i <= ilvl; i += 1) {
          const parent = numbering.getLevel(numId, i);
          levelNumFmts.push(parent?.numFmt ?? 'decimal');
        }

        const instance = numbering.getInstance(numId);
        const overrideForLevel = instance?.levelOverrides?.find((o) => o.ilvl === ilvl);

        paragraph.listRendering = {
          level: ilvl,
          numId,
          marker: level.lvlText,
          isBullet: level.numFmt === 'bullet',
          numFmt: level.numFmt,
          markerHidden: level.rPr?.hidden || undefined,
          markerFontFamily:
            level.rPr?.fontFamily?.ascii || level.rPr?.fontFamily?.hAnsi || undefined,
          // w:sz is in half-points; convert to points for downstream use
          markerFontSize: level.rPr?.fontSize ? level.rPr.fontSize / 2 : undefined,
          levelNumFmts,
          abstractNumId: instance?.abstractNumId,
          startOverride: overrideForLevel?.startOverride,
        };

        // Apply level's paragraph properties (indentation) as defaults.
        // Per OOXML spec, direct w:ind on the paragraph overrides numbering
        // level indent — only use numbering indent as fallback.
        if (level.pPr) {
          if (!paragraph.formatting) {
            paragraph.formatting = {};
          }
          const directInd = pPr ? findChild(pPr, 'w', 'ind') : null;
          const hasDirectLeft =
            directInd != null &&
            (getAttribute(directInd, 'w', 'left') !== null ||
              getAttribute(directInd, 'w', 'start') !== null);
          // Per ECMA-376 §17.3.1.12 (CT_Ind), `w:firstLine` and `w:hanging`
          // are ST_TwipsMeasure values; a value of `0` is semantically
          // identical to omitting the attribute. Treat both `firstLine="0"`
          // and `hanging="0"` as no-op so the numbering level's indent
          // still applies. A non-numeric value parses to NaN and falls
          // through as an override, preserving prior behavior on
          // malformed input.
          const hasNonZeroDirectAttr = (name: 'firstLine' | 'hanging'): boolean => {
            const raw = directInd ? getAttribute(directInd, 'w', name) : null;
            if (raw === null) return false;
            const value = parseInt(raw, 10);
            return Number.isNaN(value) || value !== 0;
          };
          const hasDirectFirstLineOrHanging =
            directInd != null &&
            (hasNonZeroDirectAttr('firstLine') || hasNonZeroDirectAttr('hanging'));

          if (!hasDirectLeft && level.pPr.indentLeft !== undefined) {
            paragraph.formatting.indentLeft = level.pPr.indentLeft;
          }
          if (!hasDirectFirstLineOrHanging) {
            if (level.pPr.indentFirstLine !== undefined) {
              paragraph.formatting.indentFirstLine = level.pPr.indentFirstLine;
            }
            if (level.pPr.hangingIndent !== undefined) {
              paragraph.formatting.hangingIndent = level.pPr.hangingIndent;
            }
          }
        }
      }
    }
  }

  return paragraph;
}

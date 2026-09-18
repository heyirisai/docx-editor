import type { ParagraphBlock } from './types';

/**
 * Word applies a paragraph's spacing whether or not it has text — ECMA-376
 * §17.3.1.33 makes `w:spacing` a property of the paragraph, with no exemption
 * for an empty one. The suppression Word DOES perform is `w:contextualSpacing`
 * (§17.3.1.9), between consecutive paragraphs of the same style, and that is
 * implemented separately in `applyContextualSpacing` (`layout-engine/index.ts`)
 * before these getters ever run.
 *
 * A second heuristic used to sit here — zero style-inherited spacing on an
 * empty paragraph unless `spacingExplicit` flagged it as direct formatting —
 * ported from deleted renderer code in #402. It has no basis in the spec and
 * broke every template that pushes a cover block down the page with empty
 * spacer paragraphs: the spacers stacked at line height and the cover's
 * anchored art overlapped the text below.
 *
 * NOT the same rule as the HF one in `layout-bridge/headerFooterLayout.ts`,
 * which strips inherited spacing from EVERY paragraph in the header/footer
 * frame (#380) — that is story-scoped and Word really does behave that way.
 */
export function getSpacingBefore(block: ParagraphBlock): number {
  return block.attrs?.spacing?.before ?? 0;
}

export function getSpacingAfter(block: ParagraphBlock): number {
  return block.attrs?.spacing?.after ?? 0;
}

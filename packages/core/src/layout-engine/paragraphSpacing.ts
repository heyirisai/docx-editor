import type { ParagraphBlock } from './types';

/**
 * Word applies a paragraph's spacing whether or not it has text, and the
 * paginator already collapses it against the previous block's `after`. Empty
 * spacer paragraphs are how templates push a cover block down the page, so
 * zeroing style-inherited spacing here stacked them at line height and left
 * anchored art overlapping.
 */
export function getSpacingBefore(block: ParagraphBlock): number {
  return block.attrs?.spacing?.before ?? 0;
}

export function getSpacingAfter(block: ParagraphBlock): number {
  return block.attrs?.spacing?.after ?? 0;
}

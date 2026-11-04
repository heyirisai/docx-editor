/**
 * Shared types and primitive helpers used across toFlowBlocks sub-modules.
 */

import type { Theme } from '../../types/document';

/**
 * Options for the conversion.
 */
export type ToFlowBlocksOptions = {
  /** Default font family. */
  defaultFont?: string;
  /** Default font size in points. */
  defaultSize?: number;
  /** Theme for resolving theme colors. */
  theme?: Theme | null;
  /** Page content height in pixels (pageHeight - marginTop - marginBottom). Images taller than this are scaled down to fit. */
  pageContentHeight?: number;
  /**
   * Document-wide `w:defaultTabStop` (§17.6.13) in twips. Stamped onto each
   * paragraph's attrs so the layout-time list-marker helper can snap body
   * text to the default tab grid when no custom `w:tabs` are defined.
   * Default 720 twips (Word's spec default).
   */
  defaultTabStopTwips?: number;
  /**
   * @internal Allocated by toFlowBlocks() and threaded through table /
   * text-box conversion so list numbering stays continuous across containers.
   * Keyed by abstractNumId when known (ECMA-376 §17.9.18: numIds sharing one
   * abstractNum share counter state); falls back to numId.
   */
  listCounters?: Map<number, number[]>;
  /**
   * @internal Tracks `${numId}:${ilvl}` pairs whose startOverride has already
   * been applied. Per ECMA-376 §17.9.27 the override fires the first time
   * each level of a numId is encountered, so a numId with overrides on
   * multiple ilvls fires each one independently.
   */
  listSeenNumIds?: Set<string>;
};

/**
 * Convert twips to pixels (1 twip = 1/1440 inch, 1 inch = 96 CSS px).
 * No rounding — precision prevents cumulative layout drift across paragraphs.
 */
export function twipsToPixels(twips: number): number {
  return (twips / 1440) * 96;
}

/**
 * Constrain image dimensions to fit within the page content area.
 * Scales proportionally if height exceeds pageContentHeight.
 */
export function constrainImageToPage(
  width: number,
  height: number,
  pageContentHeight: number | undefined
): { width: number; height: number } {
  if (!pageContentHeight || height <= pageContentHeight) {
    return { width, height };
  }
  const scale = pageContentHeight / height;
  return { width: Math.round(width * scale), height: pageContentHeight };
}

let blockIdCounter = 0;
let blockIdsByOwner = new WeakMap<object, Map<string, string[]>>();

export interface BlockIdAllocator {
  /**
   * Return the stable ID for one converted role owned by a ProseMirror node.
   *
   * The occurrence index keeps IDs unique even when callers reuse the same
   * immutable PM node instance in more than one place in a document.
   */
  idFor(owner: object, role: string): string;
}

/**
 * Generate a unique block ID.
 */
export function nextBlockId(): string {
  return `block-${++blockIdCounter}`;
}

/**
 * Create a per-conversion allocator backed by persistent PM-node identity.
 *
 * ProseMirror preserves object identity for unchanged subtrees. Reusing IDs
 * for those nodes lets incremental page rendering distinguish an edited block
 * from unchanged siblings instead of treating every layout pass as a new
 * document.
 */
export function createBlockIdAllocator(): BlockIdAllocator {
  const occurrencesByOwner = new WeakMap<object, Map<string, number>>();

  return {
    idFor(owner, role) {
      let occurrencesByRole = occurrencesByOwner.get(owner);
      if (!occurrencesByRole) {
        occurrencesByRole = new Map();
        occurrencesByOwner.set(owner, occurrencesByRole);
      }
      const occurrence = occurrencesByRole.get(role) ?? 0;
      occurrencesByRole.set(role, occurrence + 1);

      let idsByRole = blockIdsByOwner.get(owner);
      if (!idsByRole) {
        idsByRole = new Map();
        blockIdsByOwner.set(owner, idsByRole);
      }
      let ids = idsByRole.get(role);
      if (!ids) {
        ids = [];
        idsByRole.set(role, ids);
      }
      if (!ids[occurrence]) {
        ids[occurrence] = nextBlockId();
      }
      return ids[occurrence];
    },
  };
}

/**
 * Reset the block ID counter (useful for testing).
 */
export function resetBlockIdCounter(): void {
  blockIdCounter = 0;
  blockIdsByOwner = new WeakMap();
}

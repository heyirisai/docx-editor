/**
 * Page z-order bands.
 *
 * Document content takes its CSS z-index straight from OOXML
 * `relativeHeight` (arbitrary file-supplied values, typically hundreds of
 * millions), so a page-level overlay that must paint above content —
 * `w:pgBorders w:zOrder="front"` — needs a band reserved ABOVE any
 * possible content z. `relativeHeight` is unsigned 32-bit and can exceed
 * the CSS int32 maximum, so content is clamped one below the overlay
 * band: a maximal relativeHeight can neither tie with nor eclipse a
 * front overlay.
 */

/** Front page overlays (page borders) — paint above any document content. */
export const PAGE_OVERLAY_Z = 2147483647;

/** Clamp an OOXML relativeHeight into the document-content band. */
export function contentZIndex(relativeHeight: number): number {
  return Math.min(relativeHeight, PAGE_OVERLAY_Z - 1);
}

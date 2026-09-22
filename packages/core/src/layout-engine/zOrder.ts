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

/** Base of the header/footer band, above every body relativeHeight. */
export const HF_FRONT_Z_BASE = 1_000_000_000;

/**
 * Clamp an OOXML relativeHeight into the document-content band. The ceiling is
 * the HF band, not the overlay band: `relativeHeight` is unsigned 32-bit, so a
 * body float authored above 1e9 would otherwise land on top of the header.
 */
export function contentZIndex(relativeHeight: number): number {
  return Math.min(relativeHeight, HF_FRONT_Z_BASE - 1);
}

/**
 * Header/footer anchored objects live in their own OOXML story, so their
 * `relativeHeight` is not comparable with the body's — a footer date with a
 * lower value was painting under body artwork it is meant to sit on. Keep
 * in-front HF floats above all body content while preserving their order
 * among themselves. `behindDoc` objects still go behind via the -1 path.
 */
export function headerFooterFrontZIndex(relativeHeight: number): number {
  return Math.min(HF_FRONT_Z_BASE + relativeHeight, PAGE_OVERLAY_Z - 1);
}

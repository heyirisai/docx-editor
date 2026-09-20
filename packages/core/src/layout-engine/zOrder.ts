/**
 * Page z-order bands.
 *
 * Measured in Word (probe: a header picture with `behindDoc="0"` and
 * `relativeHeight="9000000"` over a body paragraph, a body text box with
 * `relativeHeight="100"`, and footer text):
 *
 *   - the header picture painted UNDER the body paragraph and UNDER the body
 *     text box, despite a `relativeHeight` five orders of magnitude higher;
 *   - it painted OVER the footer text.
 *
 * So `relativeHeight` is only comparable WITHIN a story. The header/footer
 * story as a whole paints below the body story, and `behindDoc` /
 * `relativeHeight` order objects against the other content of their own story.
 * The bands below encode exactly that, bottom to top:
 *
 *   HF_BEHIND_Z  <  header/footer flow content (auto)  <  HF front floats
 *                <  BODY_CONTENT_Z  <  PAGE_OVERLAY_Z
 *
 * Everything body-side lives inside `.layout-page-content`, which carries
 * {@link BODY_CONTENT_Z} and is therefore its own stacking context — body
 * floats order against each other inside it and cannot reach the HF band.
 */

/** Front page overlays (page borders) — paint above any document content. */
export const PAGE_OVERLAY_Z = 2147483647;

/**
 * `.layout-page-content`. Above the whole header/footer band, below the page
 * overlays. A blanket "HF front floats beat all body content" rule used to sit
 * here instead, and it hid the Hilb cover title under the header artwork.
 */
export const BODY_CONTENT_Z = 1_000_000_000;

/**
 * The band for a header/footer object with `behindDoc="1"` — behind the text
 * of its own story as well as the body's.
 *
 * Safe because `applyPageStyles` gives `.layout-page` both
 * `isolation: isolate` and the page background: inside that stacking context a
 * negative-z descendant paints ABOVE the page's own background and BELOW every
 * positioned child. Moving either property off `.layout-page` breaks this.
 */
export const HF_BEHIND_Z = -1;

/**
 * Clamp an OOXML `relativeHeight` into the document-content band. The value is
 * unsigned 32-bit in the file and is used as a raw CSS z-index, so it is
 * clamped below {@link BODY_CONTENT_Z} — a body float cannot climb out of its
 * own container's stacking context, but a sane value keeps the DOM readable.
 */
export function contentZIndex(relativeHeight: number): number {
  return Math.min(Math.max(relativeHeight, 0), BODY_CONTENT_Z - 1);
}

/**
 * A header/footer anchored object with `behindDoc="0"`: in front of the text of
 * its OWN story — the footer's included, which is what puts a full-bleed header
 * picture over the footer rule and page number (measured on a COMET divider
 * page) — but still below the body. `+ 1` keeps it above header/footer flow
 * content, which has no z-index.
 */
export function headerFooterFrontZIndex(relativeHeight: number): number {
  return 1 + Math.min(Math.max(relativeHeight, 0), BODY_CONTENT_Z - 2);
}

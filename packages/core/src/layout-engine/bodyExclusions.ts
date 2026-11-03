/**
 * Per-page body exclusion bands — the layout option header/footer floats feed.
 * Kept beside `types.ts` (re-exported from the layout-engine entry) so the
 * canonical schema file stays within its size budget.
 *
 * @packageDocumentation
 * @public
 */

import type { Page } from './types';

/**
 * A full-width horizontal band of a page that body content may not occupy, in
 * page coordinates (px from the page's top edge). Produced for header/footer
 * floats that wrap text but leave no room beside them, so the body flows past
 * them the way Word's does.
 *
 * @public
 */
export type BodyExclusionBand = {
  /** Top edge of the band (px from the page top). */
  top: number;
  /** Bottom edge of the band (px from the page top). */
  bottom: number;
};

/**
 * `LayoutOptions.pageBodyExclusions`: the full-width bands a page's body must
 * stay clear of, asked once per page as it is created. A band covering the top
 * edge of the body area starts that page's body below the band; one covering
 * the bottom edge ends the body above it, so content that no longer fits flows
 * to the next page. Bands lying wholly inside the body area are ignored, as are
 * bands that would leave the page (almost) no body at all.
 *
 * @public
 */
export type PageBodyExclusions = (
  page: Readonly<Pick<Page, 'number' | 'size' | 'margins'>>
) => readonly BodyExclusionBand[] | undefined;

/**
 * Header/footer body-margin extension.
 *
 * Word grows the header (or footer) band when its in-flow content is taller
 * than the authored top (or bottom) margin minus the header/footer distance,
 * pushing the body text down (or up). This module owns that computation so the
 * React and Vue adapters share one implementation instead of byte-identical
 * inline copies (the layout pipelines were drifting candidates — see
 * `docx-editor` engine-unification work, issue #696).
 *
 * Two correctness rules live here:
 *
 *  1. The band height is driven by `HeaderFooterContent.flowHeight` (in-flow
 *     content only), NOT `height` / `visualBottom`. A page/margin-anchored
 *     float — e.g. a full-page letterhead anchored in a header — is positioned
 *     on the page and does not push the body in Word. Counting it inflated the
 *     effective top margin past the page on real-world templates, so the
 *     paginator hard-threw "page size and margins yield no content area" and
 *     the document rendered blank (issue #705).
 *
 *  2. A clamp guarantees `top + bottom` never consumes the whole page, so a
 *     pathological in-flow header degrades to a thin content band with a
 *     warning instead of aborting pagination.
 */

import type { FlowBlock, PageMargins, SectionBreakBlock } from '../layout-engine/types';
import type { HeaderFooterContent, SectionHeaderFooterContent } from '../layout-painter/renderPage';

/** Word's default `w:header` / `w:footer` distance (0.5in = 48px). */
const DEFAULT_HF_DISTANCE_PX = 48;

/**
 * Floor on the body content area. Even when header/footer content is absurdly
 * tall, leave at least this much height so pagination produces a page instead
 * of throwing. ~one line at the default body font.
 */
const MIN_CONTENT_HEIGHT_PX = 24;

/** In-flow band height for one HF variant (falls back to total height). */
function bandHeight(hf: HeaderFooterContent | undefined): number {
  if (!hf) return 0;
  return hf.flowHeight ?? hf.height;
}

/** @public */
export interface ExtendMarginsForHeaderFooterInput {
  pageSize: { w: number; h: number };
  /** Body fallback margins. */
  margins: PageMargins;
  /** Final-section margins (last `sectPr`). */
  finalMargins: PageMargins;
  /**
   * Body flow blocks. Each `sectionBreak` block's `margins` is extended IN
   * PLACE so multi-section documents paginate with the same band growth (the
   * layout engine prefers `sectionBreak.margins` over the body fallback).
   */
  bodyBlocks?: FlowBlock[];
  /**
   * Resolved header/footer per section, indexed like `Page.sectionIndex`.
   * Each section's margins are extended by ITS OWN band: a cover that declares
   * no header keeps its authored top margin, and a body section's header still
   * pushes the body it belongs to.
   */
  sections?: Array<SectionHeaderFooterContent | undefined>;
  /** Header variants in play this layout — the fallback when `sections` is absent. */
  headers?: Array<HeaderFooterContent | undefined>;
  /** Footer variants in play this layout. */
  footers?: Array<HeaderFooterContent | undefined>;
  /** Optional diagnostic sink for the clamp (adapters pass `console.warn`). */
  warn?: (message: string) => void;
}

/** @public */
export interface ExtendMarginsForHeaderFooterResult {
  margins: PageMargins;
  finalMargins: PageMargins;
  /**
   * Margins for the FIRST page of the first / last section when `w:titlePg`
   * gives it a different header-footer pair. `undefined` when that page is
   * laid out like the rest of its section.
   */
  firstPageMargins?: PageMargins;
  finalFirstPageMargins?: PageMargins;
}

/**
 * Extend body margins so the body clears the header/footer bands, mirroring
 * Word. Returns new `margins` / `finalMargins`; mutates `sectionBreak.margins`
 * in place. When no extension is needed the original objects are returned
 * unchanged.
 *
 * @public
 */
export function extendMarginsForHeaderFooter(
  input: ExtendMarginsForHeaderFooterInput
): ExtendMarginsForHeaderFooterResult {
  const { pageSize, margins, finalMargins, bodyBlocks, sections, headers, footers, warn } = input;

  const maxMargins = Math.max(0, pageSize.h - MIN_CONTENT_HEIGHT_PX);
  let clamped = false;

  /**
   * The band heights a section pushes its own body with, split by page.
   *
   * `w:titlePg` gives the section's FIRST page its own header/footer pair, so
   * its bands are measured separately: taking `max(header, firstHeader)` for
   * the whole section let a cover's full-page first-page header shrink every
   * page of that section to nothing.
   */
  const bandsFor = (
    index: number
  ): { rest: { header: number; footer: number }; first: { header: number; footer: number } } => {
    if (sections) {
      const sec = sections[index];
      const rest = { header: bandHeight(sec?.header), footer: bandHeight(sec?.footer) };
      // With `w:titlePg` the first page uses its own pair — and a section that
      // declares `titlePg` with no first-page reference has NO header there,
      // which is a zero band, not the default one.
      const first = sec?.titlePg
        ? { header: bandHeight(sec.firstHeader), footer: bandHeight(sec.firstFooter) }
        : rest;
      return { rest, first };
    }
    // Single resolved pair for the whole document — the pre-section behaviour.
    const both = {
      header: Math.max(0, ...(headers ?? []).map(bandHeight)),
      footer: Math.max(0, ...(footers ?? []).map(bandHeight)),
    };
    return { rest: both, first: both };
  };

  /**
   * Grow one section's margins past its own bands. The header band starts at
   * the `w:header` distance and grows downward (§17.6.13), so the body clears
   * `headerDistance + height`. The footer is the mirror image (§17.6.11):
   * `w:footer` is the distance from the page bottom to the footer's BOTTOM
   * edge, so the band occupies `[h - distance - height, h - distance]` and the
   * body must clear `distance + height`. Verified against Word at three
   * different `w:footer` values — see `renderPage`'s `footerBandTop`.
   */
  const grow = (
    m: PageMargins,
    band: { header: number; footer: number },
    allowEmptyContent: boolean
  ): PageMargins => {
    const headerDistance = m.header ?? DEFAULT_HF_DISTANCE_PX;
    const footerDistance = m.footer ?? DEFAULT_HF_DISTANCE_PX;
    const footerTopOffset = footerDistance + band.footer;
    // A section with no band has nothing for the body to clear. Without the
    // `> 0` guard the distance alone pushed the body down, so a full-bleed
    // cover that declares `w:pgMar top="0"` and no header still lost the
    // `w:header` distance off the top of the page.
    const growHeader = band.header > 0 && band.header > m.top - headerDistance;
    const growFooter = band.footer > 0 && footerTopOffset > m.bottom;
    if (!growHeader && !growFooter) return m;

    const out = { ...m };
    if (growHeader) out.top = Math.max(m.top, headerDistance + band.header);
    if (growFooter) out.bottom = Math.max(m.bottom, footerTopOffset);
    // A `w:titlePg` cover page may legitimately have NO content area — Word
    // fills the sheet with the first-page header's artwork and starts the
    // body overleaf. That page is allowed to collapse to zero (bounded by the
    // sheet); every other page keeps a content band so pagination can make
    // progress.
    const limit = allowEmptyContent ? pageSize.h : maxMargins;
    if (out.top + out.bottom > limit) {
      if (!allowEmptyContent) clamped = true;
      // Clamp the footer band first (it sits at the page bottom), then the
      // header band if it alone still overflows.
      out.bottom = Math.max(0, Math.min(out.bottom, limit - out.top));
      if (out.top + out.bottom > limit) {
        out.top = Math.max(0, limit - out.bottom);
      }
    }
    return out;
  };

  const extend = (m: PageMargins, index: number): PageMargins =>
    grow(m, bandsFor(index).rest, false);

  /** `undefined` when the section's first page is laid out like the rest. */
  const extendFirstPage = (m: PageMargins, index: number): PageMargins | undefined => {
    const bands = bandsFor(index);
    if (bands.first.header === bands.rest.header && bands.first.footer === bands.rest.footer) {
      return undefined;
    }
    return grow(m, bands.first, true);
  };

  // Section 0 owns the body fallback margins; each `sectionBreak` owns the
  // section at its own index (the layout engine reads them the same way, see
  // `collectSectionConfigs`); the trailing `sectPr` owns the last one.
  const breaks: SectionBreakBlock[] = [];
  for (const block of bodyBlocks ?? []) {
    if (block.kind === 'sectionBreak') breaks.push(block as SectionBreakBlock);
  }

  const extendedMargins = extend(margins, 0);
  const extendedFinal = extend(finalMargins, breaks.length);
  const firstPageMargins = extendFirstPage(margins, 0);
  const finalFirstPageMargins = extendFirstPage(finalMargins, breaks.length);

  // A section that overrides no margin inherits the previous section's, so the
  // base has to be tracked un-extended — otherwise the inherited value would
  // already carry the WRONG section's band.
  let inheritedBase = margins;
  for (let i = 0; i < breaks.length; i++) {
    const base = breaks[i].margins ?? inheritedBase;
    breaks[i].margins = extend(base, i);
    breaks[i].firstPageMargins = extendFirstPage(base, i);
    inheritedBase = base;
  }

  if (clamped && warn) {
    warn(
      '[layout] header/footer content exceeds page height; clamping margins to ' +
        `preserve a content area. pageHeight=${Math.round(pageSize.h)} ` +
        `top=${Math.round(extendedMargins.top)} bottom=${Math.round(extendedMargins.bottom)}`
    );
  }

  return {
    margins: extendedMargins,
    finalMargins: extendedFinal,
    firstPageMargins,
    finalFirstPageMargins,
  };
}

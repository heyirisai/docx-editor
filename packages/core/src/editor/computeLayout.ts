/**
 * The pure layout COMPUTE pass shared by the React and Vue adapters — issue
 * #696 Tier 2, the clean half of the engine spine.
 *
 * This is the 6-step pass from React's `useLayoutPipeline` minus the DOM paint
 * + scroll/event side-effects (which stay adapter-side, where the framework
 * timing lives): PM doc → flow blocks → measure → header/footer resolve →
 * margin extension → `layoutDocument` (+ two-pass footnote stabilization) →
 * footnote render items. It is pure (no DOM, no refs, no rAF) and returns
 * everything the adapter needs to paint.
 *
 * The one injected seam is `measureBlocks` — each adapter passes its own
 * measurer (React's is caching), same pattern as `measureBlocksWithFloats`.
 * `getHfPmDoc` is the HF-unification seam (prefer the persistent PM doc over
 * re-parsing `HeaderFooter.content`).
 */

import type { EditorState } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';

import {
  collectSectionConfigs,
  layoutDocument,
  type ColumnLayout,
  type FlowBlock,
  type FootnoteContent,
  type Layout,
  type Measure,
  type PageMargins,
} from '../layout-engine';
import {
  toFlowBlocks,
  computePerBlockWidths,
  demoteBlockLikeFloatingTables,
  collectFootnoteRefs,
  convertHeaderFooterToContent,
  convertHeaderFooterPmDocToContent,
  buildFootnoteContentMap,
  buildFootnoteRenderItems,
  stabilizeFootnoteLayout,
  FOOTNOTE_COLUMN_GAP_PX,
  extendMarginsForHeaderFooter,
  twipsToPixels,
  type FloatPageGeometry,
} from '../layout-bridge';
import { createSectionHeaderFooterBodyExclusions } from '../layout-bridge/headerFooterFloatExclusions';
import { extendSectionMarginsForHeaderFooter } from '../layout-bridge/headerFooterMargins';
import { computePerBlockPageGeometry } from '../layout-bridge/sectionGeometry';
import {
  assignPageHeaderFooterRefs,
  resolveSectionHeaderFooters,
  type ResolvedSectionHeaderFooter,
} from '../layout-bridge/sectionHeaderFooters';
import { isAnchoredTopAndBottomImageRun } from '../layout-painter/floatingImageFlow';
import {
  type FootnoteRenderItem,
  type HeaderFooterContent,
  type SectionHeaderFooterContent,
} from '../layout-painter';
import type {
  Document,
  HeaderFooter,
  SectionProperties,
  StyleDefinitions,
  Theme,
  Watermark,
} from '../types/document';

interface PageSizePx {
  w: number;
  h: number;
}

/** Adapter-supplied block measurer (React's is caching). */
export type MeasureBlocksFn = (
  blocks: FlowBlock[],
  contentWidth: number | number[],
  /** One geometry for every block, or one per block (its own section's page). */
  pageGeometry?: FloatPageGeometry | FloatPageGeometry[]
) => Measure[];

export interface ComputeLayoutInputs {
  state: EditorState;
  document: Document | null;
  pageSize: PageSizePx;
  margins: PageMargins;
  columns: ColumnLayout | undefined;
  finalPageSize: PageSizePx;
  finalMargins: PageMargins;
  finalColumns: ColumnLayout | undefined;
  pageGap: number;
  contentWidth: number;
  theme: Theme | null | undefined;
  styles: StyleDefinitions | null | undefined;
  sectionProperties: SectionProperties | null | undefined;
  finalSectionProperties: SectionProperties | null | undefined;
  /** Resolved HF objects for the section (default + first-page). */
  headerContent: HeaderFooter | null | undefined;
  footerContent: HeaderFooter | null | undefined;
  firstPageHeaderContent: HeaderFooter | null | undefined;
  firstPageFooterContent: HeaderFooter | null | undefined;
  measureBlocks: MeasureBlocksFn;
  /** HF unification: the persistent PM doc for an HF, or null to re-parse content. */
  getHfPmDoc: (hf: HeaderFooter) => PMNode | null | undefined;
}

export interface LayoutComputation {
  blocks: FlowBlock[];
  measures: Measure[];
  layout: Layout;
  headerContentForRender: HeaderFooterContent | undefined;
  footerContentForRender: HeaderFooterContent | undefined;
  firstPageHeaderForRender: HeaderFooterContent | undefined;
  firstPageFooterForRender: HeaderFooterContent | undefined;
  hasTitlePg: boolean;
  /**
   * Header/footer stories per section, indexed by `Page.sectionIndex` — pass
   * to `renderPages` as `sectionHeaderFooters` so every page paints its own
   * section's header/footer. Undefined without a document package.
   */
  sectionHeaderFooters: SectionHeaderFooterContent[] | undefined;
  watermark: Watermark | undefined;
  headerDistancePx: number | undefined;
  footerDistancePx: number | undefined;
  pageBorders: SectionProperties['pageBorders'] | undefined;
  footnotesByPage: Map<number, FootnoteRenderItem[]> | undefined;
}

/**
 * Resolve the document-level footnote column layout from `w15:footnoteColumns`.
 *
 * Footnotes paint N-up when any section opts into multiple footnote columns.
 * In a mixed-section document we take the first multi-column section's count
 * and full content width (a documented limitation — per-section footnote
 * column counts are a follow-up); the overwhelmingly common case is a single
 * uniform setting. Returns `{ columns: 1, columnWidth: fallback }` — i.e. the
 * unchanged single-column path — when no section opts in.
 */
function resolveFootnoteColumnLayout(
  document: Document | null,
  fallbackColumnWidth: number
): { columns: number; columnWidth: number } {
  const body = document?.package?.document;
  const sectionProps: Array<SectionProperties | null | undefined> = body
    ? [...(body.sections ?? []).map((s) => s.properties), body.finalSectionProperties]
    : [];
  const fnSection = sectionProps.find((p) => (p?.footnoteColumns ?? 1) > 1);
  if (!fnSection?.footnoteColumns) {
    return { columns: 1, columnWidth: fallbackColumnWidth };
  }

  const columns = fnSection.footnoteColumns;
  // Footnote columns span the section's full content width, independent of the
  // body's w:cols. Mirror the painter's width math so a footnote measured here
  // wraps exactly as it paints.
  const sectionContentWidthPx =
    fnSection.pageWidth != null
      ? twipsToPixels(
          fnSection.pageWidth - (fnSection.marginLeft ?? 1440) - (fnSection.marginRight ?? 1440)
        )
      : fallbackColumnWidth;
  const columnWidth = (sectionContentWidthPx - (columns - 1) * FOOTNOTE_COLUMN_GAP_PX) / columns;
  return { columns, columnWidth: Math.max(1, columnWidth) };
}

/** Flag body-level topAndBottom anchored pictures (`ImageRun.anchoredBand`). */
function markAnchoredTopAndBottomImages(blocks: FlowBlock[]): void {
  for (const block of blocks) {
    if (block.kind !== 'paragraph') continue;
    for (const run of block.runs) {
      if (run.kind === 'image' && isAnchoredTopAndBottomImageRun(run)) run.anchoredBand = true;
    }
  }
}

/**
 * Run the pure layout compute pass (the 6 steps in this file's header), lifted
 * verbatim from `useLayoutPipeline`. The adapter performs the DOM paint
 * (`renderPages`), scroll-restore, `painter:painted`, and state writeback with
 * the returned values.
 */
export function computeLayout(inputs: ComputeLayoutInputs): LayoutComputation {
  const {
    state,
    document,
    pageSize,
    margins,
    columns,
    finalPageSize,
    finalMargins,
    finalColumns,
    pageGap,
    contentWidth,
    theme,
    styles,
    sectionProperties,
    finalSectionProperties,
    headerContent,
    footerContent,
    firstPageHeaderContent,
    firstPageFooterContent,
    measureBlocks,
    getHfPmDoc,
  } = inputs;

  // Step 1: PM doc → flow blocks.
  const pageContentHeight = pageSize.h - margins.top - margins.bottom;
  const blocks = toFlowBlocks(state.doc, { theme, pageContentHeight });

  // Step 2: Measure all blocks (per-section widths; full measure for float context).
  const blockWidths = computePerBlockWidths(
    blocks,
    { pageSize, margins, columns },
    { pageSize: finalPageSize, margins: finalMargins, columns: finalColumns }
  );

  // Step 1.5: Demote full-width "floating" tables to inline. A positioned table
  // that leaves no room for text to wrap beside it (a common full-width contract
  // form table) is block-like in Word/Google Docs — it paginates across pages.
  // Our floating path instead paints it as one overflowing fragment AND makes
  // the next paragraph skip past the whole table height (a wrap zone), stranding
  // it off-page. Clearing `floating` here — before measure and layout — routes
  // it through `layoutTable` (which breaks rows across pages) and suppresses the
  // wrap zone. Purely a layout transform on the ephemeral FlowBlocks; the PM doc
  // and the saved DOCX keep the original floating table.
  demoteBlockLikeFloatingTables(blocks, blockWidths, contentWidth);

  // Body topAndBottom anchors sit at their own offset and reserve a full-width
  // band there (text above stays above), instead of a block line at the run.
  markAnchoredTopAndBottomImages(blocks);

  const measures = measureBlocks(
    blocks,
    blockWidths,
    computePerBlockPageGeometry(
      blocks,
      { pageSize, margins, columns },
      { pageSize: finalPageSize, margins: finalMargins, columns: finalColumns }
    )
  );

  // Step 2.5: Footnote references.
  const footnoteRefs = collectFootnoteRefs(blocks);
  const hasFootnotes = footnoteRefs.length > 0 && !!document?.package?.footnotes;

  // Step 2.75: Header/footer content for rendering (needed before layout to
  // compute effective margins when HF content exceeds available space).
  const hfMetricsHeader = { section: 'header' as const, pageSize, margins };
  const hfMetricsFooter = { section: 'footer' as const, pageSize, margins };
  const defaultTabStopTwips = state.doc.attrs?.defaultTabStopTwips as number | null;
  const hfOptions = { styles, theme, measureBlocks, defaultTabStopTwips };

  // HF unification phase 1: prefer the persistent PM doc when mounted.
  const convertHf = (
    hf: HeaderFooter | null | undefined,
    metrics: typeof hfMetricsHeader | typeof hfMetricsFooter
  ): HeaderFooterContent | undefined => {
    if (!hf) return undefined;
    const pmDoc = getHfPmDoc(hf);
    if (pmDoc) {
      return convertHeaderFooterPmDocToContent(pmDoc, contentWidth, metrics, hfOptions);
    }
    return convertHeaderFooterToContent(hf, contentWidth, metrics, hfOptions);
  };

  const headerContentForRender = convertHf(headerContent, hfMetricsHeader);
  const footerContentForRender = convertHf(footerContent, hfMetricsFooter);
  const hasTitlePg = sectionProperties?.titlePg === true;
  const firstPageHeaderForRender = hasTitlePg
    ? convertHf(firstPageHeaderContent, hfMetricsHeader)
    : undefined;
  const firstPageFooterForRender = hasTitlePg
    ? convertHf(firstPageFooterContent, hfMetricsFooter)
    : undefined;

  // Per-section stories (ECMA-376 §17.10): each page shows the header/footer
  // of its own section, inheriting references a section omits from the one
  // before it. Each part is converted once per page geometry.
  let sectionHeaderFooters: SectionHeaderFooterContent[] | undefined;
  let resolvedSections: ResolvedSectionHeaderFooter[] = [];
  if (document?.package) {
    const resolved = resolveSectionHeaderFooters(
      blocks,
      finalSectionProperties,
      document.package.headers,
      document.package.footers
    );
    const { configs } = collectSectionConfigs(
      blocks,
      { pageSize, margins, columns },
      { pageSize: finalPageSize, margins: finalMargins, columns: finalColumns }
    );
    const converted = new Map<string, HeaderFooterContent | undefined>();
    const partIds = new Map<HeaderFooter, number>();
    const convertForSection = (
      hf: HeaderFooter | null,
      kind: 'header' | 'footer',
      sectionIndex: number
    ): HeaderFooterContent | undefined => {
      if (!hf) return undefined;
      const geometry = configs[Math.min(sectionIndex, configs.length - 1)] ?? {
        pageSize,
        margins,
      };
      if (!partIds.has(hf)) partIds.set(hf, partIds.size);
      const { w, h } = geometry.pageSize;
      const m = geometry.margins;
      const key = `${partIds.get(hf)}|${kind}|${w}|${h}|${m.left}|${m.right}|${m.top}|${m.bottom}`;
      if (!converted.has(key)) {
        const sectionContentWidth = w - m.left - m.right;
        const metrics = { section: kind, pageSize: geometry.pageSize, margins: m };
        const pmDoc = getHfPmDoc(hf);
        converted.set(
          key,
          pmDoc
            ? convertHeaderFooterPmDocToContent(pmDoc, sectionContentWidth, metrics, hfOptions)
            : convertHeaderFooterToContent(hf, sectionContentWidth, metrics, hfOptions)
        );
      }
      return converted.get(key);
    };
    resolvedSections = resolved;
    sectionHeaderFooters = resolved.map((section, i) => ({
      header: convertForSection(section.header, 'header', i),
      footer: convertForSection(section.footer, 'footer', i),
      firstPageHeader: section.titlePg
        ? convertForSection(section.firstHeader, 'header', i)
        : undefined,
      firstPageFooter: section.titlePg
        ? convertForSection(section.firstFooter, 'footer', i)
        : undefined,
      titlePg: section.titlePg,
      headerDistance: section.headerDistance,
      footerDistance: section.footerDistance,
    }));
  }

  // Watermark rides PM state as a doc attr (so it's undoable).
  const watermark = (state.doc.attrs?.watermark as Watermark | null) ?? undefined;

  // Nullish, not truthy: an explicit `w:header="0"` must paint the header at
  // the page top, not fall back to the painter's 0.5in default (#740).
  const headerDistancePx =
    sectionProperties?.headerDistance != null
      ? twipsToPixels(sectionProperties.headerDistance)
      : undefined;
  const footerDistancePx =
    sectionProperties?.footerDistance != null
      ? twipsToPixels(sectionProperties.footerDistance)
      : undefined;

  // Margin extension — push body clear of the header/footer bands (Word grows
  // the band when in-flow content exceeds the authored margin). Shared core
  // helper: uses in-flow `flowHeight` so page/margin-anchored floats (e.g. a
  // letterhead) don't push the body (issue #705), with a content-area clamp;
  // mutates each `sectionBreak.margins` in place.
  const warnMargins = (msg: string) => console.warn(`[computeLayout] ${msg}`);
  const { margins: effectiveMargins, finalMargins: effectiveFinalMargins } = sectionHeaderFooters
    ? extendSectionMarginsForHeaderFooter({
        pageSize,
        margins,
        finalPageSize,
        finalMargins,
        bodyBlocks: blocks,
        sections: sectionHeaderFooters.map((section) => ({
          headers: [section.header, section.firstPageHeader],
          footers: [section.footer, section.firstPageFooter],
          headerDistance: section.headerDistance,
          footerDistance: section.footerDistance,
        })),
        warn: warnMargins,
      })
    : extendMarginsForHeaderFooter({
        pageSize,
        margins,
        finalMargins,
        bodyBlocks: blocks,
        headers: [headerContentForRender, firstPageHeaderForRender],
        footers: [footerContentForRender, firstPageFooterForRender],
        warn: warnMargins,
      });

  // Step 3: Layout onto pages (two-pass when footnotes exist).
  const bodyBreakType = finalSectionProperties?.sectionStart as
    | 'continuous'
    | 'nextPage'
    | 'evenPage'
    | 'oddPage'
    | undefined;
  const layoutOpts = {
    pageSize,
    margins: effectiveMargins,
    finalPageSize,
    finalMargins: effectiveFinalMargins,
    columns: finalColumns,
    bodyBreakType,
    pageGap,
    // Word keeps the body clear of header/footer floats that wrap text and
    // leave no room beside them (a page-wide logo with square wrap pushes every
    // page's body below it). Same stories, distances and page pick as the paint.
    getPageBodyExclusions: createSectionHeaderFooterBodyExclusions(
      sectionHeaderFooters
        ? sectionHeaderFooters.map((section) => ({
            header: section.header,
            footer: section.footer,
            firstPageHeader: section.firstPageHeader,
            firstPageFooter: section.firstPageFooter,
            titlePg: section.titlePg,
            headerDistance: section.headerDistance,
            footerDistance: section.footerDistance,
          }))
        : [
            {
              header: headerContentForRender,
              footer: footerContentForRender,
              firstPageHeader: firstPageHeaderForRender,
              firstPageFooter: firstPageFooterForRender,
              titlePg: hasTitlePg,
              headerDistance: headerDistancePx,
              footerDistance: footerDistancePx,
            },
          ]
    ),
  };

  let layout: Layout;
  let pageFootnoteMap = new Map<number, number[]>();
  let footnoteContentMap = new Map<number, FootnoteContent>();

  if (hasFootnotes) {
    const pass1Layout = layoutDocument(blocks, measures, layoutOpts);
    // w15:footnoteColumns: when a section lays its footnotes out in multiple
    // columns, measure each footnote at the column width (so it wraps the way
    // it will paint) rather than the full content width.
    const { columns: footnoteColumns, columnWidth: footnoteColumnWidth } =
      resolveFootnoteColumnLayout(document, contentWidth);
    footnoteContentMap = buildFootnoteContentMap(
      document!.package.footnotes!,
      footnoteRefs,
      footnoteColumnWidth,
      {
        styles: styles ?? undefined,
        theme: theme ?? null,
        measureBlocks,
        defaultTabStopTwips,
      }
    );
    const stabilized = stabilizeFootnoteLayout({
      blocks,
      measures,
      layoutOpts,
      footnoteRefs,
      footnoteContentMap,
      initialLayout: pass1Layout,
      footnoteColumns,
    });
    layout = stabilized.layout;
    pageFootnoteMap = stabilized.pageFootnoteMap;
  } else {
    layout = layoutDocument(blocks, measures, layoutOpts);
  }
  assignPageHeaderFooterRefs(layout.pages, resolvedSections);

  const footnotesByPage = hasFootnotes
    ? buildFootnoteRenderItems(pageFootnoteMap, footnoteContentMap, document)
    : undefined;

  return {
    blocks,
    measures,
    layout,
    headerContentForRender,
    footerContentForRender,
    firstPageHeaderForRender,
    firstPageFooterForRender,
    hasTitlePg,
    sectionHeaderFooters,
    watermark,
    headerDistancePx,
    footerDistancePx,
    pageBorders: sectionProperties?.pageBorders,
    footnotesByPage: footnotesByPage?.size ? footnotesByPage : undefined,
  };
}

/**
 * Paginator - manages page state during layout
 *
 * Tracks the current page, cursor position, and available space.
 * Creates new pages when content doesn't fit.
 */

import type { Page, PageMargins, Fragment, ColumnLayout } from './types';

/**
 * Current state of a page being laid out.
 */
export type PageState = {
  /** The page being built. */
  page: Page;
  /** Current Y position (cursor) from page top. */
  cursorY: number;
  /** Current column index (0-based). */
  columnIndex: number;
  /** Top margin of content area. */
  topMargin: number;
  /** Bottom boundary of content area (page height - bottom margin). */
  contentBottom: number;
  /** Accumulated trailing spacing (space after previous block). */
  trailingSpacing: number;
};

/**
 * Options for creating a paginator.
 */
export type PaginatorOptions = {
  /** Page size (width, height). */
  pageSize: { w: number; h: number };
  /** Page margins. */
  margins: PageMargins;
  /**
   * Margins for the FIRST page of the section only (`w:titlePg`). Word
   * measures the header band per page, so a cover whose first-page header
   * holds full-page artwork pushes the body off page 1 without touching the
   * rest of the section. Omit when the section's pages are all alike.
   */
  firstPageMargins?: PageMargins;
  /** Index of the section the first page belongs to (defaults to 0). */
  sectionIndex?: number;
  /** Column configuration (optional). */
  columns?: ColumnLayout;
  /** Per-page footnote reserved heights (pageNumber → height in pixels). */
  footnoteReservedHeights?: Map<number, number>;
  /** Callback when a new page is created. */
  onNewPage?: (state: PageState) => void;
};

/**
 * Calculate the width of a single column.
 */
function calculateColumnWidth(
  pageWidth: number,
  leftMargin: number,
  rightMargin: number,
  columns: ColumnLayout
): number {
  const contentWidth = pageWidth - leftMargin - rightMargin;
  const totalGaps = (columns.count - 1) * columns.gap;
  return (contentWidth - totalGaps) / columns.count;
}

/**
 * Creates a paginator for managing page layout state.
 */
export function createPaginator(options: PaginatorOptions) {
  let pageSize = { ...options.pageSize };
  let margins = { ...options.margins };
  let firstPageMargins: PageMargins | undefined = options.firstPageMargins
    ? { ...options.firstPageMargins }
    : undefined;
  let columns: ColumnLayout = options.columns ?? { count: 1, gap: 0 };
  let warnedOversizedFragment = false;

  // Geometry queued by a continuous section break — applied when the next
  // page is naturally created so the current page keeps the old section's
  // size and margins per ECMA-376 §17.6.22.
  let pendingPageSize: { w: number; h: number } | undefined;
  let pendingMargins: PageMargins | undefined;
  let pendingFirstPageMargins: PageMargins | undefined;
  let pendingHasFirstPageMargins = false;
  // The section whose geometry is active. A `continuous` break defers the swap
  // to the next page, and the header follows the geometry, so it is pended the
  // same way.
  let sectionIndex = options.sectionIndex ?? 0;
  let pendingSectionIndex: number | undefined;

  const pages: Page[] = [];
  const states: PageState[] = [];

  function getContentBottom(): number {
    return pageSize.h - margins.bottom;
  }

  function getContentHeight(): number {
    return getContentBottom() - margins.top;
  }

  function getContentWidth(): number {
    return pageSize.w - margins.left - margins.right;
  }

  if (getContentHeight() <= 0) {
    throw new Error(
      'Paginator: page size and margins yield no content area ' +
        `(pageSize=${Math.round(pageSize.w)}x${Math.round(pageSize.h)} ` +
        `margins top=${Math.round(margins.top)} bottom=${Math.round(margins.bottom)} ` +
        `left=${Math.round(margins.left)} right=${Math.round(margins.right)})`
    );
  }

  // Calculate column width
  let columnWidth = calculateColumnWidth(pageSize.w, margins.left, margins.right, columns);

  // Track where column content starts on the current page.
  // Defaults to topMargin but gets updated when columns change mid-page
  // (continuous section break). When advanceColumn moves to the next column,
  // it resets cursorY to this value instead of topMargin.
  let columnRegionTop = margins.top;

  /**
   * Lowest point any column of the CURRENT column region has reached. Single
   * column content that follows a multi-column region must clear the tallest
   * column, not just the one the cursor happens to sit in.
   */
  let columnRegionMaxY = margins.top;

  /**
   * The page's own content bottom, saved while a balanced column region
   * shortens it (see `setColumnRegionBottom`). Restored when the region ends
   * so the rest of the page keeps its full height.
   */
  let columnRegionSavedBottom: number | null = null;

  /**
   * Get X position for a given column index.
   */
  function getColumnX(columnIndex: number): number {
    return margins.left + columnIndex * (columnWidth + columns.gap);
  }

  /**
   * Create a new page and add it to the list.
   */
  function createNewPage(): PageState {
    // Apply any geometry queued by a continuous section break before
    // computing the new page's size / margins.
    if (pendingPageSize || pendingMargins || pendingHasFirstPageMargins) {
      if (pendingPageSize) pageSize = pendingPageSize;
      if (pendingMargins) margins = pendingMargins;
      if (pendingHasFirstPageMargins) firstPageMargins = pendingFirstPageMargins;
      if (pendingSectionIndex !== undefined) sectionIndex = pendingSectionIndex;
      pendingPageSize = undefined;
      pendingMargins = undefined;
      pendingFirstPageMargins = undefined;
      pendingHasFirstPageMargins = false;
      pendingSectionIndex = undefined;
      columnWidth = calculateColumnWidth(pageSize.w, margins.left, margins.right, columns);
    }
    const pageNumber = pages.length + 1;
    const isSectionFirstPage =
      pages.length === 0 || pages[pages.length - 1].sectionIndex !== sectionIndex;
    // `w:titlePg`: the first page of a section carries its own header/footer
    // pair, so its bands — and therefore its margins — differ from the rest.
    const pageMargins = isSectionFirstPage && firstPageMargins ? firstPageMargins : margins;
    const topMargin = pageMargins.top;
    const contentBottom = pageSize.h - pageMargins.bottom;

    // Reduce content bottom by footnote reserved height for this page
    const footnoteHeight = options.footnoteReservedHeights?.get(pageNumber) ?? 0;
    const pageContentBottom = contentBottom - footnoteHeight;

    const page: Page = {
      number: pageNumber,
      fragments: [],
      margins: { ...pageMargins },
      size: { ...pageSize },
      // Which section's geometry this page is under — the painter reads it to
      // pick that section's header and footer.
      sectionIndex,
      // `w:titlePg` is per section, so the painter needs the section's own
      // first page, not the document's.
      isSectionFirstPage,
      footnoteReservedHeight: footnoteHeight > 0 ? footnoteHeight : undefined,
      // Set initial columns; may be overwritten by updateColumns() for continuous section breaks
      columns: columns.count > 1 ? { ...columns } : undefined,
    };

    const state: PageState = {
      page,
      cursorY: topMargin,
      columnIndex: 0,
      topMargin,
      contentBottom: pageContentBottom,
      trailingSpacing: 0,
    };

    pages.push(page);
    states.push(state);

    // Reset column region to page top on new page
    columnRegionTop = topMargin;
    columnRegionMaxY = topMargin;
    columnRegionSavedBottom = null;

    if (options.onNewPage) {
      options.onNewPage(state);
    }

    return state;
  }

  /**
   * Get the current page state, creating one if none exists.
   */
  function getCurrentState(): PageState {
    if (states.length === 0) {
      return createNewPage();
    }
    return states[states.length - 1];
  }

  /**
   * Get available height remaining on the current column.
   */
  function getAvailableHeight(state: PageState): number {
    return state.contentBottom - state.cursorY;
  }

  /**
   * Check if the given height fits in the current column.
   */
  function fits(height: number, state?: PageState): boolean {
    const s = state || getCurrentState();
    return getAvailableHeight(s) >= height;
  }

  /**
   * Advance to the next column, or create a new page if no more columns.
   */
  function advanceColumn(state: PageState): PageState {
    columnRegionMaxY = Math.max(columnRegionMaxY, state.cursorY);
    // Check if there are more columns on this page
    if (state.columnIndex < columns.count - 1) {
      state.columnIndex += 1;
      state.cursorY = columnRegionTop;
      state.trailingSpacing = 0;
      // The balanced height is a target for the columns that come BEFORE the
      // last one. The last column takes whatever is left and is bounded only
      // by the page, so a few pixels of estimation error spill into it
      // instead of breaking the page (Word grows the region the same way).
      if (columnRegionSavedBottom !== null && state.columnIndex === columns.count - 1) {
        state.contentBottom = columnRegionSavedBottom;
      }
      return state;
    }

    // No more columns, create new page
    return createNewPage();
  }

  /**
   * Ensure content of given height can fit.
   * Advances column or creates new page if needed.
   * Returns the state to use for placement.
   */
  function ensureFits(height: number, hasVisibleContent = true): PageState {
    let state = getCurrentState();
    const safeHeight = Number.isFinite(height) && height > 0 ? height : 0;
    // Guards the zero-capacity escape below against a document whose EVERY
    // page has no content area — advance once, then park rather than loop.
    let advancedForZeroCapacity = false;

    while (!fits(safeHeight, state)) {
      // Oversized-fragment guard: re-checked each iteration because page
      // geometry can change between iterations (a continuous section break
      // queues new size/margins that take effect on `createNewPage`). If a
      // single fragment is taller than the content area of an EMPTY page or
      // column, place it with overflow rather than loop forever.
      const columnCapacity = state.contentBottom - state.topMargin;
      if (safeHeight > columnCapacity) {
        // A `w:titlePg` cover whose header fills the sheet leaves NO content
        // area (capacity <= 0). That is the file doing its job, not a
        // pathological fragment: Word parks the first block off the bottom of
        // that sheet — where it is clipped — and starts the body overleaf.
        // Same placement, no warning.
        //
        // That only holds for the EMPTY spacer paragraph such a cover ends
        // with. Parking a fragment that actually paints something would clip
        // real text off the page and lose it, so visible content moves to the
        // next page instead — where there IS a content area.
        if (columnCapacity <= 0 && hasVisibleContent && !advancedForZeroCapacity) {
          advancedForZeroCapacity = true;
          state = advanceColumn(state);
          continue;
        }
        if (!warnedOversizedFragment && columnCapacity > 0) {
          warnedOversizedFragment = true;
          console.warn(
            `Paginator: fragment height ${safeHeight.toFixed(0)}px exceeds page content height ${columnCapacity.toFixed(0)}px; placing with overflow.`
          );
        }
        if (state.cursorY !== state.topMargin) {
          state = advanceColumn(state);
        }
        return state;
      }
      state = advanceColumn(state);
    }

    return state;
  }

  /**
   * Add a fragment to the current page at the cursor position.
   * Updates cursor position after placement.
   */
  function addFragment(
    fragment: Fragment,
    height: number,
    spaceBefore: number = 0,
    spaceAfter: number = 0,
    /**
     * Whether this fragment paints anything. Only consulted on a page with no
     * content area at all (see `ensureFits`); defaults to true so a caller
     * that cannot tell never risks clipping real content away.
     */
    hasVisibleContent: boolean = true
  ): { state: PageState; x: number; y: number } {
    // Word collapses adjacent paragraphs' spaceAfter / next.spaceBefore to
    // the larger of the two (CSS-style margin-collapse), not the sum.
    const effectiveSpaceBefore = Math.max(spaceBefore, getCurrentState().trailingSpacing);
    const totalHeight = effectiveSpaceBefore + height;

    // Ensure we have space
    const state = ensureFits(totalHeight, hasVisibleContent);

    // Word 2013+ (compatibilityMode ≥ 15) honors an explicit w:before on the
    // first paragraph of a page/column — it's not auto-suppressed. trailingSpacing
    // is already reset to 0 when a new page/column starts (so we don't carry
    // spacing across page breaks), so applying effectiveSpaceBefore here is safe.
    const actualSpaceBefore = effectiveSpaceBefore;

    // Calculate position
    const x = getColumnX(state.columnIndex);
    const y = state.cursorY + actualSpaceBefore;

    // Position the fragment
    fragment.x = x;
    fragment.y = y;

    // Add to page
    state.page.fragments.push(fragment);

    // Update cursor
    state.cursorY = y + height;
    state.trailingSpacing = spaceAfter;

    return { state, x, y };
  }

  /**
   * Re-decide whether `state`'s page opens a section, and re-apply that
   * page's margins.
   *
   * An `evenPage` / `oddPage` break can make a blank parity page and then
   * open the section on the NEXT one, so the page that `createNewPage`
   * stamped is not always the one that ends up first. Margins follow the
   * stamp (`w:titlePg`), so flipping the flag alone left the opening page
   * laid out against the wrong geometry.
   */
  function restampSectionFirstPage(state: PageState, isFirst: boolean): void {
    state.page.isSectionFirstPage = isFirst;
    const pageMargins = isFirst && firstPageMargins ? firstPageMargins : margins;
    if (
      state.page.margins.top === pageMargins.top &&
      state.page.margins.bottom === pageMargins.bottom
    ) {
      return;
    }
    const footnoteHeight = options.footnoteReservedHeights?.get(state.page.number) ?? 0;
    const atTop = state.cursorY === state.topMargin;
    state.page.margins = { ...pageMargins };
    state.topMargin = pageMargins.top;
    state.contentBottom = pageSize.h - pageMargins.bottom - footnoteHeight;
    // Only rewind the cursor on a page nothing has been placed on yet.
    if (atTop) state.cursorY = pageMargins.top;
    columnRegionTop = state.cursorY;
    columnRegionMaxY = state.cursorY;
  }

  /**
   * Force a page break - move to a new page.
   *
   * Idempotent when the current page is empty: a section break followed by
   * `pageBreakBefore` on the next paragraph (or any other chain of forced
   * breaks) collapses to a single break instead of leaving a phantom page.
   */
  function forcePageBreak(): PageState {
    if (states.length > 0) {
      const current = states[states.length - 1];
      if (current.page.fragments.length === 0 && current.cursorY === current.topMargin) {
        return current;
      }
    }
    return createNewPage();
  }

  /**
   * Force a column break - move to next column or new page.
   */
  function forceColumnBreak(): PageState {
    const state = getCurrentState();
    return advanceColumn(state);
  }

  /**
   * Update column configuration mid-document (for section breaks).
   * Recalculates column width based on current page/margin dimensions.
   * Sets columnRegionTop to the current cursor position so that
   * column advancement stays below existing content (for continuous breaks).
   */
  function updateColumns(newColumns: ColumnLayout): void {
    // Close the outgoing region before switching: whatever follows starts
    // below the TALLEST column and gets the page's full height back.
    const state = getCurrentState();
    if (columns.count > 1) {
      state.cursorY = Math.max(state.cursorY, columnRegionMaxY);
      state.trailingSpacing = 0;
      if (columnRegionSavedBottom !== null) {
        state.contentBottom = columnRegionSavedBottom;
        columnRegionSavedBottom = null;
      }
    }

    columns = newColumns;
    columnWidth = calculateColumnWidth(pageSize.w, margins.left, margins.right, columns);

    // Update current page's column info for rendering
    state.page.columns = columns.count > 1 ? { ...columns } : undefined;

    // Set column region top to current cursor position.
    // This ensures that when advancing columns, new columns start
    // at the same Y as where the multi-column content began (not page top).
    columnRegionTop = state.cursorY;
    columnRegionMaxY = state.cursorY;

    // Reset to column 0 for the new column layout
    state.columnIndex = 0;
  }

  /**
   * Shorten the current column region so its content balances across the
   * columns (Word balances a continuous multi-column section rather than
   * filling column 1 to the page bottom). The page's own bottom is restored
   * when the region ends — see `updateColumns`.
   */
  function setColumnRegionBottom(bottom: number): void {
    const state = getCurrentState();
    if (columnRegionSavedBottom === null) columnRegionSavedBottom = state.contentBottom;
    state.contentBottom = bottom;
  }

  /**
   * Update page geometry for pages created after a section break.
   *
   * `applyImmediately = true` (default) swaps the active geometry so the
   * NEXT page created by `forcePageBreak`/`createNewPage` and the column
   * width on the current page both reflect the new section. Used by
   * `nextPage` / `evenPage` / `oddPage` breaks where the next content
   * starts on a fresh page anyway.
   *
   * `applyImmediately = false` defers the swap until `createNewPage`
   * actually fires. Used by `continuous` breaks: ECMA-376 §17.6.22 keeps
   * the current page in the OLD section's geometry but applies the new
   * section's page size / margins to the NEXT naturally-created page.
   * The current page's `columnWidth` is left intact under the old
   * geometry; columns for the new section are still applied via
   * `updateColumns`.
   */
  function updatePageLayout(
    newPageSize?: { w: number; h: number },
    newMargins?: PageMargins,
    applyImmediately = true,
    newSectionIndex?: number,
    newFirstPageMargins?: PageMargins
  ): void {
    if (!applyImmediately) {
      pendingPageSize = newPageSize ? { ...newPageSize } : pendingPageSize;
      pendingMargins = newMargins ? { ...newMargins } : pendingMargins;
      // Always queued, including `undefined`: the incoming section may have
      // no title page where the outgoing one did.
      pendingFirstPageMargins = newFirstPageMargins ? { ...newFirstPageMargins } : undefined;
      pendingHasFirstPageMargins = true;
      pendingSectionIndex = newSectionIndex ?? pendingSectionIndex;
      return;
    }
    if (newPageSize) {
      pageSize = { ...newPageSize };
    }
    if (newMargins) {
      margins = { ...newMargins };
    }
    firstPageMargins = newFirstPageMargins ? { ...newFirstPageMargins } : undefined;
    if (newSectionIndex !== undefined) {
      sectionIndex = newSectionIndex;
    }
    if (getContentHeight() <= 0) {
      throw new Error('Paginator: section page size and margins yield no content area');
    }
    columnWidth = calculateColumnWidth(pageSize.w, margins.left, margins.right, columns);
    // A pending swap is now superseded by this immediate swap.
    pendingPageSize = undefined;
    pendingMargins = undefined;
    pendingFirstPageMargins = undefined;
    pendingHasFirstPageMargins = false;
    pendingSectionIndex = undefined;
  }

  return {
    /** All pages created so far. */
    pages,
    /** All page states. */
    states,
    /** Column width in pixels (use getColumnWidth() for current value after updates). */
    get columnWidth() {
      return columnWidth;
    },
    /** Get current column layout (returns copy to prevent external mutation). */
    get columns() {
      return { ...columns };
    },
    /** Get current state. */
    getCurrentState,
    /** Get available height in current column. */
    getAvailableHeight: () => getAvailableHeight(getCurrentState()),
    /** Content height of a whole page under the active geometry. */
    getContentHeight,
    /** Get content width for the active section. */
    getContentWidth,
    /** Check if height fits in current column. */
    fits: (height: number) => fits(height),
    /** Ensure height fits, advancing if needed. */
    ensureFits,
    /** Add a fragment to current page. */
    addFragment,
    /** Force a page break. */
    forcePageBreak,
    /** Force a column break. */
    forceColumnBreak,
    /** Get X position for column. */
    getColumnX,
    /** Update column layout (for section breaks). */
    updateColumns,
    /** Shorten the current column region for balancing. */
    setColumnRegionBottom,
    /** Update page size/margins for subsequent pages. */
    updatePageLayout,
    restampSectionFirstPage,
  };
}

export type Paginator = ReturnType<typeof createPaginator>;

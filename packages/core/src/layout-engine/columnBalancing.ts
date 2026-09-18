import type { FlowBlock, Measure } from './types';
import type { Paginator } from './paginator';
import { getSpacingAfter, getSpacingBefore } from './paragraphSpacing';

function getBalancedTextSectionHeight(
  blocks: FlowBlock[],
  measures: Measure[],
  start: number,
  end: number
): number | null {
  let totalHeight = 0;
  let hasText = false;
  // Mirror exactly what the paginator consumes per paragraph: the line heights
  // it places, plus ONE collapsed gap (`addFragment` applies
  // `max(spaceBefore, trailingSpacing)`).
  //
  // `ParagraphMeasure.totalHeight` is the wrong input here on two counts: it
  // already folds spacing.before/after in, and that spacing is the value from
  // BEFORE `applyContextualSpacing` zeroed it. Adding `getSpacingBefore/After`
  // on top counted each gap two or three times and inflated a section of a
  // dozen short paragraphs by ~30% — the balance point landed a third of the
  // way down and the second column came out nearly empty.
  let trailingSpacing = 0;

  for (let i = start; i < end; i++) {
    const block = blocks[i];
    const measure = measures[i];

    if (block.kind === 'paragraph' && measure.kind === 'paragraph') {
      const linesHeight = measure.lines.reduce(
        (sum, line) => sum + line.lineHeight + (line.floatSkipBefore ?? 0),
        0
      );
      totalHeight += Math.max(getSpacingBefore(block), trailingSpacing) + linesHeight;
      trailingSpacing = getSpacingAfter(block);
      hasText = hasText || measure.lines.length > 0;
      continue;
    }

    if (block.kind === 'sectionBreak') {
      continue;
    }

    return null;
  }

  return hasText ? totalHeight : null;
}

function balanceCurrentColumnRegion(paginator: Paginator, totalContentHeight: number): void {
  const columns = paginator.columns;
  if (columns.count <= 1 || !Number.isFinite(totalContentHeight) || totalContentHeight <= 0) {
    return;
  }

  const state = paginator.getCurrentState();
  const columnRegionTop = state.cursorY;
  const maxRegionHeight = state.contentBottom - columnRegionTop;
  if (maxRegionHeight <= 0 || totalContentHeight > maxRegionHeight * columns.count) {
    return;
  }

  const balancedHeight = Math.ceil(totalContentHeight / columns.count);
  if (balancedHeight <= 0 || balancedHeight >= maxRegionHeight) {
    return;
  }

  // Through the paginator so the page's own bottom is restored when the
  // region ends — a mid-document column section is followed by full-width
  // content that must get the rest of the page back.
  paginator.setColumnRegionBottom(columnRegionTop + balancedHeight);
}

/**
 * Balance a `continuous` multi-column section across its columns.
 *
 * Word lays a continuous multi-column section out as a balanced block: the
 * content is split evenly between the columns rather than filling column one
 * to the page bottom. Without this a short two-column section stacked into
 * column one and the page ran long. `start`/`end` bound the section's blocks
 * (the next section break, or the end of the document).
 */
export function balanceContinuousTextColumns({
  blocks,
  measures,
  paginator,
  start,
  end,
}: {
  blocks: FlowBlock[];
  measures: Measure[];
  paginator: Paginator;
  start: number;
  end: number;
}): void {
  const balancedHeight = getBalancedTextSectionHeight(blocks, measures, start, end);
  if (balancedHeight !== null) {
    balanceCurrentColumnRegion(paginator, balancedHeight);
  }
}

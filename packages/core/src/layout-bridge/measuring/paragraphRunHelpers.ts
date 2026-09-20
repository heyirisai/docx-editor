/**
 * Run predicates and the small text helpers paragraph measurement leans on.
 *
 * Split out of `measureParagraph.ts` purely for size — it had grown past the
 * 1100-line lint bound. Everything here is pure: no line state, no floats.
 */

import type {
  Run,
  TextRun,
  TabRun,
  ImageRun,
  LineBreakRun,
  FieldRun,
} from '../../layout-engine/types';
import { measureTextWidth, type FontStyle } from './measureContainer';
import { wrapsAroundText } from '../../docx/wrapTypes';
import { formatWordDate } from '../../docx/dateFormat';

/** Default font family used when a run declares none. */
const DEFAULT_FONT_FAMILY = 'Calibri';
/** Default font size in points used when a run declares none. */
const DEFAULT_FONT_SIZE = 11;

/**
 * Extract FontStyle from a text run for measurement
 */
export function runToFontStyle(run: TextRun | TabRun): FontStyle {
  return {
    fontFamily: run.fontFamily ?? DEFAULT_FONT_FAMILY,
    fontSize: run.fontSize ?? DEFAULT_FONT_SIZE,
    bold: run.bold,
    italic: run.italic,
    letterSpacing: run.letterSpacing,
    allCaps: run.allCaps,
  };
}

/**
 * The text a field run will actually paint. DATE/TIME are recomputed on open
 * and rendered through their `\@` picture, so measuring the value cached in
 * the file sizes the line for the wrong string — which is how a cover date
 * overflows its text box even after the painter was taught the picture.
 * Mirrors `renderFieldRun`.
 */
export function fieldMeasurementText(run: FieldRun): string {
  if (run.fieldType === 'DATE' || run.fieldType === 'TIME') {
    const now = new Date();
    if (run.fieldFormat) return formatWordDate(now, run.fieldFormat);
    return run.fieldType === 'DATE' ? now.toLocaleDateString() : now.toLocaleTimeString();
  }
  return run.fallback || '1';
}

/**
 * Check if a run is a text run
 */
export function isTextRun(run: Run): run is TextRun {
  return run.kind === 'text';
}

/**
 * Check if a run is a tab run
 */
export function isTabRun(run: Run): run is TabRun {
  return run.kind === 'tab';
}

/**
 * Check if a run is an image run
 */
export function isImageRun(run: Run): run is ImageRun {
  return run.kind === 'image';
}

/**
 * Check if a run is a line break run
 */
export function isLineBreakRun(run: Run): run is LineBreakRun {
  return run.kind === 'lineBreak';
}

/**
 * Check if a run is a field run
 */
export function isFieldRun(run: Run): run is FieldRun {
  return run.kind === 'field';
}

/**
 * Check if text run is empty (only whitespace or no text)
 */
export function isEmptyTextRun(run: TextRun): boolean {
  return !run.text || run.text.replace(/\u00a0/g, ' ').trim().length === 0;
}

/**
 * Sum the inline pixel widths of runs after a tab, up to (but not including)
 * the next tab or line break. Measured per-run so widths reserved match what
 * the painter draws even when trailing runs use different fonts/sizes.
 */
export function measureInlineWidthAfterTab(runs: Run[], tabIndex: number): number {
  let width = 0;
  for (let i = tabIndex + 1; i < runs.length; i++) {
    const next = runs[i];
    if (isTabRun(next) || isLineBreakRun(next)) break;
    if (isTextRun(next)) {
      width += measureTextWidth(next.text || '', runToFontStyle(next));
    } else if (isFieldRun(next)) {
      const style: FontStyle = {
        fontFamily: next.fontFamily ?? DEFAULT_FONT_FAMILY,
        fontSize: next.fontSize ?? DEFAULT_FONT_SIZE,
        bold: next.bold,
        italic: next.italic,
      };
      width += measureTextWidth(fieldMeasurementText(next), style);
    } else if (isImageRun(next)) {
      // Floating / anchored images are positioned at the page level and
      // contribute no inline width — counting them (e.g. a footer's decorative
      // full-width wave images that trail the text) would massively inflate the
      // reserved width. Mirrors the painter's measureFollowingContentWidth.
      const isFloating = next.displayMode === 'float' || wrapsAroundText(next.wrapType);
      if (!(next.position && isFloating)) {
        width += next.width || 0;
      }
    }
  }
  return width;
}

/**
 * Find word break points in text
 * Returns array of indices where words end (after space/punctuation)
 */
export function findWordBreaks(text: string): number[] {
  const breaks: number[] = [];

  for (let i = 0; i < text.length; i++) {
    const char = text[i];
    // Break after space or certain punctuation
    if (char === ' ' || char === '-' || char === '\t') {
      breaks.push(i + 1);
    }
  }

  return breaks;
}

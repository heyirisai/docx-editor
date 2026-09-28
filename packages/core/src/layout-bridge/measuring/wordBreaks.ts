/**
 * Line-break opportunities for the paragraph measurer.
 *
 * The breaker wraps after a space, hyphen or tab. A run boundary is NOT a
 * break opportunity: Word splits formatting mid-word all the time
 * (`<w:t> u</w:t>` + `<w:t>tilized?</w:t>`, spell-check and revision-id
 * splits) and still wraps the word as one unit.
 */

import type { Run, TextRun } from '../../layout-engine/types';
import { measureTextWidth, type FontStyle } from './measureContainer';

/** Characters after which the line breaker may wrap. */
export function isBreakChar(char: string | undefined): boolean {
  return char === ' ' || char === '-' || char === '\t';
}

/**
 * Find word break points in text.
 * Returns the indices where words end (just after a space/hyphen/tab).
 */
export function findWordBreaks(text: string): number[] {
  const breaks: number[] = [];
  for (let i = 0; i < text.length; i++) {
    if (isBreakChar(text[i])) breaks.push(i + 1);
  }
  return breaks;
}

/**
 * Width of the word fragment that continues past the end of `runs[runIndex]`
 * into the following runs — i.e. the text up to (and including) the next
 * break opportunity. Only text runs continue a word; tabs, breaks, fields and
 * images end it.
 */
export function measureWordContinuation(
  runs: Run[],
  runIndex: number,
  styleOf: (run: TextRun) => FontStyle
): number {
  let width = 0;
  for (let i = runIndex + 1; i < runs.length; i++) {
    const next = runs[i];
    if (!next || next.kind !== 'text') break;
    const text = next.text ?? '';
    if (!text) continue;
    let end = 0;
    while (end < text.length && !isBreakChar(text[end])) end++;
    // The break char itself (a trailing space or hyphen) stays with the word.
    const fragmentEnd = end < text.length ? end + 1 : end;
    width += measureTextWidth(text.slice(0, fragmentEnd), styleOf(next));
    if (end < text.length) break;
  }
  return width;
}

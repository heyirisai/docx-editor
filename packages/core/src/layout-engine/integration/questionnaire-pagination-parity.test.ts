/**
 * Word pagination parity against a real questionnaire (Google Docs authored,
 * A4, Arial throughout, no explicit page breaks).
 *
 * The fixture is NOT in the repo. Point `DOCX_FIXTURE_PAGINATION` at it to run
 * it; otherwise the suite skips loudly.
 *
 * Word's truth for this document:
 *  - "Table of Contents" (body block 54) is the FIRST block on page 2;
 *  - Heading 1 "Introduction" starts on page 3.
 *
 * Three engine rules decide this, and each was wrong before the fix that
 * added this suite:
 *  1. Arial single spacing is 1.1499 × size (GDI external leading), so the
 *     51 empty 8.5pt spacer lines, the 18pt title and the TOC entries all
 *     measure ~3% taller than with the bare usWin sum;
 *  2. the header (11 empty 12pt lines) pushes the body top to
 *     max(top margin, header distance + header height) — the paragraph
 *     that only carries the anchored logo must be sized from its 12pt mark;
 *  3. the footer (a floating "Internal" text box + 4 empty 12pt lines)
 *     pushes the body bottom to footer distance + footer height, because
 *     Word anchors the footer's BOTTOM at the `w:footer` distance.
 *
 * The full compute pass (`computeLayout`) runs headless: a canvas stub
 * measures text at 0.5em per character (wrapping is irrelevant to the two
 * assertions — every block that matters is a single line) and line heights
 * come from the OS/2 ratio table, as in the browser.
 */

import { describe, test, expect, beforeAll, afterAll } from 'bun:test';
import fs from 'node:fs';
import { installCanvasDocumentStub } from './helpers';
import { layoutDocxHeadless, parseDocxHeadless } from './headlessDocxLayout';

const FIXTURE = process.env.DOCX_FIXTURE_PAGINATION;
const available = !!FIXTURE && fs.existsSync(FIXTURE);
if (!available) {
  console.warn(
    '[questionnaire-pagination-parity] SKIPPED — set DOCX_FIXTURE_PAGINATION to the path of ' +
      'the questionnaire fixture to run the Word pagination parity check.'
  );
}
const run = available ? test : test.skip;

type Fragment = { blockId: unknown };
type ParagraphLike = { kind: string; id: unknown; runs?: Array<{ text?: string }> };

async function layoutFixture() {
  const doc = await parseDocxHeadless(new Uint8Array(fs.readFileSync(FIXTURE!)));
  return layoutDocxHeadless(doc);
}

function paragraphText(b: ParagraphLike): string {
  return (b.runs ?? []).map((r) => r.text ?? '').join('');
}

describe('questionnaire fixture paginates like Word', () => {
  let restoreDocument: () => void = () => {};
  beforeAll(() => {
    if (available) restoreDocument = installCanvasDocumentStub();
  });
  afterAll(() => restoreDocument());

  run(
    '"Table of Contents" opens page 2 and "Introduction" is on page 3',
    async () => {
      const res = await layoutFixture();
      const blocks = res.blocks as unknown as ParagraphLike[];
      const pages = res.layout.pages as unknown as Array<{ fragments: Fragment[] }>;

      const pageOf = (id: unknown): number[] =>
        pages.flatMap((p, i) => (p.fragments.some((f) => f.blockId === id) ? [i + 1] : []));

      const toc = blocks.find(
        (b) => b.kind === 'paragraph' && paragraphText(b) === 'Table of Contents'
      );
      const intro = blocks.find(
        (b) => b.kind === 'paragraph' && paragraphText(b) === 'Introduction'
      );
      expect(toc).toBeDefined();
      expect(intro).toBeDefined();

      // Word: the TOC heading is the first thing on page 2.
      expect(pageOf(toc!.id)).toEqual([2]);
      expect(pages[1].fragments[0]?.blockId).toBe(toc!.id);
      // Word: "Introduction" (Heading 1) starts page 3's content.
      expect(pageOf(intro!.id)).toEqual([3]);

      // The header (11 × 12pt Arial lines below a 567-twip header distance)
      // pushes the body top well past the 680-twip top margin: 38px + 11 × 18.4px.
      const page1 = res.layout.pages[0];
      expect(page1.margins.top).toBeCloseTo(38 + 11 * 12 * (2355 / 2048) * (96 / 72), 0);
    },
    30_000
  );
});

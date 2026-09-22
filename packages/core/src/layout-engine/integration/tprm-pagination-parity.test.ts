/**
 * Word pagination parity against a real customer questionnaire (Google Docs
 * authored, A4, Arial throughout, no explicit page breaks).
 *
 * The fixture is NOT in the repo (customer file). Point `DOCX_TPRM_FIXTURE`
 * at "TPRM - FULL CYBER QUESTIONNAIRE.docx" to run it; otherwise the suite
 * skips loudly.
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

import { describe, test, expect } from 'bun:test';
import fs from 'node:fs';

if (typeof document === 'undefined') {
  (globalThis as Record<string, unknown>).document = {
    createElement: () => ({
      getContext: () => ({
        font: '',
        measureText(text: string) {
          const m = /([\d.]+)px/.exec(this.font as string);
          const px = m ? parseFloat(m[1]) : 16;
          return { width: text.length * px * 0.5 };
        },
      }),
    }),
    documentElement: { style: {} },
    head: { appendChild() {} },
    fonts: { check: () => true, load: async () => [] },
  };
}

const FIXTURE = process.env.DOCX_TPRM_FIXTURE;
const available = !!FIXTURE && fs.existsSync(FIXTURE);
if (!available) {
  console.warn(
    '[tprm-pagination-parity] SKIPPED — set DOCX_TPRM_FIXTURE to the path of ' +
      '"TPRM - FULL CYBER QUESTIONNAIRE.docx" to run the Word pagination parity check.'
  );
}
const run = available ? test : test.skip;

type Fragment = { blockId: unknown };
type ParagraphLike = { kind: string; id: unknown; runs?: Array<{ text?: string }> };

async function layoutFixture() {
  const { EditorState } = await import('prosemirror-state');
  const { parseDocx } = await import('../../docx/parser');
  const { toProseDoc } = await import('../../prosemirror/conversion/toProseDoc');
  const { schema } = await import('../../prosemirror/schema');
  const { computeLayout } = await import('../../editor/computeLayout');
  const bridge = await import('../../layout-bridge');
  const { measureParagraph } = await import('../../layout-bridge/measuring/measureParagraph');

  // preloadFonts=false: the embedded/Google font step is network-bound and
  // irrelevant here (line heights come from the OS/2 ratio table).
  const doc = await parseDocx(new Uint8Array(fs.readFileSync(FIXTURE!)), { preloadFonts: false });
  const body = doc.package.document!;
  const sectionProperties = body.sections?.[0]?.properties ?? body.finalSectionProperties ?? null;
  const finalSectionProperties = body.finalSectionProperties ?? sectionProperties;
  const styles = doc.package.styles;
  const theme = doc.package.theme ?? null;

  const state = EditorState.create({ doc: toProseDoc(doc, { styles }), schema });
  const pageSize = bridge.getPageSize(sectionProperties);
  const margins = bridge.getMargins(sectionProperties);

  // Mirrors the React adapter's measureBlock (minus caching).
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const measureBlock = (block: any, width: number, zones?: any, cumulativeY?: number): any => {
    switch (block.kind) {
      case 'paragraph':
        return measureParagraph(block, width, {
          floatingZones: zones,
          paragraphYOffset: cumulativeY ?? 0,
        });
      case 'table':
        return bridge.measureTableBlock(block, width, measureBlock);
      case 'image':
        return { kind: 'image', width: block.width ?? 100, height: block.height ?? 100 };
      case 'textBox': {
        const m = block.margins ?? { top: 3.6, bottom: 3.6, left: 7.2, right: 7.2 };
        const innerWidth = (block.width ?? 200) - m.left - m.right;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const inner = block.content.map((p: any) => measureParagraph(p, innerWidth));
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const h = inner.reduce((s: number, x: any) => s + x.totalHeight, 0);
        return {
          kind: 'textBox',
          width: block.width ?? 200,
          height: block.height ?? h + m.top + m.bottom,
          innerMeasures: inner,
        };
      }
      default:
        return { kind: block.kind };
    }
  };

  const hf = bridge.resolveHeaderFooter(doc, sectionProperties);
  return computeLayout({
    state,
    document: doc,
    pageSize,
    margins,
    columns: bridge.getColumns(sectionProperties),
    finalPageSize: bridge.getPageSize(finalSectionProperties),
    finalMargins: bridge.getMargins(finalSectionProperties),
    finalColumns: bridge.getColumns(finalSectionProperties),
    pageGap: 20,
    contentWidth: pageSize.w - margins.left - margins.right,
    theme,
    styles,
    sectionProperties,
    finalSectionProperties,
    headerContent: hf.header,
    footerContent: hf.footer,
    firstPageHeaderContent: hf.firstHeader,
    firstPageFooterContent: hf.firstFooter,
    measureBlocks: (blocks, w, geom) =>
      bridge.measureBlocksWithFloats(blocks, w, measureBlock, geom),
    getHfPmDoc: () => null,
  });
}

function paragraphText(b: ParagraphLike): string {
  return (b.runs ?? []).map((r) => r.text ?? '').join('');
}

describe('TPRM questionnaire paginates like Word', () => {
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

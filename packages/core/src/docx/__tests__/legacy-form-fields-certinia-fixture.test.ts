/**
 * Legacy `FORMDROPDOWN` display parity against a real customer matrix: the
 * 6sense "PSA Platform Evaluation RFP" (Certinia), whose Response Code column
 * is 93 legacy dropdowns with `w:listEntry` FS/SC/SX/TP/NS/RM, no
 * `w:ddList/w:result` and no result run. Word displays the current entry
 * ("FS") in every cell; the editor painted them empty until the parser
 * synthesized the display run.
 *
 * The fixture is NOT in the repo (customer file). Point `DOCX_CERTINIA_FIXTURE`
 * at it, or keep `6sense_PSA_Platform_Evaluation_RFP.docx` in `~/Downloads`;
 * otherwise the suite skips loudly.
 *
 * The layout check runs `computeLayout` headless with a canvas stub (0.5em per
 * character — wrapping is irrelevant, every code is a single two-letter line).
 */

import { describe, test, expect, beforeAll, afterAll } from 'bun:test';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import JSZip from 'jszip';
import { installCanvasDocumentStub } from '../../layout-engine/integration/helpers';

const CANDIDATES = [
  process.env.DOCX_CERTINIA_FIXTURE,
  path.join(os.homedir(), 'Downloads', '6sense_PSA_Platform_Evaluation_RFP.docx'),
].filter((p): p is string => !!p);
const FIXTURE = CANDIDATES.find((p) => fs.existsSync(p));
if (!FIXTURE) {
  console.warn(
    '[legacy-form-fields-certinia-fixture] SKIPPED — set DOCX_CERTINIA_FIXTURE to the path of ' +
      '"6sense_PSA_Platform_Evaluation_RFP.docx" (or keep it in ~/Downloads) to run the ' +
      'legacy FORMDROPDOWN display parity check.'
  );
}
const run = FIXTURE ? test : test.skip;

const CODES = ['FS', 'SC', 'SX', 'TP', 'NS', 'RM'];
const EXPECTED_FIELDS = 93;

type ParagraphLike = {
  kind: string;
  paraId?: string;
  runs?: Array<{ kind: string; text?: string }>;
};
type TableLike = { kind: string; rows?: Array<{ cells: Array<{ blocks: unknown[] }> }> };

async function loadFixture() {
  const { parseDocx } = await import('../parser');
  const bytes = new Uint8Array(fs.readFileSync(FIXTURE!));
  const doc = await parseDocx(bytes, { preloadFonts: false });
  const zip = await JSZip.loadAsync(bytes);
  const sourceXml = await zip.file('word/document.xml')!.async('string');
  return { doc, sourceXml };
}

/** paraIds of the paragraphs holding a legacy dropdown, in document order. */
function dropdownParagraphIds(body: { content: unknown[] }): string[] {
  const ids: string[] = [];
  const walk = (blocks: unknown[]) => {
    for (const b of blocks as Array<Record<string, unknown>>) {
      if (b.type === 'paragraph') {
        const content = b.content as Array<Record<string, unknown>>;
        const hit = content.some(
          (c) =>
            c.type === 'inlineSdt' &&
            (c.properties as { legacyFormField?: { fieldType?: string } }).legacyFormField
              ?.fieldType === 'dropdown'
        );
        if (hit) ids.push(String(b.paraId));
      } else if (b.type === 'table') {
        for (const row of b.rows as Array<{ cells: Array<{ content: unknown[] }> }>) {
          for (const cell of row.cells) walk(cell.content);
        }
      } else if (b.type === 'blockSdt') {
        walk(b.content as unknown[]);
      }
    }
  };
  walk(body.content);
  return ids;
}

describe('legacy FORMDROPDOWN display parity — 6sense / Certinia RFP fixture', () => {
  // Canvas stub for the headless layout check, scoped to this suite (see
  // installCanvasDocumentStub for why it must not live at module scope).
  let restoreDocument: () => void = () => {};
  beforeAll(() => {
    if (FIXTURE) restoreDocument = installCanvasDocumentStub();
  });
  afterAll(() => restoreDocument());

  run('parses 93 result-less dropdowns and reports the current entry as their text', async () => {
    const { findContentControls } = await import('../../agent/contentControls');
    const { doc } = await loadFixture();

    const fields = findContentControls(doc, { source: 'legacy' });
    expect(fields).toHaveLength(EXPECTED_FIELDS);
    for (const f of fields) {
      expect(f.legacyFormField!.fieldType).toBe('dropdown');
      expect(f.legacyFormField!.options).toEqual(CODES);
      expect(f.legacyFormField!.hasResult).toBe(false);
      expect(CODES).toContain(f.text);
      expect(f.text).toBe(f.legacyFormField!.value!);
    }
    // Nobody has answered yet: every cell shows entry 0, as in Word.
    expect(new Set(fields.map((f) => f.text))).toEqual(new Set(['FS']));
  });

  run('lays out every Response Code cell with its code as visible text', async () => {
    const { EditorState } = await import('prosemirror-state');
    const { toProseDoc } = await import('../../prosemirror/conversion/toProseDoc');
    const { schema } = await import('../../prosemirror/schema');
    const { computeLayout } = await import('../../editor/computeLayout');
    const bridge = await import('../../layout-bridge');
    const { measureParagraph } = await import('../../layout-bridge/measuring/measureParagraph');
    const { doc } = await loadFixture();

    const body = doc.package.document!;
    const wanted = new Set(dropdownParagraphIds(body));
    expect(wanted.size).toBe(EXPECTED_FIELDS);

    const sectionProperties = body.sections?.[0]?.properties ?? body.finalSectionProperties ?? null;
    const finalSectionProperties = body.finalSectionProperties ?? sectionProperties;
    const styles = doc.package.styles;
    const theme = doc.package.theme ?? null;
    const state = EditorState.create({ doc: toProseDoc(doc, { styles }), schema });
    const pageSize = bridge.getPageSize(sectionProperties);
    const margins = bridge.getMargins(sectionProperties);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const measureBlock = (block: any, width: number, zones?: any, y?: number): any => {
      switch (block.kind) {
        case 'paragraph':
          return measureParagraph(block, width, { floatingZones: zones, paragraphYOffset: y ?? 0 });
        case 'table':
          return bridge.measureTableBlock(block, width, measureBlock);
        case 'image':
          return { kind: 'image', width: block.width ?? 100, height: block.height ?? 100 };
        default:
          return { kind: block.kind };
      }
    };
    const hf = bridge.resolveHeaderFooter(doc, sectionProperties);
    const res = computeLayout({
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
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      measureBlocks: (blocks: any[], w: number | number[], geom?: any) =>
        bridge.measureBlocksWithFloats(blocks, w, measureBlock, geom),
      getHfPmDoc: () => null,
    });

    const cellTexts = new Map<string, string>();
    const walk = (blocks: unknown[]) => {
      for (const b of blocks as Array<ParagraphLike & TableLike>) {
        if (b.kind === 'paragraph' && b.paraId && wanted.has(b.paraId)) {
          cellTexts.set(
            b.paraId,
            (b.runs ?? []).map((r) => (r.kind === 'text' ? (r.text ?? '') : '')).join('')
          );
        } else if (b.kind === 'table') {
          for (const row of b.rows ?? []) for (const cell of row.cells) walk(cell.blocks);
        }
      }
    };
    walk(res.blocks as unknown[]);

    expect(cellTexts.size).toBe(EXPECTED_FIELDS);
    for (const [paraId, text] of cellTexts) {
      expect(CODES, `Response Code cell ${paraId} laid out as ${JSON.stringify(text)}`).toContain(
        text
      );
    }
    expect(res.layout.pages.length).toBeGreaterThan(1);
  });

  run('serializes every untouched field byte for byte (no display run written back)', async () => {
    const { serializeDocumentBody } = await import('../serializer/documentSerializer');
    const { doc, sourceXml } = await loadFixture();

    // Every `begin … end` run sequence in the source, compared modulo two
    // markers that are not field content: Word's bookmark for the field name
    // sits inside the first sequence and the parser lifts bookmarks out of
    // field sequences (pre-existing, harmless to Word); and the two sequences
    // that start a page carry `w:lastRenderedPageBreak`, Word's render cache,
    // which the paragraph serializer re-emits at the run start rather than
    // after `w:rPr`.
    const stripMarkers = (xml: string) =>
      xml.replace(/<w:bookmark(?:Start|End)\b[^>]*\/>|<w:lastRenderedPageBreak\/>/g, '');
    const sequences =
      stripMarkers(sourceXml).match(
        /<w:r>(?:(?!<w:r>).)*?<w:fldChar w:fldCharType="begin">.*?<w:fldChar w:fldCharType="end"\/><\/w:r>/gs
      ) ?? [];
    expect(sequences).toHaveLength(EXPECTED_FIELDS);

    const out = stripMarkers(serializeDocumentBody(doc.package.document!));
    for (const seq of sequences) expect(out).toContain(seq);
    expect(out.match(/ FORMDROPDOWN /g)).toHaveLength(EXPECTED_FIELDS);
    expect(out).not.toContain('<w:t>FS</w:t>');
    expect(out).not.toContain('<w:sdt>');
  });
});

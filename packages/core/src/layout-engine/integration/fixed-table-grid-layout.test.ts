/**
 * Fixed-layout tables take their column geometry from `w:tblGrid`, and rows
 * that don't cover the whole grid (`w:gridBefore` / `w:gridAfter`) stay short.
 *
 * Shape under test (synthetic RFP-style requirement table):
 *  - `w:tblLayout w:type="fixed"`, tblW = sum of the grid;
 *  - a 10-column grid with hairline spacer columns:
 *      1461 ×4 | 1455 | 6 | 1449 | 12 | 1449 | 6   (dxa);
 *  - the header row has 7 cells spanning 1,1,1,1,1,2,2 plus `w:gridAfter=1`;
 *  - body rows have 7 cells spanning 1,1,1,1,2,2,2 — so header and body rows
 *    cover DIFFERENT grid columns, and their tcW values (1461 / 1455) don't
 *    match the grid columns they sit on.
 *
 * Word's truth: every cell is exactly as wide as the grid columns it spans, so
 * all seven header and body columns are ~1" (~97px) wide and line up.
 *
 * Before the fix the header row's `w:gridAfter` was dropped on parse, so the
 * row looked one grid column short. prosemirror-tables' `fixTables` (run by
 * `tableEditing()` on the first doc-changing transaction — e.g. typing an
 * answer, or a collab sync) "repaired" it by inserting an empty cell at the
 * START of the header row, shifting every header cell one grid column right:
 * "Compliance" landed on the 6-dxa (0.4px) spacer column and rendered one
 * character wide, text stacked vertically, with an enormously tall header row.
 *
 * A second table body row carries an answer taller than a page: Word's default
 * (no `w:cantSplit`) breaks the row across pages rather than clipping it.
 *
 * A real document with this table shape can be pointed at via
 * `DOCX_FIXTURE_TABLE_GRID` for an extra whole-document check (skipped when
 * unset).
 */

import { describe, test, expect, beforeAll, afterAll } from 'bun:test';
import fs from 'node:fs';
import JSZip from 'jszip';
import { EditorState } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import { fixTables } from 'prosemirror-tables';
import { installCanvasDocumentStub } from './helpers';
import { layoutDocxHeadless, parseDocxHeadless, toProseDocHeadless } from './headlessDocxLayout';
import { singletonManager } from '../../prosemirror/schema';
import { resolveCellGrid } from '../../layout-bridge/tableWidthUtils';
import { repackDocx } from '../../docx/rezip';
import type { TableBlock, TableFragment, TableMeasure } from '../types';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const GRID = [1461, 1461, 1461, 1461, 1455, 6, 1449, 12, 1449, 6];
const HEADER_SPANS = [1, 1, 1, 1, 1, 2, 2];
const BODY_SPANS = [1, 1, 1, 1, 2, 2, 2];
const HEADER_TEXT = [
  'ID',
  'Requirement',
  'Area',
  'Priority',
  'Compliance',
  'Comments',
  'Reference',
];
const LONG_ANSWER_PARAGRAPHS = 60;

const px = (twips: number) => twips / 15;

function para(text: string): string {
  return `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
}

function cell(tcW: number, span: number, content: string): string {
  const gridSpan = span > 1 ? `<w:gridSpan w:val="${span}"/>` : '';
  return `<w:tc><w:tcPr><w:tcW w:w="${tcW}" w:type="dxa"/>${gridSpan}</w:tcPr>${content}</w:tc>`;
}

function grid(cols: number[]): string {
  return `<w:tblGrid>${cols.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>`;
}

const BORDERS =
  '<w:tblBorders>' +
  ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
    .map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`)
    .join('') +
  '</w:tblBorders>';

function requirementTable(): string {
  // tcW deliberately disagrees with the grid columns each cell spans.
  const header =
    '<w:tr><w:trPr><w:gridAfter w:val="1"/><w:wAfter w:w="6" w:type="dxa"/><w:tblHeader/></w:trPr>' +
    HEADER_SPANS.map((span, i) => cell(i < 4 ? 1461 : 1455, span, para(HEADER_TEXT[i]))).join('') +
    '</w:tr>';
  const body = (id: string, notes: string) =>
    '<w:tr>' +
    BODY_SPANS.map((span, i) => {
      const text = [id, 'Requirement text', 'Data Security', 'Required', 'Yes', '', ''][i];
      const content = i === 5 ? notes : para(text);
      return cell(i < 6 ? 1461 : 1455, span, content);
    }).join('') +
    '</w:tr>';
  const longAnswer = Array.from({ length: LONG_ANSWER_PARAGRAPHS }, (_, i) =>
    para(`Answer paragraph ${i + 1} of a long vendor response`)
  ).join('');
  return (
    '<w:tbl><w:tblPr><w:tblW w:w="10221" w:type="dxa"/>' +
    BORDERS +
    '<w:tblLayout w:type="fixed"/></w:tblPr>' +
    grid(GRID) +
    header +
    body('R-01', para('Short answer')) +
    body('R-02', longAnswer) +
    body('R-03', para('Short answer')) +
    '</w:tbl>'
  );
}

/** Autofit (no tblLayout) table whose tcW agree with its grid. */
function autofitTable(): string {
  return (
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>' +
    BORDERS +
    '</w:tblPr>' +
    grid([3000, 7224]) +
    `<w:tr>${cell(3000, 1, para('Field'))}${cell(7224, 1, para('Detail'))}</w:tr>` +
    `<w:tr>${cell(3000, 1, para('Issued By'))}${cell(7224, 1, para('IT Department'))}</w:tr>` +
    '</w:tbl>'
  );
}

/** Fixed table whose second row starts one grid column in (`w:gridBefore`). */
function gridBeforeTable(): string {
  return (
    '<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/>' +
    BORDERS +
    '<w:tblLayout w:type="fixed"/></w:tblPr>' +
    grid([2000, 3000, 3000]) +
    `<w:tr>${cell(2000, 1, para('A'))}${cell(3000, 1, para('B'))}${cell(3000, 1, para('C'))}</w:tr>` +
    '<w:tr><w:trPr><w:gridBefore w:val="1"/><w:wBefore w:w="2000" w:type="dxa"/></w:trPr>' +
    `${cell(3000, 1, para('B2'))}${cell(3000, 1, para('C2'))}</w:tr>` +
    '</w:tbl>'
  );
}

const DOCUMENT_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}"><w:body>` +
  para('Requirements') +
  requirementTable() +
  para('') +
  autofitTable() +
  para('') +
  gridBeforeTable() +
  para('') +
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="936" w:right="1008" w:bottom="936" w:left="1008" w:header="720" w:footer="720" w:gutter="0"/>' +
  '</w:sectPr></w:body></w:document>';

async function buildDocx(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `</Types>`
  );
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
      `</Relationships>`
  );
  zip.file('word/document.xml', DOCUMENT_XML);
  return new Uint8Array(await zip.generateAsync({ type: 'arraybuffer' }));
}

function tables(doc: PMNode): Array<{ node: PMNode; pos: number }> {
  const out: Array<{ node: PMNode; pos: number }> = [];
  doc.descendants((node, pos) => {
    if (node.type.name === 'table') {
      out.push({ node, pos });
      return false;
    }
    return true;
  });
  return out;
}

function rowTexts(row: PMNode): string[] {
  const out: string[] = [];
  row.forEach((c) => out.push(c.textContent));
  return out;
}

/** Editor state with the real plugin set (incl. the table-editing plugins). */
function editorState(pmDoc: PMNode): EditorState {
  return EditorState.create({
    doc: pmDoc,
    schema: singletonManager.getSchema(),
    plugins: [...singletonManager.getPlugins()],
  });
}

/**
 * Type an answer into the first body row's Compliance cell — a doc-changing
 * transaction, which is what makes `tableEditing()` run its table repair.
 */
function typeAnswer(state: EditorState, tableIndex = 0): EditorState {
  const { node, pos } = tables(state.doc)[tableIndex]!;
  const bodyRow = node.child(1);
  let offset = pos + 1 + node.child(0).nodeSize + 1; // inside body row 1
  for (let i = 0; i < 4; i++) offset += bodyRow.child(i).nodeSize;
  // offset now points at the Compliance cell; +1 cell open, +1 paragraph open.
  const tr = state.tr.insertText('No', offset + 2);
  return state.applyTransaction(tr).state;
}

type LayoutResult = Awaited<ReturnType<typeof layoutDocxHeadless>>;

function tableAt(res: LayoutResult, n: number): { block: TableBlock; measure: TableMeasure } {
  const idx = res.blocks.map((b, i) => (b.kind === 'table' ? i : -1)).filter((i) => i >= 0)[n]!;
  return {
    block: res.blocks[idx] as TableBlock,
    measure: res.measures[idx] as TableMeasure,
  };
}

/** Pixel width of the grid columns `[from, from + span)`. */
function gridWidth(cols: number[], from: number, span: number): number {
  return cols.slice(from, from + span).reduce((a, b) => a + b, 0);
}

describe('fixed-layout tables follow w:tblGrid like Word', () => {
  let restoreDocument: () => void = () => {};
  let bytes: Uint8Array;
  beforeAll(async () => {
    restoreDocument = installCanvasDocumentStub();
    bytes = await buildDocx();
  });
  afterAll(() => restoreDocument());

  test('w:gridBefore / w:gridAfter / w:wBefore / w:wAfter parse and survive a save', async () => {
    const doc = await parseDocxHeadless(bytes);
    const blocks = doc.package.document!.content.filter((b) => b.type === 'table');
    const header = blocks[0]!.type === 'table' ? blocks[0]!.rows[0]! : null;
    expect(header?.formatting?.gridAfter).toBe(1);
    expect(header?.formatting?.widthAfter).toEqual({ value: 6, type: 'dxa' });
    const shifted = blocks[2]!.type === 'table' ? blocks[2]!.rows[1]! : null;
    expect(shifted?.formatting?.gridBefore).toBe(1);
    expect(shifted?.formatting?.widthBefore).toEqual({ value: 2000, type: 'dxa' });

    const saved = await JSZip.loadAsync(await repackDocx(doc));
    const xml = await saved.file('word/document.xml')!.async('string');
    expect(xml).toContain('<w:gridAfter w:val="1"/>');
    expect(xml).toContain('<w:wAfter w:w="6" w:type="dxa"/>');
    expect(xml).toContain('<w:gridBefore w:val="1"/>');
    expect(xml).toContain('<w:wBefore w:w="2000" w:type="dxa"/>');
  });

  test('editing does not pad a row that is short of its grid', async () => {
    const doc = await parseDocxHeadless(bytes);
    const pmDoc = await toProseDocHeadless(doc);

    // The hazard: upstream fixTables WOULD insert a cell into the header row.
    expect(fixTables(EditorState.create({ doc: pmDoc }))).toBeDefined();

    const edited = typeAnswer(editorState(pmDoc));
    const [req, , shifted] = tables(edited.doc);
    expect(req!.node.child(0).childCount).toBe(7);
    expect(rowTexts(req!.node.child(0))).toEqual(HEADER_TEXT);
    expect(rowTexts(req!.node.child(1))[4]).toBe('NoYes');
    expect(shifted!.node.child(1).childCount).toBe(2);
  });

  test('cell widths are the sum of the grid columns each cell spans', async () => {
    const doc = await parseDocxHeadless(bytes);
    const state = typeAnswer(editorState(await toProseDocHeadless(doc)));
    const res = await layoutDocxHeadless(doc, state);
    const { block, measure } = tableAt(res, 0);

    const gridPx = GRID.map(px);
    measure.columnWidths.forEach((w, i) => expect(w).toBeCloseTo(gridPx[i]!, 3));

    for (const g of resolveCellGrid(block)) {
      const width = measure.rows[g.rowIndex]!.cells[g.cellIndex]!.width;
      expect(width).toBeCloseTo(gridWidth(gridPx, g.columnIndex, g.colSpan), 3);
    }

    // Header and body columns line up: every visible column is ~1" wide, and
    // the Compliance header cell is NOT the 0.4px spacer column.
    const headerWidths = measure.rows[0]!.cells.map((c) => c.width);
    const bodyWidths = measure.rows[1]!.cells.map((c) => c.width);
    for (const w of [...headerWidths, ...bodyWidths]) expect(w).toBeGreaterThan(96);
    expect(headerWidths[4]).toBeCloseTo(px(1455), 3);
    expect(bodyWidths[4]).toBeCloseTo(px(1455 + 6), 3);

    // The header row stays a normal height (text doesn't stack vertically).
    expect(measure.rows[0]!.height).toBeLessThan(60);
  });

  test('w:gridBefore indents the row by the skipped grid columns', async () => {
    const doc = await parseDocxHeadless(bytes);
    const state = typeAnswer(editorState(await toProseDocHeadless(doc)));
    const res = await layoutDocxHeadless(doc, state);
    const { block, measure } = tableAt(res, 2);
    const shifted = resolveCellGrid(block).filter((g) => g.rowIndex === 1);
    expect(shifted.map((g) => g.columnIndex)).toEqual([1, 2]);
    expect(measure.rows[1]!.cells.map((c) => c.width)).toEqual([px(3000), px(3000)]);
  });

  test('autofit tables keep their grid widths', async () => {
    const doc = await parseDocxHeadless(bytes);
    const res = await layoutDocxHeadless(doc);
    const { measure } = tableAt(res, 1);
    expect(measure.columnWidths[0]).toBeCloseTo(px(3000), 3);
    expect(measure.columnWidths[1]).toBeCloseTo(px(7224), 3);
  });

  test('a row taller than the page breaks across pages instead of clipping', async () => {
    const doc = await parseDocxHeadless(bytes);
    const res = await layoutDocxHeadless(doc);
    const { block, measure } = tableAt(res, 0);
    const LONG_ROW = 2;
    const contentHeight = (15840 - 936 * 2) / 15;
    expect(measure.rows[LONG_ROW]!.height).toBeGreaterThan(contentHeight);

    const frags = res.layout.pages.flatMap((p) =>
      p.fragments.filter(
        (f): f is TableFragment =>
          f.kind === 'table' &&
          f.blockId === block.id &&
          f.fromRow <= LONG_ROW &&
          LONG_ROW < f.toRow
      )
    );
    expect(frags.length).toBeGreaterThanOrEqual(2);

    // The visible bands of the long row tile its full height: nothing lost.
    let shown = 0;
    for (const f of frags) {
      const start = f.fromRow === LONG_ROW ? (f.topClip ?? 0) : 0;
      const end =
        f.toRow - 1 === LONG_ROW && f.bottomClip !== undefined
          ? f.bottomClip
          : measure.rows[LONG_ROW]!.height;
      shown += end - start;
    }
    expect(shown).toBeCloseTo(measure.rows[LONG_ROW]!.height, 0);
  });
});

// ---------------------------------------------------------------------------
// Optional: a real document with this table shape (not in the repo).
// ---------------------------------------------------------------------------

const FIXTURE = process.env.DOCX_FIXTURE_TABLE_GRID;
const available = !!FIXTURE && fs.existsSync(FIXTURE);
if (!available) {
  console.warn(
    '[fixed-table-grid-layout] real-file check SKIPPED — set DOCX_FIXTURE_TABLE_GRID to the path ' +
      'of a .docx with fixed-layout gridBefore/gridAfter tables to run it.'
  );
}
const runFixture = available ? test : test.skip;

describe('real fixed-grid requirement tables (fixture) render like Word', () => {
  let restoreDocument: () => void = () => {};
  beforeAll(() => {
    if (available) restoreDocument = installCanvasDocumentStub();
  });
  afterAll(() => restoreDocument());

  runFixture(
    'every cell is as wide as its grid columns after an edit',
    async () => {
      const doc = await parseDocxHeadless(new Uint8Array(fs.readFileSync(FIXTURE!)));
      const pmDoc = await toProseDocHeadless(doc);
      const before = tables(pmDoc).map(({ node }) => {
        const counts: number[] = [];
        node.forEach((row) => counts.push(row.childCount));
        return counts;
      });
      // First requirement table (7 cells per body row).
      const reqTable = tables(pmDoc).findIndex(
        ({ node }) => node.childCount > 1 && node.child(1).childCount === 7
      );
      expect(reqTable).toBeGreaterThanOrEqual(0);
      const state = typeAnswer(editorState(pmDoc), reqTable);
      tables(state.doc).forEach(({ node }, t) => {
        node.forEach((row, _o, r) => expect(row.childCount).toBe(before[t]![r]!));
      });

      const res = await layoutDocxHeadless(doc, state);
      const tableCount = res.blocks.filter((b) => b.kind === 'table').length;
      expect(tableCount).toBeGreaterThan(0);
      for (let t = 0; t < tableCount; t++) {
        const { block, measure } = tableAt(res, t);
        for (const g of resolveCellGrid(block)) {
          const width = measure.rows[g.rowIndex]!.cells[g.cellIndex]!.width;
          expect(width).toBeCloseTo(gridWidth(measure.columnWidths, g.columnIndex, g.colSpan), 3);
          // No visible column collapses onto a hairline spacer column.
          expect(width).toBeGreaterThan(40);
        }
      }
    },
    30_000
  );
});

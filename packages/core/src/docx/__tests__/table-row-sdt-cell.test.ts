/**
 * Content controls that wrap a whole table cell at ROW level:
 *
 * ```xml
 * <w:tr><w:tc>…question…</w:tc>
 *   <w:sdt><w:sdtPr>…</w:sdtPr><w:sdtContent><w:tc>…answer…</w:tc></w:sdtContent></w:sdt>
 * </w:tr>
 * ```
 *
 * ECMA-376 `CT_Row` content is `tc | customXml | sdt` (the sdt's content being
 * `CT_SdtContentCell`). Word writes this for questionnaire answer cells. The
 * row parser used to collect only direct `w:tc` children, so the wrapped cell
 * and all its paragraphs vanished and the row came out one cell short.
 *
 * The wrapped cell is now a real cell whose content is the control's
 * `BlockSdt`; the save re-emits the wrapper around the `w:tc` byte-for-byte.
 *
 * A real questionnaire with row-level answer controls can be pointed at via
 * `DOCX_FIXTURE_CELL_CONTROLS` for an extra whole-document check (skipped when
 * unset).
 */

import { describe, test, expect, beforeAll, afterAll } from 'bun:test';
import fs from 'node:fs';
import JSZip from 'jszip';
import { EditorState } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import { installCanvasDocumentStub } from '../../layout-engine/integration/helpers';
import {
  layoutDocxHeadless,
  parseDocxHeadless,
  toProseDocHeadless,
} from '../../layout-engine/integration/headlessDocxLayout';
import { schema } from '../../prosemirror/schema';
import { fromProseDoc } from '../../prosemirror/conversion/fromProseDoc';
import { repackDocx } from '../rezip';
import {
  findContentControlsInPM,
  setContentControlValueTr,
} from '../../prosemirror/contentControls';
import { findContentControls } from '../../agent/contentControls';
import { setContentControlValue } from '../../agent/contentControlValues';
import { resolveCellGrid } from '../../layout-bridge/tableWidthUtils';
import { isTrustedRowCustomXml } from '../serializer/rowWrapperSerializer';
import { projectProseMirrorDocument, rehydrateCollaborationDocument } from '../../collaboration';
import type {
  BlockContent,
  BlockSdt,
  Document,
  Paragraph,
  Table,
  TableCell,
} from '../../types/document';
import type { ParagraphBlock, TableBlock, TableMeasure } from '../../layout-engine/types';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';

const TEXT_SDT_PR =
  '<w:sdtPr><w:rPr><w:sz w:val="18"/></w:rPr><w:id w:val="500000001"/>' +
  '<w:placeholder><w:docPart w:val="00000000000000000000000000000002"/></w:placeholder>' +
  '<w:showingPlcHdr/><w:text/></w:sdtPr>';

const DROPDOWN_SDT_PR =
  '<w:sdtPr><w:alias w:val="Yes/No"/><w:tag w:val="Answer"/><w:id w:val="500000002"/>' +
  '<w:showingPlcHdr/><w:dropDownList><w:listItem w:value="Choose an item."/>' +
  '<w:listItem w:displayText="Yes" w:value="Yes"/><w:listItem w:displayText="No" w:value="No"/>' +
  '</w:dropDownList></w:sdtPr>';

const GROUP_SDT_PR = '<w:sdtPr><w:tag w:val="Pair"/><w:id w:val="99"/></w:sdtPr>';

const para = (id: string, text: string, placeholder = false) =>
  `<w:p w14:paraId="${id}"><w:r>` +
  (placeholder ? '<w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr>' : '') +
  `<w:t>${text}</w:t></w:r></w:p>`;

const tc = (width: number, body: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr>${body}</w:tc>`;

const rowSdt = (sdtPr: string, cells: string) =>
  `<w:sdt>${sdtPr}<w:sdtEndPr/><w:sdtContent>${cells}</w:sdtContent></w:sdt>`;

/** Two-column question grid: each answer cell is wrapped by a row-level control. */
const TABLE_A =
  '<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="6000"/></w:tblGrid>' +
  '<w:tr>' +
  tc(3000, para('2A000003', 'Company name')) +
  rowSdt(TEXT_SDT_PR, tc(6000, para('2A000001', 'Click or tap here to enter text.', true))) +
  '</w:tr>' +
  '<w:tr>' +
  tc(3000, para('1A2B3C4D', 'Do you encrypt data at rest?')) +
  rowSdt(DROPDOWN_SDT_PR, tc(6000, para('2A000002', 'Choose an item.', true))) +
  '</w:tr>' +
  '</w:tbl>';

/**
 * Three-column grid with the rarer shapes: one control wrapping two cells,
 * and a control nested in a row-level customXml.
 */
const TABLE_B =
  '<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
  '<w:tr>' +
  tc(3000, para('00000B01', 'Pair')) +
  rowSdt(GROUP_SDT_PR, tc(3000, para('00000B02', 'Left')) + tc(3000, para('00000B03', 'Right'))) +
  '</w:tr>' +
  '<w:tr>' +
  tc(3000, para('00000B04', 'Wrapped')) +
  '<w:customXml w:uri="urn:x" w:element="answer"><w:customXmlPr><w:attr w:name="q" w:val="7"/></w:customXmlPr>' +
  rowSdt(
    '<w:sdtPr><w:tag w:val="Inner"/><w:id w:val="100"/><w:text/></w:sdtPr>',
    tc(3000, para('00000B05', 'Inside'))
  ) +
  '</w:customXml>' +
  tc(3000, para('00000B06', 'After')) +
  '</w:tr>' +
  '</w:tbl>';

const DOCUMENT_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<w:document xmlns:w="${NS_W}" xmlns:w14="${NS_W14}"><w:body>` +
  TABLE_A +
  '<w:p/>' +
  TABLE_B +
  '<w:p/>' +
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
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

async function savedDocumentXml(doc: Document): Promise<string> {
  const saved = await JSZip.loadAsync(await repackDocx(doc));
  return saved.file('word/document.xml')!.async('string');
}

/** The saved `w:tbl` elements, in order. */
function tablesXml(xml: string): string[] {
  return xml.match(/<w:tbl>.*?<\/w:tbl>(?=<w:p)/gs) ?? [];
}

function tables(doc: Document): Table[] {
  return doc.package.document!.content.filter((b): b is Table => b.type === 'table');
}

function textOf(blocks: BlockContent[] | TableCell['content']): string {
  const out: string[] = [];
  const walk = (list: readonly BlockContent[]) => {
    for (const b of list) {
      if (b.type === 'paragraph') {
        for (const r of (b as Paragraph).content) {
          if (r.type === 'run') for (const c of r.content) if (c.type === 'text') out.push(c.text);
        }
      } else if (b.type === 'blockSdt') walk(b.content);
    }
  };
  walk(blocks as BlockContent[]);
  return out.join('');
}

function editorState(pmDoc: PMNode): EditorState {
  return EditorState.create({ doc: pmDoc, schema });
}

describe('row-level content controls wrapping a table cell', () => {
  let restoreDocument: () => void = () => {};
  let bytes: Uint8Array;
  beforeAll(async () => {
    restoreDocument = installCanvasDocumentStub();
    bytes = await buildDocx();
  });
  afterAll(() => restoreDocument());

  test('the wrapped cell is parsed as a real cell holding the control', async () => {
    const doc = await parseDocxHeadless(bytes);
    const [a] = tables(doc);
    expect(a!.rows.map((r) => r.cells.length)).toEqual([2, 2]);

    const textCell = a!.rows[0]!.cells[1]!;
    expect(textCell.formatting?.width?.value).toBe(6000);
    expect(textCell.rowWrappers).toEqual([{ kind: 'sdt', id: 0, leading: true }]);
    expect(textCell.content.map((b) => b.type)).toEqual(['blockSdt']);
    const textSdt = textCell.content[0] as BlockSdt;
    expect(textSdt.properties.sdtType).toBe('plainText');
    expect(textSdt.properties.id).toBe(500000001);
    expect(textSdt.properties.showingPlaceholder).toBe(true);
    expect(textSdt.properties.rawPropertiesXml).toBe(TEXT_SDT_PR);
    expect((textSdt.content[0] as Paragraph).paraId).toBe('2A000001');

    const ddSdt = a!.rows[1]!.cells[1]!.content[0] as BlockSdt;
    expect(ddSdt.properties.sdtType).toBe('dropDownList');
    expect(ddSdt.properties.tag).toBe('Answer');
    expect((ddSdt.content[0] as Paragraph).paraId).toBe('2A000002');
  });

  test('multi-cell and customXml-nested wrappers keep every cell', async () => {
    const doc = await parseDocxHeadless(bytes);
    const [, b] = tables(doc);
    expect(b!.rows.map((r) => r.cells.length)).toEqual([3, 3]);

    const [, left, right] = b!.rows[0]!.cells;
    expect(left!.rowWrappers).toEqual([{ kind: 'sdt', id: 0, leading: true }]);
    expect(right!.rowWrappers).toEqual([{ kind: 'sdt', id: 0, leading: false }]);
    expect(left!.content.map((x) => x.type)).toEqual(['blockSdt']);
    expect(right!.content.map((x) => x.type)).toEqual(['paragraph']);
    expect(textOf(right!.content)).toBe('Right');

    const inside = b!.rows[1]!.cells[1]!;
    expect(inside.rowWrappers?.map((w) => w.kind)).toEqual(['customXml', 'sdt']);
    expect((inside.content[0] as BlockSdt).properties.tag).toBe('Inner');
    expect(textOf(b!.rows[1]!.cells[2]!.content)).toBe('After');
    expect(b!.rows[1]!.cells[2]!.rowWrappers).toBeUndefined();
  });

  test('a save re-emits the row-level wrappers byte-for-byte', async () => {
    const doc = await parseDocxHeadless(bytes);
    const saved = tablesXml(await savedDocumentXml(doc));
    expect(saved).toEqual([TABLE_A, TABLE_B]);
  });

  test('the editor round-trips it (Document -> PM -> Document -> XML)', async () => {
    const doc = await parseDocxHeadless(bytes);
    const pmDoc = await toProseDocHeadless(doc);
    const rowCells: number[] = [];
    pmDoc.descendants((node) => {
      if (node.type.name === 'tableRow') rowCells.push(node.childCount);
      return true;
    });
    expect(rowCells).toEqual([2, 2, 3, 3]);

    const back = fromProseDoc(pmDoc, doc);
    expect(tablesXml(await savedDocumentXml(back))).toEqual([TABLE_A, TABLE_B]);
  });

  test('the wrapped cell lays out as a real grid cell', async () => {
    const doc = await parseDocxHeadless(bytes);
    const res = await layoutDocxHeadless(doc);
    const idx = res.blocks.findIndex((b) => b.kind === 'table');
    const block = res.blocks[idx] as TableBlock;
    const measure = res.measures[idx] as TableMeasure;

    expect(block.rows.map((r) => r.cells.length)).toEqual([2, 2]);
    expect(resolveCellGrid(block).map((g) => [g.rowIndex, g.columnIndex])).toEqual([
      [0, 0],
      [0, 1],
      [1, 0],
      [1, 1],
    ]);
    expect(measure.rows[0]!.cells.map((c) => c.width)).toEqual([3000 / 15, 6000 / 15]);

    const answer = block.rows[0]!.cells[1]!.blocks[0] as ParagraphBlock;
    expect(answer.runs.map((r) => (r.kind === 'text' ? r.text : '')).join('')).toBe(
      'Click or tap here to enter text.'
    );
    expect(answer.sdtGroups?.map((g) => g.sdtType)).toEqual(['plainText']);
  });

  test('the PM content-control API finds and answers them; the save keeps the wrapper', async () => {
    const doc = await parseDocxHeadless(bytes);
    let state = editorState(await toProseDocHeadless(doc));
    const all = findContentControlsInPM(state.doc);
    expect(all.map((c) => c.id)).toEqual([500000001, 500000002, 99, 100]);

    state = state.apply(
      setContentControlValueTr(state, { id: 500000001 }, { kind: 'text', text: 'Example Corp' })
    );
    state = state.apply(
      setContentControlValueTr(state, { tag: 'Answer' }, { kind: 'dropdown', value: 'Yes' })
    );
    expect(findContentControlsInPM(state.doc, { id: 500000001 })[0]!.text).toBe('Example Corp');
    expect(findContentControlsInPM(state.doc, { tag: 'Answer' })[0]!.text).toBe('Yes');

    const [tblA] = tablesXml(await savedDocumentXml(fromProseDoc(state.doc, doc)));
    const rows = tblA!.split('<w:tr>').slice(1);
    // Wrapper still sits around the answer cell, now without the placeholder flag.
    expect(rows[0]).toMatch(
      /<\/w:tc><w:sdt><w:sdtPr>.*<w:text\/><\/w:sdtPr><w:sdtEndPr\/><w:sdtContent><w:tc>.*Example Corp<\/w:t>.*<\/w:tc><\/w:sdtContent><\/w:sdt><\/w:tr>/
    );
    expect(rows[0]).not.toContain('<w:showingPlcHdr/>');
    expect(rows[1]).toMatch(
      /<\/w:tc><w:sdt><w:sdtPr>.*<w:dropDownList[ >].*<w:sdtContent><w:tc>.*Yes<\/w:t>.*<\/w:tc><\/w:sdtContent><\/w:sdt><\/w:tr>/
    );
  });

  test('the headless content-control API finds and answers them', async () => {
    const doc = await parseDocxHeadless(bytes);
    expect(findContentControls(doc).map((c) => c.id)).toEqual([500000001, 500000002, 99, 100]);

    const answered = setContentControlValue(
      doc,
      { tag: 'Answer' },
      { kind: 'dropdown', value: 'No' }
    );
    expect(findContentControls(answered, { tag: 'Answer' })[0]!.text).toBe('No');
    const cell = tables(answered)[0]!.rows[1]!.cells[1]!;
    expect(cell.content[0]!.type).toBe('blockSdt');
    const [tblA] = tablesXml(await savedDocumentXml(answered));
    expect(tblA).toMatch(
      /<w:sdtContent><w:tc>.*No<\/w:t>.*<\/w:tc><\/w:sdtContent><\/w:sdt><\/w:tr><\/w:tbl>$/
    );
  });

  test('removing the control from the cell drops its wrapper, not the cell', async () => {
    const doc = await parseDocxHeadless(bytes);
    const cell = tables(doc)[0]!.rows[0]!.cells[1]!;
    cell.content = (cell.content[0] as BlockSdt).content as TableCell['content'];
    const [tblA] = tablesXml(await savedDocumentXml(doc));
    const row = tblA!.split('<w:tr>')[1]!;
    expect(row).not.toContain('<w:text/>');
    expect(row).toContain('2A000001');
    expect(row.match(/<w:tc>/g)).toHaveLength(2);
  });
});

describe('row-level customXml wrapper trust boundary', () => {
  const START =
    '<w:customXml w:uri="urn:x" w:element="answer"><w:customXmlPr><w:attr w:name="q" w:val="7"/></w:customXmlPr>';
  const END = '</w:customXml>';

  test('accepts the shape the parser captures', () => {
    expect(isTrustedRowCustomXml(START, END)).toBe(true);
    expect(isTrustedRowCustomXml('<w:customXml w:element="answer">', END)).toBe(true);
    expect(
      isTrustedRowCustomXml(
        '<w:customXml w:element="a"><w:customXmlPr><w:placeholder w:val="x"/></w:customXmlPr>',
        END
      )
    ).toBe(true);
  });

  test('rejects anything that could land foreign markup in the row', () => {
    const injected = '<w:tc><w:p><w:r><w:t>injected</w:t></w:r></w:p></w:tc>';
    for (const [start, end] of [
      // closes early, then smuggles a cell in front of the real one
      [`<w:customXml w:element="a"></w:customXml>${injected}<w:customXml w:element="b">`, END],
      [
        `<w:customXml w:element="a"><w:customXmlPr><w:attr w:name="q" w:val="1"/>${injected}</w:customXmlPr>`,
        END,
      ],
      [
        '<w:customXml w:element="a"><w:customXmlPr><w:attr w:name="q"><w:r/></w:attr></w:customXmlPr>',
        END,
      ],
      ['<w:customXml w:element="a">text', END],
      ['<w:sdt>', '</w:sdt>'],
      [START, `${END}${injected}`],
      ['<w:customXml w:element="a"><w:customXmlPr/><w:customXmlPr/>', END],
    ] as const) {
      expect(isTrustedRowCustomXml(start, end)).toBe(false);
    }
  });

  test('a save drops a tampered wrapper but keeps its cell', async () => {
    const restore = installCanvasDocumentStub();
    try {
      const doc = await parseDocxHeadless(await buildDocx());
      const cell = tables(doc)[1]!.rows[1]!.cells[1]!;
      const custom = cell.rowWrappers![0]!;
      expect(custom.kind).toBe('customXml');
      if (custom.kind !== 'customXml') throw new Error('expected customXml');
      custom.startXml =
        '<w:customXml w:element="a"></w:customXml><w:tc><w:p><w:r><w:t>injected</w:t></w:r></w:p></w:tc><w:customXml w:element="b">';

      const [, tblB] = tablesXml(await savedDocumentXml(doc));
      expect(tblB).not.toContain('injected');
      expect(tblB).not.toContain('<w:customXml');
      const row = tblB!.split('<w:tr>')[2]!;
      expect(row.match(/<w:tc>/g)).toHaveLength(3);
      expect(row).toContain('<w:sdtContent><w:tc>');
      expect(row).toContain('00000B05');
    } finally {
      restore();
    }
  });

  test('collaboration keeps row wrappers out of the shared document', async () => {
    const restore = installCanvasDocumentStub();
    try {
      const doc = await parseDocxHeadless(await buildDocx());
      const pmDoc = await toProseDocHeadless(doc);
      const projected = projectProseMirrorDocument(pmDoc, { baseRevision: 'base-1' });
      const shared = JSON.stringify(projected.document);
      expect(shared).not.toContain('"rowWrappers":[');
      expect(shared).not.toContain('customXml');
      expect(JSON.stringify(projected.sidecar)).toContain('<w:customXml w:uri=\\"urn:x\\"');

      const hydrated = schema.nodeFromJSON(
        rehydrateCollaborationDocument(projected.document, projected.sidecar)
      );
      expect(tablesXml(await savedDocumentXml(fromProseDoc(hydrated, doc)))).toEqual([
        TABLE_A,
        TABLE_B,
      ]);
    } finally {
      restore();
    }
  });
});

// ---------------------------------------------------------------------------
// Optional: a real questionnaire with row-level answer controls (not in the repo).
// ---------------------------------------------------------------------------

const FIXTURE = process.env.DOCX_FIXTURE_CELL_CONTROLS;
const available = !!FIXTURE && fs.existsSync(FIXTURE);
if (!available) {
  console.warn(
    '[table-row-sdt-cell] real-file check SKIPPED — set DOCX_FIXTURE_CELL_CONTROLS to the path ' +
      'of a .docx with row-level (w:tr > w:sdt > w:tc) controls to run it.'
  );
}
const runFixture = available ? test : test.skip;

/** Cells that open a row-level wrapper, with the paraIds of their paragraphs. */
function wrappedCells(doc: Document): Array<{ cell: TableCell; paraIds: string[] }> {
  const out: Array<{ cell: TableCell; paraIds: string[] }> = [];
  const ids = (list: readonly BlockContent[]): string[] =>
    list.flatMap((b) =>
      b.type === 'paragraph'
        ? b.paraId
          ? [b.paraId]
          : []
        : b.type === 'blockSdt'
          ? ids(b.content)
          : []
    );
  for (const table of tables(doc)) {
    for (const row of table.rows) {
      for (const cell of row.cells) {
        if (cell.rowWrappers?.some((w) => w.kind === 'sdt' && w.leading)) {
          out.push({ cell, paraIds: ids(cell.content as BlockContent[]) });
        }
      }
    }
  }
  return out;
}

describe('real questionnaire (fixture) row-level controls', () => {
  let restoreDocument: () => void = () => {};
  beforeAll(() => {
    if (available) restoreDocument = installCanvasDocumentStub();
  });
  afterAll(() => restoreDocument());

  runFixture(
    'answer cells wrapped by row-level controls are kept and saved',
    async () => {
      const raw = new Uint8Array(fs.readFileSync(FIXTURE!));
      const sourceXml = await (await JSZip.loadAsync(raw))
        .file('word/document.xml')!
        .async('string');
      const wrappedInSource = (sourceXml.match(/<w:sdtContent><w:tc>/g) ?? []).length;
      expect(wrappedInSource).toBeGreaterThan(0);

      const doc = await parseDocxHeadless(raw);
      // Every wrapped cell in the source is kept as a cell opening its wrapper.
      const wrapped = wrappedCells(doc);
      expect(wrapped).toHaveLength(wrappedInSource);
      const paraIds = wrapped.flatMap((w) => w.paraIds);

      const state = editorState(await toProseDocHeadless(doc));
      const saved = await savedDocumentXml(fromProseDoc(state.doc, doc));
      expect((saved.match(/<w:sdtContent><w:tc>/g) ?? []).length).toBe(wrappedInSource);
      // The answer paragraphs keep their ids through the save.
      for (const id of paraIds) expect(saved).toContain(`w14:paraId="${id}"`);
    },
    60_000
  );
});

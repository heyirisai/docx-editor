/**
 * Content controls (and customXml) that wrap whole table ROWS:
 *
 * ```xml
 * <w:tbl>…<w:tr>…</w:tr>
 *   <w:sdt><w:sdtPr>…</w:sdtPr><w:sdtContent><w:tr>…</w:tr></w:sdtContent></w:sdt>
 * </w:tbl>
 * ```
 *
 * ECMA-376 `CT_Tbl` row content is `tr | customXml | sdt` (`CT_SdtRow`,
 * `CT_CustomXmlRow`). The table parser used to collect only direct `w:tr`
 * children, so a wrapped row and everything in it vanished from the parse,
 * the layout and the save.
 */

import { describe, test, expect, beforeAll, afterAll } from 'bun:test';
import JSZip from 'jszip';
import { installCanvasDocumentStub } from '../../layout-engine/integration/helpers';
import {
  layoutDocxHeadless,
  parseDocxHeadless,
  toProseDocHeadless,
} from '../../layout-engine/integration/headlessDocxLayout';
import { schema } from '../../prosemirror/schema';
import { fromProseDoc } from '../../prosemirror/conversion/fromProseDoc';
import { repackDocx } from '../rezip';
import { projectProseMirrorDocument, rehydrateCollaborationDocument } from '../../collaboration';
import type { Document, Paragraph, Table, TableRow } from '../../types/document';
import type { TableBlock } from '../../layout-engine/types';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const NS_W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';

const ROW_SDT_PR = '<w:sdtPr><w:tag w:val="row-control"/><w:id w:val="700000001"/></w:sdtPr>';
const SECTION_SDT_PR = '<w:sdtPr><w:id w:val="700000002"/><w15:repeatingSection/></w:sdtPr>';
const ITEM_SDT_PR = '<w:sdtPr><w:id w:val="700000003"/><w15:repeatingSectionItem/></w:sdtPr>';
const ITEM2_SDT_PR = '<w:sdtPr><w:id w:val="700000004"/><w15:repeatingSectionItem/></w:sdtPr>';

const para = (id: string, text: string) =>
  `<w:p w14:paraId="${id}"><w:r><w:t>${text}</w:t></w:r></w:p>`;
const tc = (body: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>${body}</w:tc>`;
const tr = (id: string, label: string) =>
  `<w:tr>${tc(para(`${id}1`, label))}${tc(`<w:p w14:paraId="${id}2"></w:p>`)}</w:tr>`;
const sdt = (pr: string, body: string, endPr = '') =>
  `<w:sdt>${pr}${endPr}<w:sdtContent>${body}</w:sdtContent></w:sdt>`;

/**
 * Header, a row in a control (with sdtEndPr), a row in customXml, a
 * repeating section of two items (nested sdts), a two-row control, a plain row.
 */
const TABLE =
  '<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid>' +
  tr('0000010', 'Item') +
  sdt(ROW_SDT_PR, tr('0000020', 'Wrapped row'), '<w:sdtEndPr/>') +
  '<w:customXml w:uri="urn:x" w:element="row"><w:customXmlPr><w:attr w:name="n" w:val="3"/></w:customXmlPr>' +
  tr('0000030', 'Custom row') +
  '</w:customXml>' +
  sdt(
    SECTION_SDT_PR,
    sdt(ITEM_SDT_PR, tr('0000040', 'Section item A')) +
      sdt(ITEM2_SDT_PR, tr('0000050', 'Section item B'))
  ) +
  sdt(
    '<w:sdtPr><w:tag w:val="pair"/><w:id w:val="700000005"/></w:sdtPr>',
    tr('0000060', 'Pair top') + tr('0000070', 'Pair bottom')
  ) +
  tr('0000080', 'Plain row') +
  '</w:tbl>';

const LABELS = [
  'Item',
  'Wrapped row',
  'Custom row',
  'Section item A',
  'Section item B',
  'Pair top',
  'Pair bottom',
  'Plain row',
];

const DOCUMENT_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<w:document xmlns:w="${NS_W}" xmlns:w14="${NS_W14}" xmlns:w15="${NS_W15}"><w:body>` +
  TABLE +
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

async function savedTableXml(doc: Document): Promise<string> {
  const saved = await JSZip.loadAsync(await repackDocx(doc));
  const xml = await saved.file('word/document.xml')!.async('string');
  return xml.match(/<w:tbl>.*<\/w:tbl>/s)?.[0] ?? '';
}

function table(doc: Document): Table {
  return doc.package.document!.content.find((b): b is Table => b.type === 'table')!;
}

function rowLabel(row: TableRow): string {
  const p = row.cells[0]!.content[0] as Paragraph;
  return p.content
    .flatMap((r) => (r.type === 'run' ? r.content : []))
    .map((c) => (c.type === 'text' ? c.text : ''))
    .join('');
}

describe('table-level content controls wrapping rows', () => {
  let restoreDocument: () => void = () => {};
  let bytes: Uint8Array;
  beforeAll(async () => {
    restoreDocument = installCanvasDocumentStub();
    bytes = await buildDocx();
  });
  afterAll(() => restoreDocument());

  test('every wrapped row is parsed, in document order, with its wrapper chain', async () => {
    const t = table(await parseDocxHeadless(bytes));
    expect(t.rows.map(rowLabel)).toEqual(LABELS);
    expect(t.rows.map((r) => r.cells.length)).toEqual(LABELS.map(() => 2));

    const [head, wrapped, custom, itemA, itemB, top, bottom, plain] = t.rows;
    expect(head!.tableWrappers).toBeUndefined();
    expect(plain!.tableWrappers).toBeUndefined();

    const w = wrapped!.tableWrappers!;
    expect(w).toHaveLength(1);
    if (w[0]!.kind !== 'sdt') throw new Error('expected sdt');
    expect(w[0]!.properties.tag).toBe('row-control');
    expect(w[0]!.properties.rawPropertiesXml).toBe(ROW_SDT_PR);

    expect(custom!.tableWrappers!.map((x) => x.kind)).toEqual(['customXml']);
    // Repeating section: same outer control, distinct item controls.
    const [secA, itemWA] = itemA!.tableWrappers!;
    const [secB, itemWB] = itemB!.tableWrappers!;
    expect(secA!.id).toBe(secB!.id);
    expect(itemWA!.id).not.toBe(itemWB!.id);
    // One control holding two rows.
    expect(top!.tableWrappers![0]!.id).toBe(bottom!.tableWrappers![0]!.id);
  });

  test('a save re-emits the rows and their wrappers byte-for-byte', async () => {
    expect(await savedTableXml(await parseDocxHeadless(bytes))).toBe(TABLE);
  });

  test('the editor round-trips it (Document -> PM -> Document -> XML)', async () => {
    const doc = await parseDocxHeadless(bytes);
    const pmDoc = await toProseDocHeadless(doc);
    let rows = 0;
    pmDoc.descendants((node) => {
      if (node.type.name === 'tableRow') rows++;
      return true;
    });
    expect(rows).toBe(LABELS.length);
    expect(await savedTableXml(fromProseDoc(pmDoc, doc))).toBe(TABLE);
  });

  test('the wrapped rows lay out', async () => {
    const doc = await parseDocxHeadless(bytes);
    const res = await layoutDocxHeadless(doc);
    const block = res.blocks.find((b) => b.kind === 'table') as TableBlock;
    expect(block.rows).toHaveLength(LABELS.length);
  });

  test('a row inserted inside a multi-row control does not duplicate the control', async () => {
    const doc = await parseDocxHeadless(bytes);
    const t = table(doc);
    const inserted: TableRow = { ...t.rows[7]!, tableWrappers: undefined };
    t.rows.splice(6, 0, inserted); // between "Pair top" and "Pair bottom"
    const xml = await savedTableXml(doc);
    expect(xml.match(/w:val="pair"/g)).toHaveLength(1);
    expect(xml.match(/<w:tr>/g)).toHaveLength(LABELS.length + 1);
    expect(xml).toContain('Pair bottom');
  });

  test('a save drops a tampered wrapper but keeps its row', async () => {
    const doc = await parseDocxHeadless(bytes);
    const t = table(doc);
    const sdtW = t.rows[1]!.tableWrappers![0]!;
    if (sdtW.kind !== 'sdt') throw new Error('expected sdt');
    sdtW.properties = {
      ...sdtW.properties,
      rawPropertiesXml: '<w:sdtPr/></w:sdtContent></w:sdt><w:tr><w:tc><w:p/></w:tc></w:tr><w:sdt>',
    };
    const customW = t.rows[2]!.tableWrappers![0]!;
    if (customW.kind !== 'customXml') throw new Error('expected customXml');
    customW.startXml = '<w:customXml w:element="a"></w:customXml><w:tr/><w:customXml>';

    const xml = await savedTableXml(doc);
    expect(xml).not.toContain('row-control');
    expect(xml).not.toContain('<w:customXml');
    expect(xml.match(/<w:tr>/g)).toHaveLength(LABELS.length);
    expect(xml).toContain('Wrapped row');
    expect(xml).toContain('Custom row');
  });

  test('collaboration keeps table wrappers out of the shared document', async () => {
    const doc = await parseDocxHeadless(bytes);
    const pmDoc = await toProseDocHeadless(doc);
    const projected = projectProseMirrorDocument(pmDoc, { baseRevision: 'base-1' });
    const shared = JSON.stringify(projected.document);
    expect(shared).not.toContain('"tableWrappers":[');
    expect(shared).not.toContain('row-control');

    const hydrated = schema.nodeFromJSON(
      rehydrateCollaborationDocument(projected.document, projected.sidecar)
    );
    expect(await savedTableXml(fromProseDoc(hydrated, doc))).toBe(TABLE);
  });
});

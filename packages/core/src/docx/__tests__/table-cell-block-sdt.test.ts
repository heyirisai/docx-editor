/**
 * Block-level content controls directly inside a table cell (`w:tc > w:sdt`).
 *
 * Shape under test (synthetic, the usual security-questionnaire layout): each
 * response cell holds a block `w:sdt` drop-down (Yes / No / N/A) wrapping its
 * placeholder paragraph, followed by a "Comments:" paragraph carrying an
 * inline plain-text control.
 *
 * Word shows the drop-down. The parser used to skip any `w:sdt` child of a
 * `w:tc`, so the control AND the paragraph it wraps vanished — consumers could
 * neither see nor answer it, and a save deleted it.
 *
 * A real questionnaire with cell-level drop-downs can be pointed at via
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
import { toFlowBlocks } from '../../layout-bridge/toFlowBlocks';
import type { BlockSdt, Document, Table } from '../../types/document';
import type { ParagraphBlock, TableBlock } from '../../layout-engine/types';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const DROPDOWN_SDT_PR =
  '<w:sdtPr><w:rPr><w:sz w:val="18"/></w:rPr><w:alias w:val="Yes/No"/><w:tag w:val="Answer"/>' +
  '<w:id w:val="500000002"/><w:placeholder><w:docPart w:val="00000000000000000000000000000001"/></w:placeholder>' +
  '<w:showingPlcHdr/><w:dropDownList><w:listItem w:value="Choose an item."/>' +
  '<w:listItem w:displayText="Yes" w:value="Yes"/><w:listItem w:displayText="No" w:value="No"/>' +
  '<w:listItem w:displayText="N/A" w:value="N/A"/></w:dropDownList></w:sdtPr>';

const RESPONSE_CELL =
  '<w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>' +
  `<w:sdt>${DROPDOWN_SDT_PR}<w:sdtContent>` +
  '<w:p><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>Choose an item.</w:t></w:r></w:p>' +
  '</w:sdtContent></w:sdt>' +
  '<w:p/>' +
  '<w:p><w:r><w:t xml:space="preserve">Comments: </w:t></w:r>' +
  '<w:sdt><w:sdtPr><w:id w:val="500000003"/><w:showingPlcHdr/><w:text/></w:sdtPr><w:sdtContent>' +
  '<w:r><w:t>Click or tap here to enter text.</w:t></w:r></w:sdtContent></w:sdt></w:p>' +
  '</w:tc>';

const DOCUMENT_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}"><w:body>` +
  '<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid>' +
  '<w:tr><w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>' +
  '<w:p><w:r><w:t>Do you encrypt data at rest?</w:t></w:r></w:p></w:tc>' +
  RESPONSE_CELL +
  '</w:tr></w:tbl><w:p/>' +
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

function responseCell(doc: Document) {
  const table = doc.package.document!.content[0] as Table;
  return table.rows[0]!.cells[1]!;
}

function editorState(pmDoc: PMNode): EditorState {
  return EditorState.create({ doc: pmDoc, schema });
}

describe('block content controls inside table cells', () => {
  let restoreDocument: () => void = () => {};
  let bytes: Uint8Array;
  beforeAll(async () => {
    restoreDocument = installCanvasDocumentStub();
    bytes = await buildDocx();
  });
  afterAll(() => restoreDocument());

  test('the cell keeps the drop-down control and the paragraphs after it', async () => {
    const doc = await parseDocxHeadless(bytes);
    const cell = responseCell(doc);
    expect(cell.content.map((b) => b.type)).toEqual(['blockSdt', 'paragraph', 'paragraph']);

    const sdt = cell.content[0] as BlockSdt;
    expect(sdt.properties.sdtType).toBe('dropDownList');
    expect(sdt.properties.tag).toBe('Answer');
    expect(sdt.properties.alias).toBe('Yes/No');
    expect(sdt.properties.id).toBe(500000002);
    expect(sdt.properties.showingPlaceholder).toBe(true);
    expect(sdt.properties.listItems?.map((i) => i.value)).toEqual([
      'Choose an item.',
      'Yes',
      'No',
      'N/A',
    ]);
    expect(sdt.content.map((b) => b.type)).toEqual(['paragraph']);
  });

  test('a save round-trips the cell-level control unchanged', async () => {
    const doc = await parseDocxHeadless(bytes);
    const xml = await savedDocumentXml(doc);
    expect(xml).toContain(`<w:sdt>${DROPDOWN_SDT_PR}<w:sdtContent>`);
    // Still inside the response cell, before the "Comments" paragraph.
    const tc = xml.slice(xml.lastIndexOf('<w:tc>'), xml.lastIndexOf('</w:tc>'));
    expect(tc.indexOf('<w:dropDownList>')).toBeGreaterThan(0);
    expect(tc.indexOf('<w:dropDownList>')).toBeLessThan(tc.indexOf('Comments'));
  });

  test('the editor round-trips it (Document -> PM -> Document)', async () => {
    const doc = await parseDocxHeadless(bytes);
    const pmDoc = await toProseDocHeadless(doc);
    const kinds: string[] = [];
    pmDoc.descendants((node, _pos, parent) => {
      if (parent?.type.name === 'tableCell') kinds.push(node.type.name);
      return true;
    });
    expect(kinds).toEqual(['paragraph', 'blockSdt', 'paragraph', 'paragraph']);

    const back = fromProseDoc(pmDoc, doc);
    const cell = responseCell(back);
    expect(cell.content.map((b) => b.type)).toEqual(['blockSdt', 'paragraph', 'paragraph']);
    expect((cell.content[0] as BlockSdt).properties.rawPropertiesXml).toBe(DROPDOWN_SDT_PR);
    expect(await savedDocumentXml(back)).toContain(`<w:sdt>${DROPDOWN_SDT_PR}<w:sdtContent>`);
  });

  test('the PM content-control API finds and answers it', async () => {
    const doc = await parseDocxHeadless(bytes);
    const state = editorState(await toProseDocHeadless(doc));
    const dropdowns = findContentControlsInPM(state.doc, { type: 'dropDownList' });
    expect(dropdowns).toHaveLength(1);
    expect(dropdowns[0]!.tag).toBe('Answer');
    expect(dropdowns[0]!.listItems?.map((i) => i.value)).toContain('N/A');
    expect(dropdowns[0]!.text).toBe('Choose an item.');

    const next = state.apply(
      setContentControlValueTr(state, { tag: 'Answer' }, { kind: 'dropdown', value: 'Yes' })
    );
    const answered = findContentControlsInPM(next.doc, { tag: 'Answer' })[0]!;
    expect(answered.text).toBe('Yes');
    expect(answered.showingPlaceholder).toBeUndefined();

    const xml = await savedDocumentXml(fromProseDoc(next.doc, doc));
    const tc = xml.slice(xml.lastIndexOf('<w:tc>'), xml.lastIndexOf('</w:tc>'));
    expect(tc).toContain('<w:dropDownList');
    expect(tc).not.toContain('Choose an item.</w:t>');
    expect(tc).toMatch(/<w:sdtContent><w:p>.*Yes<\/w:t>/);
    expect(tc).not.toContain('<w:showingPlcHdr/><w:dropDownList');
  });

  test('the headless content-control API finds and answers it', async () => {
    const doc = await parseDocxHeadless(bytes);
    const found = findContentControls(doc, { type: 'dropDownList' });
    expect(found).toHaveLength(1);
    expect(found[0]!.tag).toBe('Answer');

    const answered = setContentControlValue(
      doc,
      { tag: 'Answer' },
      { kind: 'dropdown', value: 'No' }
    );
    const sdt = responseCell(answered).content[0] as BlockSdt;
    expect(sdt.type).toBe('blockSdt');
    expect(findContentControls(answered, { tag: 'Answer' })[0]!.text).toBe('No');
  });

  test('the cell lays out the control content, tagged with its SDT group', async () => {
    const doc = await parseDocxHeadless(bytes);
    const blocks = toFlowBlocks(await toProseDocHeadless(doc));
    const table = blocks.find((b): b is TableBlock => b.kind === 'table')!;
    const cellBlocks = table.rows[0]!.cells[1]!.blocks as ParagraphBlock[];
    expect(cellBlocks).toHaveLength(3);
    const first = cellBlocks[0]!;
    expect(first.runs.map((r) => (r.kind === 'text' ? r.text : '')).join('')).toBe(
      'Choose an item.'
    );
    expect(first.sdtGroups?.map((g) => [g.sdtType, g.tag])).toEqual([['dropDownList', 'Answer']]);
    expect(cellBlocks[2]!.sdtGroups).toBeUndefined();
  });
});

// ---------------------------------------------------------------------------
// Optional: a real questionnaire with cell-level drop-downs (not in the repo).
// ---------------------------------------------------------------------------

const FIXTURE = process.env.DOCX_FIXTURE_CELL_CONTROLS;
const available = !!FIXTURE && fs.existsSync(FIXTURE);
if (!available) {
  console.warn(
    '[table-cell-block-sdt] real-file check SKIPPED — set DOCX_FIXTURE_CELL_CONTROLS to the ' +
      'path of a .docx with drop-downs inside table cells to run it.'
  );
}
const runFixture = available ? test : test.skip;

describe('real questionnaire (fixture) cell-level drop-downs', () => {
  let restoreDocument: () => void = () => {};
  beforeAll(() => {
    if (available) restoreDocument = installCanvasDocumentStub();
  });
  afterAll(() => restoreDocument());

  runFixture(
    'every cell-level Yes/No drop-down is parsed, editable and saved',
    async () => {
      const raw = new Uint8Array(fs.readFileSync(FIXTURE!));
      const sourceXml = await (await JSZip.loadAsync(raw))
        .file('word/document.xml')!
        .async('string');
      const expected = (sourceXml.match(/<w:dropDownList\b/g) ?? []).length;
      expect(expected).toBeGreaterThan(0);

      const doc = await parseDocxHeadless(raw);
      expect(findContentControls(doc, { type: 'dropDownList' })).toHaveLength(expected);

      const state = editorState(await toProseDocHeadless(doc));
      const dropdowns = findContentControlsInPM(state.doc, { type: 'dropDownList' });
      expect(dropdowns).toHaveLength(expected);
      for (const d of dropdowns) expect(d.listItems?.length ?? 0).toBeGreaterThan(1);
      // The answer controls in the question tables are answerable.
      const target = dropdowns.find((d) => d.id !== undefined && (d.listItems?.length ?? 0) > 1)!;
      expect(target).toBeDefined();
      const choice = target.listItems!.at(-1)!.value;
      const next = state.apply(
        setContentControlValueTr(state, { id: target.id }, { kind: 'dropdown', value: choice })
      );
      expect(findContentControlsInPM(next.doc, { id: target.id })[0]!.text).toBe(
        target.listItems!.at(-1)!.displayText
      );

      const saved = await savedDocumentXml(fromProseDoc(state.doc, doc));
      expect((saved.match(/<w:dropDownList\b/g) ?? []).length).toBe(expected);
    },
    60_000
  );
});

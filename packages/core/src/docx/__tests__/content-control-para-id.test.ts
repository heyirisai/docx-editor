/**
 * Answering a content control keeps the `w14:paraId` / `w14:textId` (and the
 * paragraph properties) of the paragraphs it rewrites, as Word does. Hosts key
 * answer anchors by paraId, so an answer that re-minted or dropped it made the
 * slot unfindable after a save.
 *
 * Covered for every control placement — inline in a paragraph, block in the
 * body, block directly in a table cell, and a row-level control wrapping a
 * whole cell — through both the headless API and the editor (PM) transactions,
 * each asserted after save + reparse. When one paragraph becomes several, the
 * first keeps the original id and the rest get fresh ids unique in the document.
 */

import { describe, test, expect } from 'bun:test';
import JSZip from 'jszip';
import { EditorState } from 'prosemirror-state';
import {
  parseDocxHeadless,
  toProseDocHeadless,
} from '../../layout-engine/integration/headlessDocxLayout';
import { schema } from '../../prosemirror/schema';
import { fromProseDoc } from '../../prosemirror/conversion/fromProseDoc';
import { repackDocx } from '../rezip';
import {
  setContentControlContentTr,
  setContentControlValueTr,
  findContentControlsInPM,
} from '../../prosemirror/contentControls';
import { setContentControlContent, type ContentControlFilter } from '../../agent/contentControls';
import { setContentControlValue, type ContentControlValue } from '../../agent/contentControlValues';
import { isValidLongHexId } from '../../utils/hexId';
import type { BlockContent, Document, Paragraph } from '../../types/document';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';

const DROPDOWN = (tag: string, id: number) =>
  `<w:sdtPr><w:tag w:val="${tag}"/><w:id w:val="${id}"/><w:showingPlcHdr/>` +
  '<w:dropDownList><w:listItem w:value="Choose an item."/>' +
  '<w:listItem w:displayText="Yes" w:value="Yes"/><w:listItem w:displayText="No" w:value="No"/>' +
  '</w:dropDownList></w:sdtPr>';
const RICH = (tag: string, id: number) =>
  `<w:sdtPr><w:tag w:val="${tag}"/><w:id w:val="${id}"/><w:showingPlcHdr/></w:sdtPr>`;
const PLAIN = (tag: string, id: number) =>
  `<w:sdtPr><w:tag w:val="${tag}"/><w:id w:val="${id}"/><w:showingPlcHdr/><w:text/></w:sdtPr>`;

/** A paragraph with identity + direct pPr (centered, 240tw after). */
const para = (paraId: string, textId: string, text: string) =>
  `<w:p w14:paraId="${paraId}" w14:textId="${textId}"><w:pPr><w:jc w:val="center"/>` +
  `<w:spacing w:after="240"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;

const blockSdt = (sdtPr: string, body: string) =>
  `<w:sdt>${sdtPr}<w:sdtContent>${body}</w:sdtContent></w:sdt>`;

const tc = (body: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${body}</w:tc>`;

const BODY =
  // Inline controls inside a paragraph.
  `<w:p w14:paraId="10000001" w14:textId="20000001"><w:pPr><w:jc w:val="right"/></w:pPr>` +
  '<w:r><w:t xml:space="preserve">Name: </w:t></w:r>' +
  `<w:sdt>${RICH('InlineText', 1)}<w:sdtContent><w:r><w:t>Enter name</w:t></w:r></w:sdtContent></w:sdt>` +
  '<w:r><w:t xml:space="preserve"> Encrypt: </w:t></w:r>' +
  `<w:sdt>${DROPDOWN('InlineDrop', 2)}<w:sdtContent><w:r><w:t>Choose an item.</w:t></w:r></w:sdtContent></w:sdt>` +
  '</w:p>' +
  // Body block controls: a two-paragraph rich text, a plain text, a dropdown.
  blockSdt(
    RICH('BodyRich', 3),
    para('10000002', '20000002', 'First') + para('10000003', '20000003', 'Second')
  ) +
  blockSdt(PLAIN('BodyPlain', 4), para('10000004', '20000004', 'Plain here')) +
  blockSdt(DROPDOWN('BodyDrop', 5), para('10000005', '20000005', 'Choose an item.')) +
  // Cell-level (w:tc > w:sdt) and row-level (w:tr > w:sdt > w:tc) controls.
  '<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
  '<w:tr>' +
  tc(para('10000006', '20000006', 'Question one')) +
  tc(blockSdt(DROPDOWN('CellDrop', 6), para('10000007', '20000007', 'Choose an item.'))) +
  '</w:tr><w:tr>' +
  tc(para('10000008', '20000008', 'Question two')) +
  blockSdt(RICH('RowRich', 7), tc(para('10000009', '20000009', 'Click here'))) +
  '</w:tr><w:tr>' +
  tc(para('1000000A', '2000000A', 'Question three')) +
  blockSdt(DROPDOWN('RowDrop', 8), tc(para('1000000B', '2000000B', 'Choose an item.'))) +
  '</w:tr></w:tbl>' +
  '<w:p w14:paraId="1000000C"/>';

const DOCUMENT_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<w:document xmlns:w="${NS_W}" xmlns:w14="${NS_W14}"><w:body>${BODY}` +
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

/** Save + reparse; returns the reparsed document and the saved document.xml. */
async function roundTrip(doc: Document): Promise<{ doc: Document; xml: string }> {
  const bytes = new Uint8Array(await repackDocx(doc));
  const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
  return { doc: await parseDocxHeadless(bytes), xml };
}

/** Every paragraph in body order (through controls and table cells). */
function allParagraphs(doc: Document): Paragraph[] {
  const out: Paragraph[] = [];
  const walk = (blocks: readonly BlockContent[]) => {
    for (const b of blocks) {
      if (b.type === 'paragraph') out.push(b);
      else if (b.type === 'blockSdt') walk(b.content);
      else if (b.type === 'table')
        for (const row of b.rows) for (const cell of row.cells) walk(cell.content);
    }
  };
  walk(doc.package.document!.content);
  return out;
}

function paragraphText(p: Paragraph): string {
  const out: string[] = [];
  const walk = (content: readonly unknown[]) => {
    for (const node of content as Array<{ type: string; content?: unknown[]; text?: string }>) {
      if (node.type === 'text') out.push(node.text ?? '');
      else if (node.content) walk(node.content);
    }
  };
  walk(p.content);
  return out.join('');
}

function byParaId(doc: Document, paraId: string): Paragraph | undefined {
  return allParagraphs(doc).find((p) => p.paraId === paraId);
}

/** Every paragraph has a valid paraId and no two share one. */
function expectUniqueValidIds(doc: Document): void {
  const ids = allParagraphs(doc).map((p) => p.paraId);
  for (const id of ids) {
    expect(id).toBeDefined();
    expect(isValidLongHexId(id)).toBe(true);
    expect(parseInt(id!, 16)).toBeLessThan(0x80000000);
  }
  expect(new Set(ids).size).toBe(ids.length);
}

type Answer =
  | { kind: 'content'; filter: ContentControlFilter; text: string }
  | { kind: 'value'; filter: ContentControlFilter; value: ContentControlValue };

/** The headless path. */
function answerHeadless(doc: Document, answers: Answer[]): Document {
  return answers.reduce(
    (d, a) =>
      a.kind === 'content'
        ? setContentControlContent(d, a.filter, a.text)
        : setContentControlValue(d, a.filter, a.value),
    doc
  );
}

/** The editor path: plugin-free PM state (no paraId allocator), then save. */
async function answerInEditor(doc: Document, answers: Answer[]): Promise<Document> {
  let state = EditorState.create({ doc: await toProseDocHeadless(doc), schema });
  for (const a of answers) {
    state = state.apply(
      a.kind === 'content'
        ? setContentControlContentTr(state, a.filter, a.text)
        : setContentControlValueTr(state, a.filter, a.value)
    );
  }
  return fromProseDoc(state.doc, doc);
}

const PATHS = [
  ['headless', async (doc: Document, answers: Answer[]) => answerHeadless(doc, answers)],
  ['editor', answerInEditor],
] as const;

const SINGLE_ANSWERS: Answer[] = [
  { kind: 'content', filter: { tag: 'InlineText' }, text: 'Acme' },
  { kind: 'value', filter: { tag: 'InlineDrop' }, value: { kind: 'dropdown', value: 'Yes' } },
  { kind: 'content', filter: { tag: 'BodyRich' }, text: 'One answer' },
  { kind: 'content', filter: { tag: 'BodyPlain' }, text: 'Plain answer' },
  { kind: 'value', filter: { tag: 'BodyDrop' }, value: { kind: 'dropdown', value: 'No' } },
  { kind: 'value', filter: { tag: 'CellDrop' }, value: { kind: 'dropdown', value: 'Yes' } },
  { kind: 'value', filter: { tag: 'RowRich' }, value: { kind: 'text', text: 'Row answer' } },
  { kind: 'value', filter: { tag: 'RowDrop' }, value: { kind: 'dropdown', value: 'No' } },
];

describe.each(PATHS)('answering content controls keeps paragraph ids (%s)', (_name, answer) => {
  test('every placement keeps its paraId, textId and pPr through answer + save + reparse', async () => {
    const source = await parseDocxHeadless(await buildDocx());
    const { doc, xml } = await roundTrip(await answer(source, SINGLE_ANSWERS));

    const expectAnswer = (paraId: string, textId: string, text: string, jc = 'center') => {
      const p = byParaId(doc, paraId);
      expect(p).toBeDefined();
      expect(paragraphText(p!)).toContain(text);
      expect(p!.textId).toBe(textId);
      expect(p!.formatting?.alignment as string | undefined).toBe(jc);
      expect(xml).toContain(`w14:paraId="${paraId}" w14:textId="${textId}"`);
    };
    // Inline controls: the host paragraph is untouched.
    expectAnswer('10000001', '20000001', 'Name: Acme Encrypt: Yes', 'right');
    // Body block controls (the two-paragraph control answered with one line
    // keeps its first paragraph; the second goes away with its content).
    expectAnswer('10000002', '20000002', 'One answer');
    expect(byParaId(doc, '10000003')).toBeUndefined();
    expectAnswer('10000004', '20000004', 'Plain answer');
    expectAnswer('10000005', '20000005', 'No');
    // Cell-level and row-level controls.
    expectAnswer('10000007', '20000007', 'Yes');
    expectAnswer('10000009', '20000009', 'Row answer');
    expectAnswer('1000000B', '2000000B', 'No');
    expect(doc.package.document!.content.find((b) => b.type === 'table')).toBeDefined();
    expect(xml).toContain('<w:sdtContent><w:tc>');

    expectUniqueValidIds(doc);
  });

  test('a multi-line answer keeps the first paraId and mints unique ids for the rest', async () => {
    const source = await parseDocxHeadless(await buildDocx());
    const answered = await answer(source, [
      { kind: 'content', filter: { tag: 'BodyRich' }, text: 'Line A\nLine B\nLine C\nLine D' },
      { kind: 'content', filter: { tag: 'RowRich' }, text: 'Row A\nRow B' },
    ]);
    const { doc, xml } = await roundTrip(answered);
    const paragraphs = allParagraphs(doc);
    const at = (text: string) => paragraphs.find((p) => paragraphText(p) === text)!;

    // The Nth line takes over the Nth original paragraph.
    expect(at('Line A').paraId).toBe('10000002');
    expect(at('Line A').textId).toBe('20000002');
    expect(at('Line B').paraId).toBe('10000003');
    expect(at('Line B').textId).toBe('20000003');
    expect(at('Row A').paraId).toBe('10000009');
    // Lines past the original count: fresh ids, continued pPr.
    const fresh = [at('Line C'), at('Line D'), at('Row B')];
    const originals = new Set(allParagraphs(source).map((p) => p.paraId));
    for (const p of fresh) {
      expect(originals.has(p.paraId)).toBe(false);
      expect(p.textId).toBeDefined();
      expect(p.formatting?.alignment).toBe('center');
      expect(p.formatting?.spaceAfter).toBe(240);
      expect(xml).toContain(`w14:paraId="${p.paraId}"`);
    }
    expectUniqueValidIds(doc);
  });

  test('a second answer to the same control still finds the same paraId', async () => {
    const source = await parseDocxHeadless(await buildDocx());
    const once = (await roundTrip(await answer(source, SINGLE_ANSWERS))).doc;
    const twice = (
      await roundTrip(
        await answer(once, [
          { kind: 'value', filter: { tag: 'CellDrop' }, value: { kind: 'dropdown', value: 'No' } },
          { kind: 'content', filter: { tag: 'RowRich' }, text: 'Revised' },
        ])
      )
    ).doc;
    expect(paragraphText(byParaId(twice, '10000007')!)).toBe('No');
    expect(paragraphText(byParaId(twice, '10000009')!)).toBe('Revised');
    expectUniqueValidIds(twice);
  });
});

describe('explicit block replacements (headless)', () => {
  test('caller paragraphs without ids take the originals; their own ids and pPr win', async () => {
    const source = await parseDocxHeadless(await buildDocx());
    const answered = setContentControlContent(source, { tag: 'BodyRich' }, [
      { type: 'paragraph', content: [{ type: 'run', content: [{ type: 'text', text: 'X' }] }] },
      {
        type: 'paragraph',
        paraId: '0ABCDEF0',
        formatting: { alignment: 'left' },
        content: [{ type: 'run', content: [{ type: 'text', text: 'Y' }] }],
      },
      { type: 'paragraph', content: [{ type: 'run', content: [{ type: 'text', text: 'Z' }] }] },
    ]);
    const { doc } = await roundTrip(answered);
    const at = (text: string) => allParagraphs(doc).find((p) => paragraphText(p) === text)!;
    expect(at('X').paraId).toBe('10000002');
    expect(at('X').textId).toBe('20000002');
    expect(at('X').formatting?.alignment).toBeUndefined();
    expect(at('Y').paraId).toBe('0ABCDEF0');
    expect(at('Y').formatting?.alignment).toBe('left');
    expect(at('Z').paraId).toBeDefined();
    expect(['10000002', '10000003', '0ABCDEF0']).not.toContain(at('Z').paraId);
    expectUniqueValidIds(doc);
  });

  test('the PM state keeps the ids before any save', async () => {
    const source = await parseDocxHeadless(await buildDocx());
    let state = EditorState.create({ doc: await toProseDocHeadless(source), schema });
    state = state.apply(
      setContentControlValueTr(state, { tag: 'RowDrop' }, { kind: 'dropdown', value: 'Yes' })
    );
    const pos = findContentControlsInPM(state.doc, { tag: 'RowDrop' })[0]!.pos;
    const control = state.doc.nodeAt(pos)!;
    expect(control.firstChild!.attrs.paraId).toBe('1000000B');
    expect(control.firstChild!.attrs.textId).toBe('2000000B');
    expect(control.firstChild!.attrs.alignment).toBe('center');
  });
});

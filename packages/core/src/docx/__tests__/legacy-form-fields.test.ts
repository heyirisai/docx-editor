/**
 * Legacy Word form fields (`w:fldChar` + `w:ffData`) end to end.
 *
 * Phase 1 runs over `e2e/fixtures/multi-column-controls.docx`, which mixes one
 * `FORMDROPDOWN`, two `FORMCHECKBOX`, `w14:checkbox` SDTs, `w:dropDownList`
 * SDTs and hand-typed ☐ glyphs — the exact shape of the requirements matrices
 * Iris has to stage (the Certinia RFP is 93 `FORMDROPDOWN` fields and no
 * content controls at all).
 *
 * Phase 2 uses a minimal in-test fixture whose dropdown has a preset
 * `w:result` and whose checkbox has `w:default w:val="1"`, covering the
 * pre-answered documents the real matrices ship as.
 */

import { describe, test, expect } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import JSZip from 'jszip';

import { parseDocumentBody } from '../documentParser';
import { serializeDocumentBody } from '../serializer/documentSerializer';
import { toProseDoc } from '../../prosemirror/conversion/toProseDoc';
import { fromProseDoc } from '../../prosemirror/conversion/fromProseDoc';
import { EditorState } from 'prosemirror-state';

import {
  findContentControls,
  findContentControl,
  setContentControlContent,
} from '../../agent/contentControls';
import { setContentControlValue } from '../../agent/contentControlValues';
import { findGlyphCheckboxes } from '../../agent/glyphCheckboxes';
import {
  LEGACY_TEXT_PLACEHOLDER,
  setLegacyCheckbox,
  setLegacyDropdownIndex,
} from '../legacyFormField';
import type { ContentControlInfo } from '../../agent/contentControls';
import type { Document, DocumentBody, LegacyFormField, Run } from '../../types/document';

const FIXTURE = join(import.meta.dir, '../../../../../e2e/fixtures/multi-column-controls.docx');

async function fixtureXml(): Promise<string> {
  const zip = await JSZip.loadAsync(readFileSync(FIXTURE));
  const doc = zip.file('word/document.xml');
  if (!doc) throw new Error('fixture missing word/document.xml');
  return doc.async('string');
}

/** Wrap a parsed body as a `Document` so the headless agent API accepts it. */
function asDocument(body: DocumentBody): Document {
  return { package: { document: body } } as unknown as Document;
}

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';

/** `serializeDocumentBody` emits body children only — re-wrap to re-parse. */
function reparse(bodyXml: string): DocumentBody {
  return parseDocumentBody(`<w:document ${NS}><w:body>${bodyXml}</w:body></w:document>`);
}

const legacyOf = (doc: Document | DocumentBody): ContentControlInfo[] =>
  findContentControls(doc, { source: 'legacy' });

// The two legacy sequences as they appear in the fixture, byte for byte.
const CHECK1_XML =
  '<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Check1"/><w:enabled/>' +
  '<w:calcOnExit w:val="0"/><w:checkBox><w:sizeAuto/><w:default w:val="0"/></w:checkBox>' +
  '</w:ffData></w:fldChar></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> FORMCHECKBOX </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>';

const DROP1_XML =
  '<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Drop1"/><w:enabled/>' +
  '<w:calcOnExit w:val="0"/><w:ddList><w:result w:val="0"/><w:listEntry w:val="Never"/>' +
  '<w:listEntry w:val="Sometimes"/><w:listEntry w:val="Always"/></w:ddList>' +
  '</w:ffData></w:fldChar></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> FORMDROPDOWN </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:r><w:t>Never</w:t></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>';

describe('legacy form fields — parsing the multi-column-controls fixture', () => {
  test('recognizes one FORMDROPDOWN and two FORMCHECKBOX fields', async () => {
    const legacy = legacyOf(parseDocumentBody(await fixtureXml()));
    expect(legacy.map((c) => c.legacyFormField!.fieldType)).toEqual([
      'checkbox',
      'checkbox',
      'dropdown',
    ]);
    expect(legacy.map((c) => c.legacyFormField!.name)).toEqual(['Check1', 'Check2', 'Drop1']);
    // The w:ffData name is projected as tag/alias so `{ tag }` addresses them.
    expect(legacy.map((c) => c.tag)).toEqual(['Check1', 'Check2', 'Drop1']);
    for (const c of legacy) expect(c.legacyFormField!.kind).toBe('legacy');
  });

  test('projects sdtType, list entries, selection and checkbox state', async () => {
    const body = parseDocumentBody(await fixtureXml());
    const drop = findContentControl(body, { tag: 'Drop1' })!;
    expect(drop.source).toBe('legacy');
    expect(drop.sdtType).toBe('dropDownList');
    expect(drop.legacyFormField!.options).toEqual(['Never', 'Sometimes', 'Always']);
    expect(drop.legacyFormField!.selectedIndex).toBe(0);
    expect(drop.legacyFormField!.value).toBe('Never');
    // listItems mirrors the entries so the shared dropdown UI works unchanged.
    expect(drop.listItems).toEqual([
      { displayText: 'Never', value: 'Never' },
      { displayText: 'Sometimes', value: 'Sometimes' },
      { displayText: 'Always', value: 'Always' },
    ]);

    const check = findContentControl(body, { tag: 'Check1' })!;
    expect(check.sdtType).toBe('checkbox');
    expect(check.checked).toBe(false);
    expect(check.legacyFormField!.sizeAuto).toBe(true);
    // Word writes no result run for a checkbox; the parser synthesizes the
    // glyph so the painted page (and the clickable widget) match Word.
    expect(check.text).toBe('☐');
  });

  test('captures the raw ffData verbatim for the round trip', async () => {
    const drop = findContentControl(parseDocumentBody(await fixtureXml()), { tag: 'Drop1' })!;
    expect(drop.legacyFormField!.ffDataXml).toBe(
      '<w:ffData><w:name w:val="Drop1"/><w:enabled/><w:calcOnExit w:val="0"/>' +
        '<w:ddList><w:result w:val="0"/><w:listEntry w:val="Never"/>' +
        '<w:listEntry w:val="Sometimes"/><w:listEntry w:val="Always"/></w:ddList></w:ffData>'
    );
    expect(drop.legacyFormField!.hasSeparate).toBe(true);
    expect(drop.legacyFormField!.hasResult).toBe(true);
  });

  test('mixes with w:sdt controls in one list, discriminated by `source`', async () => {
    const all = findContentControls(parseDocumentBody(await fixtureXml()));
    expect(all.filter((c) => c.source === 'legacy').length).toBe(3);
    // 6 w14:checkbox SDTs + 2 w:dropDownList SDTs in the fixture.
    expect(all.filter((c) => c.source === 'sdt').length).toBe(8);
    expect(all.every((c) => (c.source === 'legacy') === (c.legacyFormField != null))).toBe(true);
  });
});

describe('legacy form fields — serialization', () => {
  test('an untouched field round-trips byte for byte', async () => {
    const out = serializeDocumentBody(parseDocumentBody(await fixtureXml()));
    expect(out).toContain(CHECK1_XML);
    expect(out).toContain(DROP1_XML);
    // Never re-encoded as a w:sdt: the fixture's 8 real SDTs and no more.
    expect((out.match(/<w:sdt>/g) ?? []).length).toBe(8);
  });

  test('survives the editor round trip (parse → PM → back → serialize)', async () => {
    const body = parseDocumentBody(await fixtureXml());
    const out = serializeDocumentBody(fromProseDoc(toProseDoc(asDocument(body))).package.document);
    expect(out).toContain(DROP1_XML);
    expect(out).toContain('<w:ffData><w:name w:val="Check1"/>');
    expect(out).toContain(' FORMCHECKBOX ');
    // The synthesized checkbox glyph must not leak into the saved bytes: the
    // separator is still followed straight by the field end.
    expect(out).toContain(
      ' FORMCHECKBOX </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    );
    expect((out.match(/<w:sdt>/g) ?? []).length).toBe(8);
  });

  test('keeps w14:paraId on the enclosing paragraph', async () => {
    const body = parseDocumentBody(await fixtureXml());
    expect(serializeDocumentBody(body)).toContain('w14:paraId="1A2B3C4D"');
    const edited = fromProseDoc(toProseDoc(asDocument(body))).package.document;
    expect(serializeDocumentBody(edited)).toContain('w14:paraId="1A2B3C4D"');
  });

  test('a malformed sequence falls back to the opaque complex-field passthrough', () => {
    // ffData says dropdown, the instruction says FORMTEXT — refusing to model
    // it keeps today's behaviour instead of corrupting the field on save.
    const xml =
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>' +
      '<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Bad"/>' +
      '<w:ddList><w:listEntry w:val="A"/></w:ddList></w:ffData></w:fldChar></w:r>' +
      '<w:r><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>A</w:t></w:r>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:body></w:document>';
    const body = parseDocumentBody(xml);
    expect(legacyOf(body)).toEqual([]);
    const para = body.content[0];
    expect(para.type === 'paragraph' && para.content[0].type).toBe('complexField');
    expect(serializeDocumentBody(body)).toContain('FORMTEXT');
  });
});

describe('legacy form fields — setContentControlValue', () => {
  test('dropdown: writes w:result and the displayed run, then re-parses', async () => {
    const doc = asDocument(parseDocumentBody(await fixtureXml()));
    const next = setContentControlValue(
      doc,
      { tag: 'Drop1' },
      { kind: 'dropdown', value: 'Always' }
    );

    const info = findContentControl(next, { tag: 'Drop1' })!;
    expect(info.legacyFormField!.selectedIndex).toBe(2);
    expect(info.legacyFormField!.value).toBe('Always');
    expect(info.text).toBe('Always');

    const out = serializeDocumentBody(next.package.document);
    expect(out).toContain('<w:result w:val="2"/>');
    expect(out).toContain('<w:t>Always</w:t>');
    // Everything else inside ffData is replayed verbatim.
    expect(out).toContain('<w:name w:val="Drop1"/><w:enabled/><w:calcOnExit w:val="0"/>');
    expect(out).toContain('<w:listEntry w:val="Sometimes"/>');

    const reparsed = findContentControl(reparse(out), { tag: 'Drop1' })!;
    expect(reparsed.legacyFormField!.selectedIndex).toBe(2);
    expect(reparsed.legacyFormField!.value).toBe('Always');
    expect(reparsed.text).toBe('Always');
  });

  test('checkbox: writes both w:default and w:checked, then re-parses', async () => {
    const doc = asDocument(parseDocumentBody(await fixtureXml()));
    const next = setContentControlValue(
      doc,
      { tag: 'Check1' },
      { kind: 'checkbox', checked: true }
    );

    expect(findContentControl(next, { tag: 'Check1' })!.checked).toBe(true);
    expect(findContentControl(next, { tag: 'Check1' })!.text).toBe('☒');

    const out = serializeDocumentBody(next.package.document);
    // Word draws the box from w:checked, so the whole sequence is the captured
    // one with only the ffData rewritten — no result run is added for the glyph
    // (which would double it in Word).
    expect(out).toContain(
      CHECK1_XML.replace('<w:default w:val="0"/>', '<w:default w:val="1"/><w:checked w:val="1"/>')
    );
    // The untouched sibling still serializes byte for byte.
    expect(out).toContain(CHECK1_XML.replace(/Check1/, 'Check2'));

    const reparsed = findContentControl(reparse(out), { tag: 'Check1' })!;
    expect(reparsed.checked).toBe(true);
    expect(findContentControl(reparse(out), { tag: 'Check2' })!.checked).toBe(false);
  });

  test('rejects a value kind the field cannot hold', async () => {
    const doc = asDocument(parseDocumentBody(await fixtureXml()));
    expect(() =>
      setContentControlValue(doc, { tag: 'Drop1' }, { kind: 'checkbox', checked: true })
    ).toThrow(/not a checkbox/);
    expect(() =>
      setContentControlValue(doc, { tag: 'Check1' }, { kind: 'dropdown', value: 'Always' })
    ).toThrow(/not a dropdown/);
    expect(() =>
      setContentControlValue(doc, { tag: 'Drop1' }, { kind: 'dropdown', value: 'Nope' })
    ).toThrow(/not one of the field's list entries/);
  });
});

describe('legacy form fields — pre-set values (minimal generated fixture)', () => {
  /**
   * Second fixture, built here rather than checked in: a dropdown already
   * answered (`w:result w:val="1"`) and a checkbox whose only state is
   * `w:default w:val="1"` — the shape a matrix arrives in when the customer
   * pre-filled some rows, and the case where `w:checked` has to be *added*
   * rather than overwritten.
   */
  const PRESET =
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' +
    '<w:p><w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Pre1"/><w:enabled/>' +
    '<w:ddList><w:result w:val="1"/><w:listEntry w:val="FS"/><w:listEntry w:val="SC"/>' +
    '<w:listEntry w:val="NS"/></w:ddList></w:ffData></w:fldChar></w:r>' +
    '<w:r><w:instrText xml:space="preserve"> FORMDROPDOWN </w:instrText></w:r>' +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>SC</w:t></w:r>' +
    '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>' +
    '<w:p><w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Pre2"/>' +
    '<w:checkBox><w:sizeAuto/><w:default w:val="1"/></w:checkBox></w:ffData></w:fldChar></w:r>' +
    '<w:r><w:instrText xml:space="preserve"> FORMCHECKBOX </w:instrText></w:r>' +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
    '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>' +
    '</w:body></w:document>';

  test('reads a preset w:result and a w:default-only checkbox', () => {
    const body = parseDocumentBody(PRESET);
    const drop = findContentControl(body, { tag: 'Pre1' })!;
    expect(drop.legacyFormField!.selectedIndex).toBe(1);
    expect(drop.legacyFormField!.value).toBe('SC');
    // No w:checked yet — w:default is the state Word renders.
    expect(findContentControl(body, { tag: 'Pre2' })!.checked).toBe(true);
    expect(findContentControl(body, { tag: 'Pre2' })!.text).toBe('☒');
  });

  test('re-selecting overwrites w:result in place', () => {
    const next = setContentControlValue(
      asDocument(parseDocumentBody(PRESET)),
      { tag: 'Pre1' },
      {
        kind: 'dropdown',
        value: 'NS',
      }
    );
    const out = serializeDocumentBody(next.package.document);
    expect(out).toContain('<w:ddList><w:result w:val="2"/><w:listEntry w:val="FS"/>');
    expect(out).toContain('<w:t>NS</w:t>');
    expect(out).not.toContain('<w:t>SC</w:t>');
  });

  test('unticking a w:default-only checkbox inserts w:checked in sequence order', () => {
    const next = setContentControlValue(
      asDocument(parseDocumentBody(PRESET)),
      { tag: 'Pre2' },
      {
        kind: 'checkbox',
        checked: false,
      }
    );
    const out = serializeDocumentBody(next.package.document);
    expect(out).toContain(
      '<w:checkBox><w:sizeAuto/><w:default w:val="0"/><w:checked w:val="0"/></w:checkBox>'
    );
    expect(findContentControl(reparse(out), { tag: 'Pre2' })!.checked).toBe(false);
  });

  test('a field with no separator still round-trips, and gains one when answered', () => {
    const noSep =
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' +
      '<w:p><w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="T1"/>' +
      '<w:textInput/></w:ffData></w:fldChar></w:r>' +
      '<w:r><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:body></w:document>';
    const body = parseDocumentBody(noSep);
    const field = findContentControl(body, { tag: 'T1' })!;
    expect(field.sdtType).toBe('plainText');
    expect(field.legacyFormField!.hasSeparate).toBe(false);
    expect(serializeDocumentBody(body)).not.toContain('separate');

    const next = setContentControlValue(
      asDocument(body),
      { tag: 'T1' },
      {
        kind: 'text',
        text: 'Yes, since 2019.',
      }
    );
    const out = serializeDocumentBody(next.package.document);
    expect(out).toContain('<w:fldChar w:fldCharType="separate"/>');
    expect(out).toContain('<w:t>Yes, since 2019.</w:t>');
    expect(findContentControl(reparse(out), { tag: 'T1' })!.text).toBe('Yes, since 2019.');
  });
});

describe('legacy form fields — display synthesis for result-less fields', () => {
  /**
   * The shape the 6sense / Certinia "PSA Platform Evaluation RFP" writes 93
   * times in its Response Code column: sized structural runs, an empty code
   * run, no `w:ddList/w:result`, and `separate` followed straight by `end`.
   * Word displays the current entry (index 0) in every such cell; before the
   * synthesis the cells rendered empty.
   */
  const RPR = '<w:rPr><w:sz w:val="13"/><w:szCs w:val="13"/></w:rPr>';
  const ENTRIES =
    '<w:listEntry w:val="FS"/><w:listEntry w:val="SC"/><w:listEntry w:val="SX"/>' +
    '<w:listEntry w:val="TP"/><w:listEntry w:val="NS"/><w:listEntry w:val="RM"/>';
  const dropdownXml = (ddList: string, result = '', name = 'Dropdown1') =>
    `<w:r>${RPR}<w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="${name}"/><w:enabled/>` +
    `<w:calcOnExit w:val="0"/><w:ddList>${ddList}</w:ddList></w:ffData></w:fldChar></w:r>` +
    `<w:r>${RPR}<w:instrText xml:space="preserve"> FORMDROPDOWN </w:instrText></w:r>` +
    `<w:r>${RPR}</w:r>` +
    `<w:r>${RPR}<w:fldChar w:fldCharType="separate"/></w:r>` +
    result +
    `<w:r>${RPR}<w:fldChar w:fldCharType="end"/></w:r>`;
  const textXml = (textInput: string, name: string) =>
    `<w:r>${RPR}<w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="${name}"/>` +
    (textInput ? `<w:textInput>${textInput}</w:textInput>` : '<w:textInput/>') +
    `</w:ffData></w:fldChar></w:r>` +
    `<w:r>${RPR}<w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r>` +
    `<w:r>${RPR}<w:fldChar w:fldCharType="separate"/></w:r>` +
    `<w:r>${RPR}<w:fldChar w:fldCharType="end"/></w:r>`;
  const checkboxXml = (checkBox: string, name: string) =>
    `<w:r>${RPR}<w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="${name}"/>` +
    `<w:checkBox><w:sizeAuto/>${checkBox}</w:checkBox></w:ffData></w:fldChar></w:r>` +
    `<w:r>${RPR}<w:instrText xml:space="preserve"> FORMCHECKBOX </w:instrText></w:r>` +
    `<w:r>${RPR}<w:fldChar w:fldCharType="separate"/></w:r>` +
    `<w:r>${RPR}<w:fldChar w:fldCharType="end"/></w:r>`;
  const doc = (...paragraphs: string[]) =>
    `<w:document ${NS}><w:body>${paragraphs.map((p) => `<w:p>${p}</w:p>`).join('')}</w:body></w:document>`;

  /** The inline SDT that is the sole content of body paragraph `i`. */
  function fieldNode(body: DocumentBody, i = 0) {
    const para = body.content[i];
    if (para.type !== 'paragraph') throw new Error('expected a paragraph');
    const node = para.content.find((c) => c.type === 'inlineSdt');
    if (!node || node.type !== 'inlineSdt') throw new Error('expected an inlineSdt');
    return node;
  }

  test('a FORMDROPDOWN with no w:result and no result run displays entry 0', () => {
    const body = parseDocumentBody(doc(dropdownXml(ENTRIES) + `<w:r>${RPR}<w:t>x</w:t></w:r>`));
    const field = findContentControl(body, { tag: 'Dropdown1' })!;
    expect(field.text).toBe('FS');
    expect(field.legacyFormField).toMatchObject({
      selectedIndex: 0,
      value: 'FS',
      hasResult: false,
      hasSeparate: true,
    });

    // The synthesized display run inherits the field's run formatting — the
    // same `w:rPr` a regular run in that paragraph would get.
    const node = fieldNode(body);
    expect(node.content).toHaveLength(1);
    const display = node.content[0] as Run;
    const sibling = (body.content[0] as { content: Run[] }).content.at(-1)!;
    expect(display.type).toBe('run');
    expect(sibling.formatting).toBeDefined();
    expect(display.formatting).toEqual(sibling.formatting);
  });

  test('w:result selects the displayed entry (index 2 → third entry, 0 → first)', () => {
    const body = parseDocumentBody(
      doc(
        dropdownXml(`<w:result w:val="2"/>${ENTRIES}`, '', 'Two'),
        dropdownXml(`<w:result w:val="0"/>${ENTRIES}`, '', 'Zero')
      )
    );
    expect(findContentControl(body, { tag: 'Two' })!.text).toBe('SX');
    expect(findContentControl(body, { tag: 'Two' })!.legacyFormField!.value).toBe('SX');
    expect(findContentControl(body, { tag: 'Zero' })!.text).toBe('FS');
  });

  test('an explicit result run is displayed unchanged', () => {
    const body = parseDocumentBody(doc(dropdownXml(ENTRIES, `<w:r>${RPR}<w:t>TP</w:t></w:r>`)));
    const field = findContentControl(body, { tag: 'Dropdown1' })!;
    expect(field.text).toBe('TP');
    expect(field.legacyFormField!.hasResult).toBe(true);
    expect(fieldNode(body).content).toHaveLength(1);
  });

  test('a FORMTEXT displays w:textInput/w:default, else the five en-space blank', () => {
    const body = parseDocumentBody(
      doc(textXml('<w:default w:val="Enter vendor name"/>', 'Named'), textXml('', 'Blank'))
    );
    const named = findContentControl(body, { tag: 'Named' })!;
    expect(named.text).toBe('Enter vendor name');
    expect(named.legacyFormField).toMatchObject({
      defaultText: 'Enter vendor name',
      value: 'Enter vendor name',
      hasResult: false,
    });

    const blank = findContentControl(body, { tag: 'Blank' })!;
    expect(blank.text).toBe(LEGACY_TEXT_PLACEHOLDER);
    expect(LEGACY_TEXT_PLACEHOLDER).toBe(' '.repeat(5));
    // The blank is a display artefact, not an answer.
    expect(blank.legacyFormField!.value).toBe('');
    expect(blank.legacyFormField!.defaultText).toBeUndefined();
  });

  test('a FORMCHECKBOX displays the box from w:checked, falling back to w:default', () => {
    const body = parseDocumentBody(
      doc(
        checkboxXml('<w:default w:val="0"/><w:checked w:val="1"/>', 'Ticked'),
        checkboxXml('<w:default w:val="1"/>', 'DefaultOn'),
        checkboxXml('', 'Bare')
      )
    );
    expect(findContentControl(body, { tag: 'Ticked' })!.text).toBe('☒');
    expect(findContentControl(body, { tag: 'DefaultOn' })!.text).toBe('☒');
    expect(findContentControl(body, { tag: 'Bare' })!.text).toBe('☐');
    expect(fieldNode(body, 2).content[0]).toMatchObject({ type: 'run', formatting: {} });
  });

  test('untouched result-less fields round-trip byte for byte, directly and via the editor', () => {
    const sequences = [
      dropdownXml(ENTRIES, '', 'D'),
      textXml('<w:default w:val="Enter vendor name"/>', 'T1'),
      textXml('', 'T2'),
      checkboxXml('<w:default w:val="0"/>', 'C'),
    ];
    const body = parseDocumentBody(doc(...sequences));

    const direct = serializeDocumentBody(body);
    const viaEditor = serializeDocumentBody(
      fromProseDoc(toProseDoc(asDocument(body))).package.document
    );
    for (const out of [direct, viaEditor]) {
      for (const seq of sequences) expect(out).toContain(seq);
      // The synthesized display never leaks into the saved bytes.
      expect(out).not.toContain('<w:t>FS</w:t>');
      expect(out).not.toContain('Enter vendor name</w:t>');
      expect(out).not.toContain(LEGACY_TEXT_PLACEHOLDER);
      expect(out).not.toContain('☐');
      expect(out).not.toContain('<w:sdt>');
    }
  });

  test('an empty result run is replayed verbatim and the display is still synthesized', () => {
    const seq = dropdownXml(ENTRIES, `<w:r>${RPR}</w:r>`);
    const body = parseDocumentBody(doc(seq));
    const field = findContentControl(body, { tag: 'Dropdown1' })!;
    expect(field.text).toBe('FS');
    expect(field.legacyFormField!.hasResult).toBe(false);
    expect(serializeDocumentBody(body)).toContain(seq);
  });

  test('setContentControlValue on a result-less dropdown writes w:result and the display run', () => {
    const next = setContentControlValue(
      asDocument(parseDocumentBody(doc(dropdownXml(ENTRIES)))),
      { tag: 'Dropdown1' },
      { kind: 'dropdown', value: 'NS' }
    );
    const info = findContentControl(next, { tag: 'Dropdown1' })!;
    expect(info.text).toBe('NS');
    expect(info.legacyFormField).toMatchObject({ selectedIndex: 4, value: 'NS', hasResult: true });

    const out = serializeDocumentBody(next.package.document);
    // `w:result` is inserted first in `w:ddList` (sequence order), the rest of
    // the captured prefix — including the empty code run — is replayed as-is.
    expect(out).toContain(`<w:ddList><w:result w:val="4"/>${ENTRIES}</w:ddList>`);
    expect(out).toContain(
      ` FORMDROPDOWN </w:instrText></w:r><w:r>${RPR}</w:r><w:r>${RPR}<w:fldChar w:fldCharType="separate"/></w:r>`
    );
    expect(out).toContain('<w:t>NS</w:t>');
    expect(out).not.toContain('<w:t>FS</w:t>');
    expect(findContentControl(reparse(out), { tag: 'Dropdown1' })!.text).toBe('NS');
  });

  test('setContentControlValue on a blank FORMTEXT replaces the placeholder with a real result', () => {
    const next = setContentControlValue(
      asDocument(parseDocumentBody(doc(textXml('', 'T')))),
      { tag: 'T' },
      { kind: 'text', text: 'Yes' }
    );
    expect(findContentControl(next, { tag: 'T' })!.legacyFormField!.hasResult).toBe(true);
    const out = serializeDocumentBody(next.package.document);
    expect(out).toContain('<w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Yes</w:t></w:r>');
    expect(out).not.toContain(LEGACY_TEXT_PLACEHOLDER);
    expect(findContentControl(reparse(out), { tag: 'T' })!.text).toBe('Yes');
  });

  test('content typed over the synthesized display is kept as a real result', () => {
    const body = parseDocumentBody(doc(dropdownXml(ENTRIES)));
    // Simulate the editor replacing the display run's text without going
    // through the typed setter (so `hasResult` is still false).
    fieldNode(body).content = [{ type: 'run', content: [{ type: 'text', text: 'RM' }] }];
    const out = serializeDocumentBody(body);
    expect(out).toContain('<w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>RM</w:t></w:r>');
  });
});

describe('glyph checkbox detection', () => {
  test('finds the hand-typed ☐ cells and leaves the real controls alone', async () => {
    const body = parseDocumentBody(await fixtureXml());
    const glyphs = findGlyphCheckboxes(body);
    expect(glyphs.length).toBeGreaterThan(0);
    for (const g of glyphs) {
      expect(g.kind).toBe('glyph');
      expect(g.encoding).toBe('text');
      expect(g.checked).toBe(false);
      expect(g.char).toBe('☐');
    }
    // The synthesized legacy-checkbox glyph lives inside an inlineSdt run, not
    // a bare paragraph run, so it is not reported as a loose glyph.
    expect(glyphs.length).toBe(9);
  });

  test('detects a Wingdings w:sym ballot box', () => {
    const xml =
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' +
      '<w:p><w:r><w:sym w:font="Wingdings" w:char="F0FE"/></w:r></w:p></w:body></w:document>';
    const [glyph] = findGlyphCheckboxes(parseDocumentBody(xml));
    expect(glyph).toMatchObject({
      kind: 'glyph',
      encoding: 'sym',
      checked: true,
      font: 'Wingdings',
    });
  });
});

describe("legacy form fields — ffData patching follows the file's own prefixes", () => {
  /**
   * The parser accepts ffData children by local name, so a file that binds the
   * WordprocessingML namespace to another prefix inside the field projects
   * exactly like a `w:` one. Patches must then speak that prefix too, or the
   * modeled state and the replayed bytes disagree after save.
   */
  const XNS = `${NS} xmlns:x="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`;
  const xDropdownXml =
    '<w:r><w:fldChar w:fldCharType="begin"><x:ffData><x:name w:val="XD"/><x:enabled/>' +
    '<x:ddList><x:listEntry w:val="A"/><x:listEntry w:val="B"/><x:listEntry w:val="C"/>' +
    '</x:ddList></x:ffData></w:fldChar></w:r>' +
    '<w:r><w:instrText xml:space="preserve"> FORMDROPDOWN </w:instrText></w:r>' +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>';

  test('a dropdown whose ffData uses another prefix keeps its edit through save and reparse', () => {
    const body = parseDocumentBody(
      `<w:document ${XNS}><w:body><w:p>${xDropdownXml}</w:p></w:body></w:document>`
    );
    expect(findContentControl(body, { tag: 'XD' })!.text).toBe('A');

    const next = setContentControlValue(
      asDocument(body),
      { tag: 'XD' },
      { kind: 'dropdown', value: 'B' }
    );
    const out = serializeDocumentBody(next.package.document);
    // `w:result` is inserted in the file's dialect, first in `x:ddList`.
    expect(out).toContain('<x:ddList><x:result w:val="1"/><x:listEntry w:val="A"/>');
    expect(out).toContain('<w:t>B</w:t>');

    const reparsed = parseDocumentBody(`<w:document ${XNS}><w:body>${out}</w:body></w:document>`);
    expect(findContentControl(reparsed, { tag: 'XD' })!.legacyFormField).toMatchObject({
      selectedIndex: 1,
      value: 'B',
    });
  });

  /** A descriptor as the parser would build it, for the setters alone. */
  const fieldWith = (ffDataXml: string, extra: Partial<LegacyFormField>): LegacyFormField => ({
    kind: 'legacy',
    fieldType: 'checkbox',
    instruction: 'FORMCHECKBOX',
    ffDataXml,
    rawPrefixXml: `<w:r><w:fldChar w:fldCharType="begin">${ffDataXml}</w:fldChar></w:r>`,
    rawSuffixXml: '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
    hasSeparate: true,
    hasResult: false,
    ...extra,
  });

  test('checkbox patches use the prefix of both the elements and the val attribute', () => {
    const ffData =
      '<x:ffData><x:name x:val="XC"/><x:checkBox><x:sizeAuto/><x:default x:val="0"/></x:checkBox></x:ffData>';
    const next = setLegacyCheckbox(fieldWith(ffData, { checked: false }), true);
    const patched =
      '<x:ffData><x:name x:val="XC"/><x:checkBox><x:sizeAuto/><x:default x:val="1"/>' +
      '<x:checked x:val="1"/></x:checkBox></x:ffData>';
    expect(next.ffDataXml).toBe(patched);
    expect(next.rawPrefixXml).toContain(patched);
    expect(next.checked).toBe(true);
  });

  test('a default-namespace (unprefixed) ffData is patched without inventing a prefix', () => {
    const ffData =
      '<ffData><name val="D"/><ddList><listEntry val="A"/><listEntry val="B"/></ddList></ffData>';
    const field = fieldWith(ffData, {
      fieldType: 'dropdown',
      instruction: 'FORMDROPDOWN',
      options: ['A', 'B'],
      selectedIndex: 0,
    });
    const next = setLegacyDropdownIndex(field, 1);
    expect(next.ffDataXml).toBe(
      '<ffData><name val="D"/><ddList><result val="1"/><listEntry val="A"/><listEntry val="B"/></ddList></ffData>'
    );
    expect(setLegacyDropdownIndex(next, 0).ffDataXml).toContain('<ddList><result val="0"/>');
    expect(next).toMatchObject({ selectedIndex: 1, value: 'B', hasResult: true });
  });

  test('the w: prefix still patches exactly as before', () => {
    const ffData = '<w:ffData><w:name w:val="C"/><w:checkBox><w:sizeAuto/></w:checkBox></w:ffData>';
    expect(setLegacyCheckbox(fieldWith(ffData, {}), true).ffDataXml).toBe(
      '<w:ffData><w:name w:val="C"/><w:checkBox><w:sizeAuto/><w:default w:val="1"/><w:checked w:val="1"/></w:checkBox></w:ffData>'
    );
  });
});

describe('legacy form fields — FORMTEXT descriptor follows generic content writes', () => {
  const RPR = '<w:rPr><w:sz w:val="13"/></w:rPr>';
  const textXml = (textInput: string, name: string) =>
    `<w:r>${RPR}<w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="${name}"/>` +
    (textInput ? `<w:textInput>${textInput}</w:textInput>` : '<w:textInput/>') +
    `</w:ffData></w:fldChar></w:r>` +
    `<w:r>${RPR}<w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r>` +
    `<w:r>${RPR}<w:fldChar w:fldCharType="separate"/></w:r>` +
    `<w:r>${RPR}<w:fldChar w:fldCharType="end"/></w:r>`;
  const docXml = (...paragraphs: string[]) =>
    `<w:document ${NS}><w:body>${paragraphs.map((p) => `<w:p>${p}</w:p>`).join('')}</w:body></w:document>`;

  test('setContentControlContent updates value/hasResult alongside the runs', () => {
    const doc = asDocument(parseDocumentBody(docXml(textXml('', 'T'))));
    expect(findContentControl(doc, { tag: 'T' })!.legacyFormField).toMatchObject({
      value: '',
      hasResult: false,
    });

    const next = setContentControlContent(doc, { tag: 'T' }, 'Yes');
    const info = findContentControl(next, { tag: 'T' })!;
    expect(info.text).toBe('Yes');
    expect(info.legacyFormField).toMatchObject({
      fieldType: 'text',
      value: 'Yes',
      hasResult: true,
    });

    const out = serializeDocumentBody(next.package.document);
    expect(out).toContain('<w:fldChar w:fldCharType="separate"/></w:r><w:r>');
    expect(out).toContain('<w:t>Yes</w:t>');
    expect(out).not.toContain(LEGACY_TEXT_PLACEHOLDER);
    const reparsed = findContentControl(reparse(out), { tag: 'T' })!;
    expect(reparsed.text).toBe('Yes');
    expect(reparsed.legacyFormField).toMatchObject({ value: 'Yes', hasResult: true });
  });

  test("writing the field's own default keeps it result-less and byte-identical", () => {
    const xml = docXml(textXml('<w:default w:val="TBD"/>', 'T'));
    const doc = asDocument(parseDocumentBody(xml));
    const next = setContentControlContent(doc, { tag: 'T' }, 'TBD');
    expect(findContentControl(next, { tag: 'T' })!.legacyFormField).toMatchObject({
      value: 'TBD',
      hasResult: false,
    });
    expect(serializeDocumentBody(next.package.document)).toBe(
      serializeDocumentBody(doc.package.document)
    );
  });

  test('typing over the field in the editor reaches the descriptor on the way back', () => {
    const body = parseDocumentBody(docXml(textXml('', 'T')));
    const pmDoc = toProseDoc(asDocument(body));
    let sdtPos = -1;
    let sdtSize = 0;
    pmDoc.descendants((node, pos) => {
      if (node.type.name === 'sdt' && sdtPos < 0) {
        sdtPos = pos;
        sdtSize = node.content.size;
      }
      return sdtPos < 0;
    });
    expect(sdtPos).toBeGreaterThanOrEqual(0);

    const state = EditorState.create({ doc: pmDoc });
    const typed = state.apply(
      state.tr.replaceWith(sdtPos + 1, sdtPos + 1 + sdtSize, state.schema.text('Typed answer'))
    );
    const edited = fromProseDoc(typed.doc).package.document;
    const info = findContentControl(edited, { tag: 'T' })!;
    expect(info.text).toBe('Typed answer');
    expect(info.legacyFormField).toMatchObject({ value: 'Typed answer', hasResult: true });
    expect(serializeDocumentBody(edited)).toContain('<w:t>Typed answer</w:t>');

    // An untouched trip through the editor still leaves the descriptor alone.
    const untouched = findContentControl(fromProseDoc(pmDoc).package.document, { tag: 'T' })!;
    expect(untouched.legacyFormField).toMatchObject({ value: '', hasResult: false });
  });
});

describe('glyph checkbox detection — several boxes in one run', () => {
  const WNS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  const para = (...runs: string[]) =>
    `<w:document ${WNS}><w:body><w:p>${runs.join('')}</w:p></w:body></w:document>`;

  test('an answer line typed as one run yields one candidate per box, with offsets', () => {
    const body = parseDocumentBody(
      para('<w:r><w:t xml:space="preserve">☐ Yes ☐ No ☒ N/A</w:t></w:r>')
    );
    const glyphs = findGlyphCheckboxes(body);
    expect(glyphs.map((g) => [g.char, g.checked, g.runIndex, g.offset])).toEqual([
      ['☐', false, 0, 0],
      ['☐', false, 0, 6],
      ['☒', true, 0, 11],
    ]);
    expect(glyphs[0].paragraphText).toBe('☐ Yes ☐ No ☒ N/A');
  });

  test('a box glued to its label, and boxes split across runs, are all found', () => {
    const body = parseDocumentBody(
      para(
        '<w:r><w:t>☑Compliant</w:t></w:r>',
        '<w:r><w:t xml:space="preserve"> or </w:t></w:r>',
        '<w:r><w:t>☐</w:t></w:r>'
      )
    );
    expect(findGlyphCheckboxes(body).map((g) => [g.runIndex, g.offset, g.checked])).toEqual([
      [0, 0, true],
      [2, 0, false],
    ]);
  });

  test('several w:sym boxes in one run are reported individually', () => {
    const body = parseDocumentBody(
      para(
        '<w:r><w:sym w:font="Wingdings" w:char="F06F"/><w:t xml:space="preserve"> Yes </w:t>' +
          '<w:sym w:font="Wingdings" w:char="F0FE"/><w:t xml:space="preserve"> No</w:t></w:r>'
      )
    );
    expect(findGlyphCheckboxes(body).map((g) => [g.encoding, g.offset, g.checked])).toEqual([
      ['sym', 0, false],
      ['sym', 6, true],
    ]);
  });

  test('a paragraph with no box yields nothing', () => {
    expect(findGlyphCheckboxes(parseDocumentBody(para('<w:r><w:t>Yes / No</w:t></w:r>')))).toEqual(
      []
    );
  });
});

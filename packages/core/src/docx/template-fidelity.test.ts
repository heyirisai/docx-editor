/**
 * Fidelity of branded Word templates through a save: untouched header/footer
 * parts, preserved `mc:AlternateContent`, and a field sharing a `w:r` with text.
 */

import { describe, expect, test } from 'bun:test';
import type { XmlElement } from './xmlParser';
import { parseXmlDocument } from './xmlParser';
import { parseParagraph } from './paragraphParser';
import { parseFooter, headerFooterSnapshot } from './headerFooterParser';
import { parseXmlDocument as parseDoc, elementToSelfContainedXml } from './xmlParser';
import {
  collectHeaderFooterUpdates,
  commitHeaderFooterSnapshots,
  adoptSavedBuffer,
} from './rezip/packaging';
import type { Document, HeaderFooter } from '../types/document';
import { serializeParagraph } from './serializer/paragraphSerializer';
import { serializeHeaderFooter } from './serializer/headerFooterSerializer';
import type { ComplexField, Run } from '../types/document';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const MC = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"';
const WPG = 'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"';
const V = 'xmlns:v="urn:schemas-microsoft-com:vml"';
const WP = 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"';

function para(inner: string) {
  const root = parseXmlDocument(`<w:p ${W} ${MC} ${WPG} ${V} ${WP}>${inner}</w:p>`);
  return parseParagraph(root as XmlElement, null, null, null, null, null);
}

describe('mc:AlternateContent without image data', () => {
  const GROUP = `
    <w:r><mc:AlternateContent>
      <mc:Choice Requires="wpg"><w:drawing><wp:anchor>
        <wp:docPr id="1" name="Group 1"/><wpg:wgp/>
      </wp:anchor></w:drawing></mc:Choice>
      <mc:Fallback><w:pict><v:group id="_x0000_s1027"/></w:pict></mc:Fallback>
    </mc:AlternateContent></w:r>`;

  test('a grouped drawing is carried through as source, not dropped', () => {
    const raw = para(GROUP)
      .content.filter((c) => c.type === 'run')
      .flatMap((r) => (r as Run).content)
      .filter((c) => c.type === 'rawXml');
    expect(raw).toHaveLength(1);
    expect((raw[0] as { xml: string }).xml).toContain('wpg:wgp');
  });

  test('it round-trips back into the serialized paragraph', () => {
    const xml = serializeParagraph(para(GROUP));
    expect(xml).toContain('mc:AlternateContent');
    expect(xml).toContain('wpg:wgp');
    expect(xml).toContain('name="Group 1"');
    // The VML fallback rides along inside the preserved source.
    expect(xml).toContain('w:pict');
  });
});

describe('a run holding both text and a complete field', () => {
  // How Word writes a TOC entry: the title, a tab, and the whole PAGEREF
  // field all live in one w:r.
  const ENTRY = `
    <w:r><w:t>2.</w:t></w:r>
    <w:r>
      <w:t>Tender Response</w:t><w:tab/>
      <w:fldChar w:fldCharType="begin"/>
      <w:instrText xml:space="preserve"> PAGEREF _Toc1 \\h </w:instrText>
      <w:fldChar w:fldCharType="separate"/>
      <w:t>3</w:t>
      <w:fldChar w:fldCharType="end"/>
    </w:r>`;

  test('the entry text survives alongside the field', () => {
    const text = para(ENTRY)
      .content.filter((c) => c.type === 'run')
      .flatMap((r) => (r as Run).content)
      .map((c) => (c.type === 'text' ? c.text : ''))
      .join('');
    expect(text).toContain('2.');
    expect(text).toContain('Tender Response');
  });

  test('the field keeps its instruction and its result', () => {
    const field = para(ENTRY).content.find((c) => c.type === 'complexField') as ComplexField;
    expect(field).toBeDefined();
    expect(field.instruction).toBe('PAGEREF _Toc1 \\h');
    const result = field.fieldResult
      .flatMap((r) => r.content)
      .map((c) => (c.type === 'text' ? c.text : ''))
      .join('');
    expect(result).toBe('3');
  });

  test('serializing it back keeps text, instruction and page number', () => {
    const xml = serializeParagraph(para(ENTRY));
    expect(xml).toContain('Tender Response');
    expect(xml).toContain('PAGEREF _Toc1');
    expect(xml).toContain('<w:t>3</w:t>');
  });
});

describe('a paragraph that opens a multi-paragraph field', () => {
  // Word's first TOC entry paragraph: the TOC field's own begin/instruction/
  // separate have no matching end here, but the entry's PAGEREF field does.
  const FIRST_ENTRY = `
    <w:r>
      <w:fldChar w:fldCharType="begin" w:dirty="true"/>
      <w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText>
      <w:fldChar w:fldCharType="separate"/>
    </w:r>
    <w:r>
      <w:t>Introduction</w:t><w:tab/>
      <w:fldChar w:fldCharType="begin"/>
      <w:instrText xml:space="preserve"> PAGEREF _Toc1 \\h </w:instrText>
      <w:fldChar w:fldCharType="separate"/>
      <w:t>1</w:t>
      <w:fldChar w:fldCharType="end"/>
    </w:r>`;

  test('the enclosed field is still modelled', () => {
    const field = para(FIRST_ENTRY).content.find((c) => c.type === 'complexField') as ComplexField;
    expect(field).toBeDefined();
    expect(field.instruction).toBe('PAGEREF _Toc1 \\h');
  });

  test('the unpaired wrapper markers ride through as run content', () => {
    const items = para(FIRST_ENTRY)
      .content.filter((c) => c.type === 'run')
      .flatMap((r) => (r as Run).content);
    expect(items.filter((c) => c.type === 'fieldChar')).toHaveLength(2);
    expect(
      items.filter((c) => c.type === 'instrText').map((c) => (c as { text: string }).text)
    ).toEqual([' TOC \\o "1-3" \\h ']);
  });

  test('the entry text is not swallowed', () => {
    const text = para(FIRST_ENTRY)
      .content.filter((c) => c.type === 'run')
      .flatMap((r) => (r as Run).content)
      .map((c) => (c.type === 'text' ? c.text : ''))
      .join('');
    expect(text).toContain('Introduction');
  });
});

describe('field code regions with more than an instruction', () => {
  // Word emits an empty formatting-only run between `begin` and the
  // instruction. The instruction must survive regardless.
  const PAGE_FIELD = `
    <w:r><w:rPr><w:b/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:rPr><w:b/></w:rPr></w:r>
    <w:r><w:rPr><w:b/></w:rPr><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
    <w:r><w:rPr><w:b/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:rPr><w:b/></w:rPr><w:t>7</w:t></w:r>
    <w:r><w:rPr><w:b/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>`;

  test('the instruction is still written out', () => {
    expect(serializeParagraph(para(PAGE_FIELD))).toContain('PAGE');
  });

  test('formatting carried only by the instruction run is kept', () => {
    // `begin` here has no w:rPr, so the fallback must come from a later run.
    const noFormatOnBegin = `
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:rPr><w:b/></w:rPr><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>`;
    const field = para(noFormatOnBegin).content.find(
      (c) => c.type === 'complexField'
    ) as ComplexField;
    expect(field.formatting?.bold).toBe(true);
  });
});

describe('a deleted run holding both text and preserved markup', () => {
  // Word packs a picture (or an mc:AlternateContent shape) into the same w:r as
  // the text around it. Under w:del only the run's OWN text becomes delText —
  // the <w:t> nested inside a textbox belongs to that inner document.
  const RAW =
    '<mc:AlternateContent><wps:txbx><w:txbxContent>' +
    '<w:p><w:r><w:t>inside the box</w:t></w:r></w:p>' +
    '</w:txbxContent></wps:txbx></mc:AlternateContent>';

  function deletedRunXml(content: Run['content']): string {
    return serializeParagraph({
      type: 'paragraph',
      content: [
        {
          type: 'deletion',
          info: { id: 1, author: 'A', date: '2026-09-17T00:00:00Z' },
          content: [{ type: 'run', content }],
        },
      ],
    } as never);
  }

  test('the run text is rewritten but the nested textbox text is not', () => {
    const xml = deletedRunXml([
      { type: 'text', text: 'removed' },
      { type: 'rawXml', xml: RAW },
    ]);
    expect(xml).toContain('<w:delText>removed</w:delText>');
    expect(xml).toContain('<w:t>inside the box</w:t>');
    expect(xml).not.toContain('<w:delText>inside the box</w:delText>');
  });

  test('a text-only deleted run is unaffected', () => {
    expect(deletedRunXml([{ type: 'text', text: 'removed' }])).toContain(
      '<w:delText>removed</w:delText>'
    );
  });

  test('content order is preserved across the split', () => {
    const xml = deletedRunXml([
      { type: 'text', text: 'before' },
      { type: 'rawXml', xml: '<w:drawing/>' },
      { type: 'text', text: 'after' },
    ]);
    expect(xml.indexOf('before')).toBeLessThan(xml.indexOf('<w:drawing/>'));
    expect(xml.indexOf('<w:drawing/>')).toBeLessThan(xml.indexOf('after'));
  });
});

describe('untouched header/footer parts', () => {
  const FOOTER = `<w:ftr ${W} ${MC} ${WPG} ${V} ${WP}>
    <w:p><w:r><w:t>SQE</w:t></w:r></w:p>
  </w:ftr>`;

  test('parsing records a snapshot used to detect an edit', () => {
    const footer = parseFooter(FOOTER);
    expect(footer.originalSnapshot).toBeDefined();
    expect(headerFooterSnapshot(footer)).toBe(footer.originalSnapshot!);
  });

  test('editing the content changes the snapshot', () => {
    const footer = parseFooter(FOOTER);
    (footer.content[0] as { content: unknown[] }).content = [];
    expect(headerFooterSnapshot(footer)).not.toBe(footer.originalSnapshot!);
  });
});

describe('preserved markup declares the namespaces it uses', () => {
  test('a sibling inheriting a prefix from the old root is still bound', () => {
    // The first a:graphic declares `a` on ITSELF; the second inherited it from
    // the document root that is about to be left behind.
    const frag = parseDoc(`<mc:AlternateContent
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <mc:Choice Requires="wpg"><w:drawing>
        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>
        <a:graphic/>
      </w:drawing></mc:Choice>
    </mc:AlternateContent>`);
    const xml = elementToSelfContainedXml(frag!);
    // The fragment root must now bind `a` for the inheriting sibling.
    expect(/^<mc:AlternateContent[^>]*xmlns:a=/.test(xml)).toBe(true);
  });
});

describe('header/footer baseline after a save', () => {
  function docWith(footer: HeaderFooter): Document {
    return {
      package: {
        document: { content: [] },
        headers: new Map(),
        footers: new Map([['rId1', footer]]),
        relationships: new Map([
          [
            'rId1',
            {
              id: 'rId1',
              type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer',
              target: 'footer1.xml',
            },
          ],
        ]),
      },
    } as unknown as Document;
  }

  test('an edit reverted after a save is written out, not skipped', () => {
    const footer = parseFooter(`<w:ftr ${W}><w:p><w:r><w:t>SQE</w:t></w:r></w:p></w:ftr>`);
    const doc = docWith(footer);
    const original = JSON.stringify(footer.content);

    // Edit + save: the part is written, and that save becomes the baseline.
    (footer.content[0] as { content: unknown[] }).content = [];
    expect([...collectHeaderFooterUpdates(doc, () => true).keys()]).toHaveLength(1);
    commitHeaderFooterSnapshots(doc);

    // Revert to exactly what was first parsed, then save again. The ZIP now
    // holds the edit, so this restoration MUST be written.
    footer.content = JSON.parse(original);
    expect([...collectHeaderFooterUpdates(doc, () => true).keys()]).toHaveLength(1);
  });

  test('exporting twice without adopting the buffer writes the edit each time', () => {
    // Callers that never replace `originalBuffer` keep saving against the
    // ORIGINAL ZIP, so the part must be written on every export.
    const footer = parseFooter(`<w:ftr ${W}><w:p><w:r><w:t>SQE</w:t></w:r></w:p></w:ftr>`);
    const doc = docWith(footer);
    (footer.content[0] as { content: unknown[] }).content = [];

    expect([...collectHeaderFooterUpdates(doc, () => true).keys()]).toHaveLength(1);
    expect([...collectHeaderFooterUpdates(doc, () => true).keys()]).toHaveLength(1);

    // Only adopting the saved buffer moves the baseline.
    adoptSavedBuffer(doc, new ArrayBuffer(0));
    expect([...collectHeaderFooterUpdates(doc, () => true).keys()]).toHaveLength(0);
  });

  test('an untouched part is still skipped', () => {
    const footer = parseFooter(`<w:ftr ${W}><w:p><w:r><w:t>SQE</w:t></w:r></w:p></w:ftr>`);
    expect([...collectHeaderFooterUpdates(docWith(footer), () => true).keys()]).toHaveLength(0);
  });
});

describe('namespaces outside the known table', () => {
  test('a prefix the source root declared is re-declared on save', () => {
    const footer = parseFooter(
      `<w:ftr ${W} xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main">
         <w:p><w:r><w:t>x</w:t></w:r></w:p>
       </w:ftr>`
    );
    expect(footer.rootNamespaces?.a16).toBe(
      'http://schemas.microsoft.com/office/drawing/2014/main'
    );
    // Force a re-serialization and confirm the binding survives.
    (footer.content[0] as { content: unknown[] }).content = [];
    expect(serializeHeaderFooter(footer)).toContain('xmlns:a16=');
  });
});

describe('a destination that does not have the part yet', () => {
  test('an unchanged part is still written when the target lacks it', () => {
    // createDocx builds a fresh ZIP; skipping there would leave the
    // relationship pointing at a file that was never written.
    const footer = parseFooter(`<w:ftr ${W}><w:p><w:r><w:t>SQE</w:t></w:r></w:p></w:ftr>`);
    const doc = {
      package: {
        document: { content: [] },
        headers: new Map(),
        footers: new Map([['rId1', footer]]),
        relationships: new Map([
          [
            'rId1',
            {
              id: 'rId1',
              type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer',
              target: 'footer1.xml',
            },
          ],
        ]),
      },
    } as unknown as Document;
    expect([...collectHeaderFooterUpdates(doc, () => false).keys()]).toHaveLength(1);
  });
});

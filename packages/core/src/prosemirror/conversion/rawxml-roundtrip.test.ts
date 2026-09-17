/**
 * Preserved OOXML must survive the ProseMirror round-trip: the editor's save
 * rebuilds the body from PM, so anything PM drops is gone from the saved file.
 */

import { afterAll, beforeAll, describe, test, expect } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { DOMParser as PMDOMParser, DOMSerializer } from 'prosemirror-model';
import { schema } from '../schema';
import { toProseDoc } from './toProseDoc';
import { fromProseDoc } from './fromProseDoc';
import type { Document, Paragraph, Run } from '../../types/document';

beforeAll(() => GlobalRegistrator.register());
afterAll(() => GlobalRegistrator.unregister());

const GROUP_XML =
  '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
  '<mc:Choice Requires="wpg"><w:drawing/></mc:Choice></mc:AlternateContent>';

function docWith(runContent: Run['content']): Document {
  const paragraph: Paragraph = {
    type: 'paragraph',
    content: [{ type: 'run', content: runContent }],
  };
  return {
    package: { document: { content: [paragraph] } },
  } as unknown as Document;
}

function firstRunContent(doc: Document): Run['content'] {
  const block = doc.package.document.content[0] as Paragraph;
  const out: Run['content'] = [];
  for (const item of block.content) {
    if (item.type === 'run') out.push(...item.content);
  }
  return out;
}

describe('rawXml through ProseMirror', () => {
  test('preserved markup survives the round-trip byte for byte', () => {
    const before = docWith([{ type: 'rawXml', xml: GROUP_XML }]);
    const after = fromProseDoc(toProseDoc(before), before);
    const raw = firstRunContent(after).filter((c) => c.type === 'rawXml');
    expect(raw).toHaveLength(1);
    expect((raw[0] as { xml: string }).xml).toBe(GROUP_XML);
  });

  test('a canvas-only picture keeps its renderOnly flag', () => {
    const before = docWith([
      {
        type: 'drawing',
        image: {
          type: 'image',
          rId: 'rId1',
          src: 'data:image/png;base64,iVBORw0KGgo=',
          size: { width: 100, height: 100 },
          wrap: { type: 'inFront' },
          renderOnly: true,
        },
      },
    ]);
    const after = fromProseDoc(toProseDoc(before), before);
    const drawings = firstRunContent(after).filter((c) => c.type === 'drawing');
    expect(drawings).toHaveLength(1);
    expect((drawings[0] as { image: { renderOnly?: boolean } }).image.renderOnly).toBe(true);
  });

  test('a paragraph-spanning field keeps its markers through PM', () => {
    const before = docWith([
      { type: 'fieldChar', charType: 'begin', dirty: true },
      { type: 'instrText', text: ' TOC \\o "1-3" \\h ' },
      { type: 'fieldChar', charType: 'separate' },
    ]);
    const after = fromProseDoc(toProseDoc(before), before);
    const xml = firstRunContent(after)
      .filter((c) => c.type === 'rawXml')
      .map((c) => (c as { xml: string }).xml)
      .join('');
    expect(xml).toContain('w:fldCharType="begin"');
    expect(xml).toContain('w:dirty="true"');
    expect(xml).toContain('TOC \\o &quot;1-3&quot; \\h');
    expect(xml).toContain('w:fldCharType="separate"');
  });

  test('markup pasted from foreign HTML is rejected at the parseDOM boundary', () => {
    const parse = (rawXml: string) => {
      const dom = new window.DOMParser().parseFromString(
        `<p><span data-raw-xml="${rawXml}"></span>text</p>`,
        'text/html'
      );
      const parsed = PMDOMParser.fromSchema(schema).parse(dom.body);
      let count = 0;
      parsed.descendants((node) => {
        if (node.type.name === 'rawXml') count++;
        return true;
      });
      return count;
    };

    // Malformed — would make Word reject the whole part.
    expect(parse('&lt;w:t&gt;unclosed')).toBe(0);
    // Two elements — not a single self-contained fragment.
    expect(parse('&lt;w:t/&gt;&lt;w:t/&gt;')).toBe(0);
    // Well-formed source, e.g. copied between documents.
    expect(parse('&lt;w:fldChar w:fldCharType=&quot;begin&quot;/&gt;')).toBe(1);
  });

  test('the source root namespaces survive the save', () => {
    // A preserved fragment can use a prefix only the original root declared, so
    // losing this on the PM rebuild exports an unbound prefix.
    const before = docWith([{ type: 'rawXml', xml: GROUP_XML }]);
    before.package.document.rootNamespaces = { zz: 'urn:example:zz', a16: 'urn:example:a16' };
    const after = fromProseDoc(toProseDoc(before), before);
    expect(after.package.document.rootNamespaces).toEqual({
      zz: 'urn:example:zz',
      a16: 'urn:example:a16',
    });
  });

  test('renderOnly survives the clipboard DOM round-trip', () => {
    // Copy/paste goes through toDOM -> parseDOM. Losing the flag turns a group
    // preview into a real picture, so the save emits it twice.
    const doc = toProseDoc(
      docWith([
        {
          type: 'drawing',
          image: {
            type: 'image',
            rId: 'rId1',
            src: 'data:image/png;base64,iVBORw0KGgo=',
            size: { width: 100, height: 100 },
            wrap: { type: 'inFront' },
            renderOnly: true,
          },
        },
      ])
    );
    const html = document.createElement('div');
    html.appendChild(DOMSerializer.fromSchema(schema).serializeFragment(doc.content));
    expect(html.innerHTML).toContain('data-render-only');

    const reparsed = PMDOMParser.fromSchema(schema).parse(html);
    let images = 0;
    reparsed.descendants((node) => {
      if (node.type.name === 'image') {
        images++;
        expect(node.attrs.renderOnly).toBe(true);
      }
      return true;
    });
    expect(images).toBe(1);
  });
});

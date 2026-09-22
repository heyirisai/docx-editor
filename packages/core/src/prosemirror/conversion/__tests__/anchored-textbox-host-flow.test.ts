/**
 * An anchored text box does not consume its host paragraph.
 *
 * `wp:anchor` shapes are out of flow, so in Word the `w:p` that holds one still
 * occupies a line — that stray empty paragraph under a floating object is why
 * you cannot delete one without deleting the other. Dropping the host pulled
 * every later block up by a line plus the host's spacing, and on a cover page
 * built from empty spacer paragraphs (`Heading2` with style spacing and nothing
 * in it) that walked the artwork up off the bottom of the page.
 *
 * An IN-FLOW box is different: it is the paragraph's content, so it replaces an
 * emptied host and carries its id.
 */
import { describe, expect, test } from 'bun:test';
import { toProseDoc } from '../toProseDoc';
import { fromProseDoc } from '../fromProseDoc';
import type { Document, Paragraph } from '../../../types/document';
import type { WrapType } from '../../../docx/wrapTypes';

const HOST_ID = 'AAAA1111';

function hostParagraph(wrapType: WrapType, styleId?: string): Paragraph {
  return {
    type: 'paragraph',
    paraId: HOST_ID,
    formatting: styleId ? { styleId, spaceBefore: 100, spaceAfter: 220 } : undefined,
    content: [
      {
        type: 'run',
        content: [
          {
            type: 'shape',
            shape: {
              type: 'shape',
              shapeType: 'textBox',
              size: { width: 2743200, height: 914400 },
              wrap: { type: wrapType },
              textBody: {
                content: [
                  {
                    type: 'paragraph',
                    content: [{ type: 'run', content: [{ type: 'text', text: 'COVER TITLE' }] }],
                  },
                ],
              },
            },
          },
        ],
      },
    ],
  } as Paragraph;
}

const docWith = (...blocks: Paragraph[]): Document =>
  ({ package: { document: { content: blocks } } }) as Document;

const topLevel = (doc: Document): string[] => {
  const out: string[] = [];
  toProseDoc(doc).forEach((n) => out.push(n.type.name));
  return out;
};

describe('an anchored text box keeps its host paragraph in the flow', () => {
  test.each<WrapType>(['square', 'tight', 'through', 'topAndBottom', 'behind', 'inFront'])(
    '%s leaves the host paragraph in place',
    (wrapType) => {
      expect(topLevel(docWith(hostParagraph(wrapType)))).toEqual(['textBox', 'paragraph']);
    }
  );

  test('an in-flow box still replaces the host it emptied', () => {
    // Nothing is floating, so there is no line for the host to occupy.
    expect(topLevel(docWith(hostParagraph('inline')))).toEqual(['textBox', 'paragraph']);
    // ...and the trailing paragraph there is the caret paragraph #861 appends,
    // not the host — the host is gone, so the box owns its id.
    const pm = toProseDoc(docWith(hostParagraph('inline')));
    expect(pm.child(0).attrs.hostParaId).toBe(HOST_ID);
  });

  test('an anchored box does not adopt the host id — the host keeps it', () => {
    const pm = toProseDoc(docWith(hostParagraph('tight')));
    expect(pm.child(0).type.name).toBe('textBox');
    expect(pm.child(0).attrs.hostParaId).toBeNull();
  });

  test('a cover spacer stack keeps every paragraph that carries no box', () => {
    const spacer = (): Paragraph =>
      ({ type: 'paragraph', formatting: { styleId: 'Heading2' }, content: [] }) as Paragraph;
    const kinds = topLevel(
      docWith(spacer(), spacer(), hostParagraph('tight', 'Heading2'), spacer(), spacer())
    );
    // 4 spacers + the host + the extracted box.
    expect(kinds).toEqual([
      'paragraph',
      'paragraph',
      'textBox',
      'paragraph',
      'paragraph',
      'paragraph',
    ]);
  });
});

describe('the host paragraph survives the export', () => {
  test('its style and spacing come back, not a bare rebuilt paragraph', () => {
    const doc = docWith(hostParagraph('tight', 'Heading2'));
    const body = fromProseDoc(toProseDoc(doc, { styles: undefined }), doc).package.document.content;

    expect(body.length).toBe(1);
    const host = body[0] as Paragraph;
    expect(JSON.stringify(host)).toContain('"shape"');
    expect(host.paraId).toBe(HOST_ID);
    // The old path rebuilt `{type:'paragraph', content:[run]}` and lost these.
    expect(host.formatting?.styleId).toBe('Heading2');
    expect(host.formatting?.spaceBefore).toBe(100);
    expect(host.formatting?.spaceAfter).toBe(220);
  });

  test('two boxes sharing one host land back on that host', () => {
    const doc = docWith(hostParagraph('tight', 'Heading2'));
    (doc.package.document.content[0] as Paragraph).content.push({
      type: 'run',
      content: [
        {
          type: 'shape',
          shape: {
            type: 'shape',
            shapeType: 'textBox',
            size: { width: 100, height: 100 },
            wrap: { type: 'tight' },
            textBody: {
              content: [
                {
                  type: 'paragraph',
                  content: [{ type: 'run', content: [{ type: 'text', text: 'second' }] }],
                },
              ],
            },
          },
        },
      ],
    } as never);

    const body = fromProseDoc(toProseDoc(doc), doc).package.document.content;
    expect(body.length).toBe(1);
    expect(JSON.stringify(body[0]).match(/"shapeType"/g)?.length).toBe(2);
  });
});

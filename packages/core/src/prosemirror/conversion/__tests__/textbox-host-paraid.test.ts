import { describe, expect, test } from 'bun:test';
import type { Document, Paragraph } from '../../../types/document';
import { toProseDoc } from '../toProseDoc';
import { fromProseDoc } from '../fromProseDoc';

const HOST_PARA_ID = 'AAAA1111';
const NEXT_PARA_ID = 'BBBB2222';

/**
 * A paragraph whose only run is an anchored text box — the shape Word uses for
 * a cover title. Extracting the box leaves the host empty.
 */
function hostParagraphWithTextBox(): Paragraph {
  return {
    type: 'paragraph',
    paraId: HOST_PARA_ID,
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
              wrap: { type: 'topAndBottom' },
              textBody: {
                content: [
                  {
                    type: 'paragraph',
                    content: [
                      { type: 'run', content: [{ type: 'text', text: 'REQUEST FOR PROPOSAL' }] },
                    ],
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

function makeDocument(): Document {
  return {
    package: {
      document: {
        content: [
          hostParagraphWithTextBox(),
          { type: 'paragraph', paraId: NEXT_PARA_ID, content: [] } as Paragraph,
        ],
      },
    },
  } as Document;
}

function holdsShape(block: unknown): boolean {
  return JSON.stringify(block).includes('"shape"');
}

describe('text box host paragraph identity', () => {
  test('round-trip keeps the text box in its own host paragraph', () => {
    const doc = makeDocument();
    const roundTripped = fromProseDoc(toProseDoc(doc), doc);
    const body = roundTripped.package.document.content;

    expect(body.length).toBe(2);
    expect(holdsShape(body[0])).toBe(true);
    expect(holdsShape(body[1])).toBe(false);
  });

  test('round-trip preserves the host paragraph paraId', () => {
    const doc = makeDocument();
    const roundTripped = fromProseDoc(toProseDoc(doc), doc);
    const body = roundTripped.package.document.content;

    const shapeBlock = body.find((b) => holdsShape(b)) as Paragraph | undefined;
    expect(shapeBlock?.paraId).toBe(HOST_PARA_ID);
  });

  test('round-trip does not reassign the following paragraph paraId', () => {
    const doc = makeDocument();
    const roundTripped = fromProseDoc(toProseDoc(doc), doc);
    const ids = roundTripped.package.document.content.map((b) => (b as Paragraph).paraId);

    expect(ids).toEqual([HOST_PARA_ID, NEXT_PARA_ID]);
  });
});

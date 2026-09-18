/**
 * Pasted HTML is the other way into a text box's preserved markup.
 *
 * `data-body-pr-xml` and `data-sp-pr-extra-xml` are written back into
 * `wps:spPr` verbatim on the next save, so what a page puts there would land
 * inside the document part. Well-formedness is not the test — the element has
 * to be one that belongs where it is going.
 */
import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { DOMParser as PMDOMParser } from 'prosemirror-model';
import { schema } from '../../schema';

beforeAll(() => GlobalRegistrator.register());
afterAll(() => GlobalRegistrator.unregister());

const A = 'xmlns:a=&quot;http://schemas.openxmlformats.org/drawingml/2006/main&quot;';
const WPS =
  'xmlns:wps=&quot;http://schemas.microsoft.com/office/word/2010/wordprocessingShape&quot;';

/** Parse a pasted text box and hand back its attrs. */
function pastedTextBox(dataAttrs: string): Record<string, unknown> {
  const container = document.createElement('div');
  container.innerHTML =
    `<div class="docx-textbox" data-width="200" data-height="100" ${dataAttrs}>` +
    '<p>pasted</p></div>';

  const doc = PMDOMParser.fromSchema(schema).parse(container);
  let attrs: Record<string, unknown> | null = null;
  doc.descendants((node) => {
    if (node.type.name === 'textBox') attrs = node.attrs as Record<string, unknown>;
    return attrs === null;
  });
  if (!attrs) throw new Error('no textBox parsed from the pasted markup');
  return attrs;
}

describe('preserved shape markup from pasted HTML', () => {
  test('our own markup survives the round-trip through the DOM', () => {
    const attrs = pastedTextBox(
      `data-body-pr-xml="&lt;wps:bodyPr ${WPS} wrap=&quot;square&quot;/&gt;" ` +
        `data-sp-pr-extra-xml="&lt;a:ln ${A}/&gt;"`
    );
    expect(attrs.bodyPrXml).toContain('wrap="square"');
    expect(attrs.spPrExtraXml).toContain('<a:ln');
  });

  test('a paragraph dressed up as body properties does not get stored', () => {
    const attrs = pastedTextBox(
      'data-body-pr-xml="&lt;w:p xmlns:w=&quot;http://schemas.openxmlformats.org/wordprocessingml/2006/main&quot;/&gt;"'
    );
    expect(attrs.bodyPrXml).toBeNull();
  });

  test('an element smuggled in beside a legitimate one takes it down with it', () => {
    const attrs = pastedTextBox(
      `data-sp-pr-extra-xml="&lt;a:ln ${A}/&gt;&lt;w:drawing xmlns:w=&quot;http://schemas.openxmlformats.org/wordprocessingml/2006/main&quot;/&gt;"`
    );
    expect(attrs.spPrExtraXml).toBeNull();
  });

  test('a doctype in the fragment is refused outright', () => {
    const attrs = pastedTextBox(`data-body-pr-xml="&lt;!DOCTYPE p&gt;&lt;wps:bodyPr ${WPS}/&gt;"`);
    expect(attrs.bodyPrXml).toBeNull();
  });

  test('a box with no preserved markup pastes as one with none', () => {
    const attrs = pastedTextBox('');
    expect(attrs.bodyPrXml).toBeNull();
    expect(attrs.spPrExtraXml).toBeNull();
  });
});

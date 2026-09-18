/**
 * A text box carries more than the model has fields for.
 *
 * `wps:bodyPr` has a dozen attributes plus an autofit child; the model holds
 * five of them. `wps:spPr` can declare `<a:ln><a:noFill/></a:ln>` — an explicit
 * "no outline" — which parses to no `outline` at all, so rebuilding spPr from
 * the model swapped it for the DEFAULT outline. And an inline content control
 * inside the box lost its `w:sdtEndPr` because only the block path passed it on.
 *
 * All three are kept verbatim now and replayed on save, so an edit anywhere in
 * the document stops rewriting the shape's geometry and chrome.
 */
import { describe, expect, test } from 'bun:test';
import { parseXml, type XmlElement } from '../xmlParser';
import { parseTextBox } from '../textBoxParser';
import { serializeRun } from '../serializer/runSerializer';
import type { Run, Shape } from '../../types/document';

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"';

const BODY_PR =
  '<wps:bodyPr rot="0" spcFirstLastPara="0" vertOverflow="overflow" horzOverflow="overflow"' +
  ' vert="horz" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" numCol="1" spcCol="0"' +
  ' rtlCol="0" fromWordArt="0" anchor="t" anchorCtr="0" forceAA="0" compatLnSpc="1">' +
  '<a:prstTxWarp prst="textNoShape"><a:avLst/></a:prstTxWarp><a:spAutoFit/></wps:bodyPr>';

function drawing(spPrExtras = '<a:ln w="6350"><a:noFill/></a:ln><a:effectLst/>'): XmlElement {
  const xml =
    `<w:drawing ${NS}><wp:anchor behindDoc="0"><wp:extent cx="4686300" cy="6720840"/>` +
    '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">' +
    '<wps:wsp><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4686300" cy="6720840"/></a:xfrm>' +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>${spPrExtras}</wps:spPr>` +
    '<wps:txbx><w:txbxContent><w:p><w:r><w:t>COVER</w:t></w:r></w:p></w:txbxContent></wps:txbx>' +
    `${BODY_PR}</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing>`;
  return parseXml(xml).elements![0];
}

/** Serialize a shape the way a save does, and return the emitted XML. */
function saved(shape: Shape): string {
  return serializeRun({ type: 'run', content: [{ type: 'shape', shape }] } as Run);
}

function shapeFrom(tb: ReturnType<typeof parseTextBox>): Shape {
  return {
    type: 'shape',
    shapeType: 'textBox',
    size: tb!.size,
    outline: tb!.outline,
    spPrExtraXml: tb!.spPrExtraXml,
    textBody: { content: [], bodyPrXml: tb!.bodyPrXml },
  };
}

describe('text box shape fidelity', () => {
  test('the source bodyPr is captured whole, not reduced to the modelled fields', () => {
    const tb = parseTextBox(drawing());
    expect(tb?.bodyPrXml).toContain('<a:spAutoFit/>');
    expect(tb?.bodyPrXml).toContain('vertOverflow="overflow"');
    expect(tb?.bodyPrXml).toContain('compatLnSpc="1"');
  });

  test('it is replayed on save, autofit and all', () => {
    const xml = saved(shapeFrom(parseTextBox(drawing())));
    expect(xml).toContain('<a:spAutoFit/>');
    for (const attr of ['wrap="square"', 'anchor="t"', 'horzOverflow="overflow"']) {
      expect(xml).toContain(attr);
    }
  });

  test('an explicit "no outline" survives instead of becoming the default one', () => {
    const tb = parseTextBox(drawing());
    expect(tb?.outline).toBeUndefined(); // nothing to model — that IS the bug
    expect(tb?.spPrExtraXml).toContain('<a:noFill/>');
    const xml = saved(shapeFrom(tb));
    expect(xml).toContain('<a:ln');
    expect(xml).toContain('<a:effectLst');
  });

  test('an outline set on the model still wins over the preserved source', () => {
    const shape = shapeFrom(parseTextBox(drawing()));
    shape.outline = { width: 12700, color: { rgb: 'FF0000' }, style: 'solid' };
    const xml = saved(shape);
    expect(xml).toContain('FF0000');
    // The preserved "no outline" must not also be emitted — two `a:ln` in one
    // `spPr` is invalid.
    expect(xml.match(/<a:ln[\s>]/g)?.length).toBe(1);
  });

  test('a shape with no preserved source still serializes', () => {
    const xml = saved({
      type: 'shape',
      shapeType: 'textBox',
      size: { width: 100, height: 100 },
      textBody: { content: [] },
    });
    expect(xml).toContain('<wps:bodyPr');
    expect(xml).not.toContain('undefined');
  });

  test('malformed preserved markup is dropped rather than written', () => {
    const shape = shapeFrom(parseTextBox(drawing()));
    shape.textBody!.bodyPrXml = '<wps:bodyPr unclosed';
    shape.spPrExtraXml = '<a:ln><a:noFill/>';
    const xml = saved(shape);
    expect(xml).not.toContain('unclosed');
    // Falls back to the rebuilt bodyPr so the shape is still valid.
    expect(xml).toContain('<wps:bodyPr');
  });
});

describe('an inline content control keeps its end-mark properties', () => {
  const SDT_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  const END_PR =
    '<w:sdtEndPr><w:rPr><w:rStyle w:val="TitleChar"/><w:sz w:val="56"/></w:rPr></w:sdtEndPr>';

  test('w:sdtEndPr round-trips through the inline path, as it always did for blocks', async () => {
    const { parseParagraphContents } = await import('../paragraphParser/content');
    const el = parseXml(
      `<w:p ${SDT_NS}><w:sdt><w:sdtPr><w:alias w:val="Title"/></w:sdtPr>${END_PR}` +
        '<w:sdtContent><w:r><w:t>REQUEST FOR PROPOSAL</w:t></w:r></w:sdtContent></w:sdt></w:p>'
    ).elements![0];

    const contents = parseParagraphContents(el, null, null, null, null, null);
    const sdt = contents.find((c) => c.type === 'inlineSdt');
    expect(sdt).toBeDefined();
    const raw = JSON.stringify((sdt as { properties?: unknown }).properties);
    expect(raw).toContain('sdtEndPr');
    expect(raw).toContain('TitleChar');
  });
});

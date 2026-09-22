/**
 * The preserved shape markup, at both boundaries it crosses.
 *
 * Captured from a `.docx` it is whatever the source element was; captured from
 * pasted HTML it is whatever the page said. Both end up written verbatim into
 * `wps:spPr` on the next save, so the element and its children are checked
 * against the schema's own lists before either is stored.
 */
import { describe, expect, test } from 'bun:test';
import { parseXml, type XmlElement } from '../xmlParser';
import { parseShape } from '../shapeParser';
import {
  preservedBodyPrXml,
  preservedSpPrExtra,
  preservedSpPrExtraXml,
} from '../preservedShapeXml';

const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const WPS = 'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"';
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

describe('preservedBodyPrXml', () => {
  test('keeps a bodyPr whole, attributes and autofit alike', () => {
    const xml = preservedBodyPrXml(
      `<wps:bodyPr ${WPS} ${A} wrap="square" vertOverflow="overflow"><a:spAutoFit/></wps:bodyPr>`
    );
    expect(xml).toContain('vertOverflow="overflow"');
    expect(xml).toContain('<a:spAutoFit/>');
  });

  test('writes the overrides over the source attributes, leaving the rest', () => {
    const xml = preservedBodyPrXml(`<wps:bodyPr ${WPS} lIns="0" wrap="square" compatLnSpc="1"/>`, {
      overrides: { lIns: '91440', anchor: 'ctr' },
    });
    expect(xml).toContain('lIns="91440"');
    expect(xml).toContain('anchor="ctr"');
    expect(xml).toContain('wrap="square"');
    expect(xml).toContain('compatLnSpc="1"');
  });

  test.each([
    ['another element entirely', `<w:p ${W}/>`],
    ['two roots', `<wps:bodyPr ${WPS}/><wps:bodyPr ${WPS}/>`],
    ['a child the schema does not allow there', `<wps:bodyPr ${WPS} ${W}><w:p/></wps:bodyPr>`],
    ['character data beside it', `text<wps:bodyPr ${WPS}/>`],
    ['unbalanced markup', `<wps:bodyPr ${WPS}>`],
    ['a doctype', `<!DOCTYPE x><wps:bodyPr ${WPS}/>`],
    ['nothing at all', ''],
  ])('refuses %s', (_name, input) => {
    expect(preservedBodyPrXml(input)).toBeUndefined();
  });
});

describe('preservedSpPrExtra', () => {
  test('splits the line from the effects so either can be replaced alone', () => {
    const { ln, effectLst } = preservedSpPrExtra(
      `<a:ln ${A} w="6350"><a:noFill/></a:ln><a:effectLst ${A}><a:outerShdw blurRad="50800"/></a:effectLst>`
    );
    expect(ln).toContain('<a:noFill/>');
    expect(effectLst).toContain('<a:outerShdw');
  });

  test('a rogue sibling drops the fragment rather than half of it', () => {
    expect(preservedSpPrExtra(`<a:ln ${A}/><w:p ${W}/>`)).toEqual({});
    expect(preservedSpPrExtraXml(`<a:ln ${A}/><w:p ${W}/>`)).toBeUndefined();
  });

  test('a child that does not belong in a:ln is refused with it', () => {
    expect(preservedSpPrExtra(`<a:ln ${A} ${W}><w:drawing/></a:ln>`)).toEqual({});
  });

  test('joins back in schema order', () => {
    const xml = preservedSpPrExtraXml(`<a:effectLst ${A}/><a:ln ${A}/>`);
    expect(xml!.indexOf('<a:ln')).toBeLessThan(xml!.indexOf('<a:effectLst'));
  });
});

describe('the paste boundary is stricter about prefixes', () => {
  // A source part gets its prefixes from the document root, which the
  // serializer re-declares, so a custom extension inside `a:extLst` is valid
  // markup that binds nothing itself. Pasted HTML has no such root.
  const withExtension =
    `<wps:bodyPr ${WPS} ${A} lIns="0"><a:extLst><a:ext uri="urn:x">` +
    '<custom:data/></a:ext></a:extLst></wps:bodyPr>';

  test('a source fragment may lean on the destination root for a prefix', () => {
    expect(preservedBodyPrXml(withExtension)).toContain('<custom:data/>');
  });

  test('the same fragment pasted in is refused', () => {
    expect(preservedBodyPrXml(withExtension, { requireBoundPrefixes: true })).toBeUndefined();
  });
});

describe('markup that was pretty-printed', () => {
  test('whitespace between children is formatting, not character data', () => {
    const xml = preservedBodyPrXml(
      `<wps:bodyPr ${WPS} ${A} wrap="square">\n  <a:spAutoFit/>\n</wps:bodyPr>`
    );
    expect(xml).toContain('wrap="square"');
    expect(xml).toContain('<a:spAutoFit/>');
  });

  test('whitespace between siblings is too', () => {
    const { ln, effectLst } = preservedSpPrExtra(`<a:ln ${A}/>\n  <a:effectLst ${A}/>`);
    expect(ln).toBeDefined();
    expect(effectLst).toBeDefined();
  });

  test('text that is not whitespace is still refused', () => {
    expect(preservedBodyPrXml(`<wps:bodyPr ${WPS}>hello</wps:bodyPr>`)).toBeUndefined();
  });
});

describe('a shape whose body properties are all unmodelled', () => {
  const shapeXml = (inner: string): XmlElement =>
    parseXml(`<wps:wsp ${WPS} ${A}>${inner}</wps:wsp>`).elements![0];

  test('keeps its bodyPr even with no text and nothing the model parses', () => {
    // `wrap` / `vertOverflow` / `compatLnSpc` have no field on ShapeTextBody,
    // so gating the capture on the modelled fields dropped the element whole.
    const shape = parseShape(
      shapeXml(
        '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></wps:spPr>' +
          '<wps:bodyPr wrap="none" vertOverflow="clip" compatLnSpc="1"/>'
      )
    );
    expect(shape.textBody?.bodyPrXml).toContain('vertOverflow="clip"');
    expect(shape.textBody?.bodyPrXml).toContain('compatLnSpc="1"');
    expect(shape.textBody?.content).toEqual([]);
  });

  test('a shape with no bodyPr at all still has no text body', () => {
    const shape = parseShape(
      shapeXml(
        '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></wps:spPr>'
      )
    );
    expect(shape.textBody).toBeUndefined();
  });
});

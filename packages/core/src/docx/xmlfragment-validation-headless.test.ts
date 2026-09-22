/**
 * The same validation, in a runtime with no `DOMParser`.
 *
 * `DocxReviewer`, server-side generation and the `agents` package all run this
 * code in bare Node, where `xml-js` alone would accept `<w:t>unclosed` and the
 * serializer would write it into `document.xml`. This file deliberately does
 * NOT register happy-dom — the sibling `xmlfragment-validation.test.ts` covers
 * the browser path, and the two must agree.
 */
import { describe, expect, test } from 'bun:test';
import { isWellFormedXmlElement } from './xmlParser';

const nest = (depth: number) => '<w:t>'.repeat(depth) + 'x' + '</w:t>'.repeat(depth);

describe('preserved-fragment validation without a DOM', () => {
  test('no DOMParser is registered in this file', () => {
    expect((globalThis as { DOMParser?: unknown }).DOMParser).toBeUndefined();
  });

  test('rejects unbalanced markup', () => {
    expect(isWellFormedXmlElement('<w:t>unclosed')).toBe(false);
    expect(isWellFormedXmlElement('<w:t>a</w:tt>')).toBe(false);
    expect(isWellFormedXmlElement('<w:t></w:x>')).toBe(false);
    expect(isWellFormedXmlElement('</w:t>')).toBe(false);
    expect(isWellFormedXmlElement('plain text')).toBe(false);
    expect(isWellFormedXmlElement('<w:t/><w:t/>')).toBe(false);
  });

  test('rejects character data outside the root element', () => {
    // Written back verbatim this puts text straight inside a `w:r`.
    expect(isWellFormedXmlElement('garbage<w:t/>')).toBe(false);
    expect(isWellFormedXmlElement('<w:t/>trailing')).toBe(false);
    expect(isWellFormedXmlElement('<w:t/>&amp;')).toBe(false);
    // Surrounding whitespace is fine.
    expect(isWellFormedXmlElement('\n  <w:t/>\n')).toBe(true);
  });

  test('rejects malformed attributes', () => {
    expect(isWellFormedXmlElement('<w:t val=1/>')).toBe(false);
    expect(isWellFormedXmlElement('<w:t val="unterminated/>')).toBe(false);
    expect(isWellFormedXmlElement('<w:t val="a" val="b"/>')).toBe(false);
    expect(isWellFormedXmlElement('<w:t val="<"/>')).toBe(false);
    // XML requires whitespace between attributes.
    expect(isWellFormedXmlElement('<w:t a="1"b="2"/>')).toBe(false);
    expect(isWellFormedXmlElement('<w:t a="1" b="2"/>')).toBe(true);
  });

  test('rejects names that are not XML names', () => {
    expect(isWellFormedXmlElement('<w:t?/>')).toBe(false);
    expect(isWellFormedXmlElement('<w:t*/>')).toBe(false);
    expect(isWellFormedXmlElement('<1w:t/>')).toBe(false);
    expect(isWellFormedXmlElement('<w:t a?="1"/>')).toBe(false);
  });

  test('rejects character references outside the XML character range', () => {
    expect(isWellFormedXmlElement('<w:t>&#x110000;</w:t>')).toBe(false); // past U+10FFFF
    expect(isWellFormedXmlElement('<w:t>&#xD800;</w:t>')).toBe(false); // surrogate half
    expect(isWellFormedXmlElement('<w:t>&#0;</w:t>')).toBe(false); // NUL
    expect(isWellFormedXmlElement('<w:t>&#x41;&#65;&#x1F600;</w:t>')).toBe(true);
  });

  test('rejects a DTD, an entity reference and a processing instruction', () => {
    expect(isWellFormedXmlElement('<!DOCTYPE t [<!ENTITY x SYSTEM "file:///etc/passwd">]>')).toBe(
      false
    );
    expect(isWellFormedXmlElement('<w:t>&xxe;</w:t>')).toBe(false);
    expect(isWellFormedXmlElement('<w:t>bare & ampersand</w:t>')).toBe(false);
    expect(isWellFormedXmlElement('<?xml-stylesheet?><w:t/>')).toBe(false);
  });

  test('rejects nesting past the depth bound instead of truncating the walk', () => {
    expect(isWellFormedXmlElement(nest(60))).toBe(true);
    expect(isWellFormedXmlElement(nest(70))).toBe(false);
  });

  test('rejects a malformed comment', () => {
    expect(isWellFormedXmlElement('<w:t><!-- a -- b --></w:t>')).toBe(false);
    expect(isWellFormedXmlElement('<w:t><!--x---></w:t>')).toBe(false); // content is `x-`
    expect(isWellFormedXmlElement('<w:t><!-- unterminated</w:t>')).toBe(false);
    expect(isWellFormedXmlElement('<w:t><!-- fine --></w:t>')).toBe(true);
  });

  test('accepts what Word actually emits inside a run', () => {
    expect(isWellFormedXmlElement('<w:fldChar w:fldCharType="begin"/>')).toBe(true);
    expect(isWellFormedXmlElement('<w:t xml:space="preserve"> </w:t>')).toBe(true);
    expect(isWellFormedXmlElement('<w:t>&amp;&lt;&#65;&#x41;</w:t>')).toBe(true);
    expect(isWellFormedXmlElement('<w:drawing><a:blip r:embed="rId4"/></w:drawing>')).toBe(true);
    expect(
      isWellFormedXmlElement(
        '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
          '<mc:Choice Requires="wpg"><w:drawing/></mc:Choice></mc:AlternateContent>'
      )
    ).toBe(true);
  });

  test('the paste boundary still narrows prefixes and run content', () => {
    const paste = (xml: string) =>
      isWellFormedXmlElement(xml, { requireBoundPrefixes: true, runContentOnly: true });
    expect(paste('<w:t>hello</w:t>')).toBe(true);
    expect(paste('<w:p><w:r><w:t>x</w:t></w:r></w:p>')).toBe(false);
    expect(paste('<w:instrText xml:space="preserve"> DDEAUTO WinWord </w:instrText>')).toBe(false);
    expect(paste('<zz:custom someAttr="1"/>')).toBe(false);
    expect(paste('<zz:custom xmlns:zz="urn:x"/>')).toBe(false); // not run content
  });
});

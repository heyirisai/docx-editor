/**
 * `isWellFormedXmlElement` gates every preserved fragment on its way into the
 * package, so it has to reject what would break the part without dropping what
 * would have exported fine.
 */
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import { isWellFormedXmlElement } from './xmlParser';

beforeAll(() => GlobalRegistrator.register());
afterAll(() => GlobalRegistrator.unregister());

describe('preserved-fragment validation', () => {
  test('rejects markup that would make Word repair the file', () => {
    expect(isWellFormedXmlElement('<w:t>unclosed')).toBe(false);
    expect(isWellFormedXmlElement('<w:t></w:x>')).toBe(false);
    expect(isWellFormedXmlElement('plain text')).toBe(false);
    expect(isWellFormedXmlElement('')).toBe(false);
    expect(isWellFormedXmlElement(']]>')).toBe(false);
  });

  test('rejects anything that is not exactly one element', () => {
    expect(isWellFormedXmlElement('<w:t/><w:t/>')).toBe(false);
  });

  test('accepts source-derived fragments', () => {
    expect(isWellFormedXmlElement('<w:fldChar w:fldCharType="begin"/>')).toBe(true);
    expect(
      isWellFormedXmlElement(
        '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
          '<mc:Choice Requires="wpg"><w:drawing/></mc:Choice></mc:AlternateContent>'
      )
    ).toBe(true);
  });

  test('accepts a prefix bound only by the destination root', () => {
    // `rootNamespaces` carries these to the serializer, so rejecting here would
    // silently drop markup that exports correctly.
    expect(isWellFormedXmlElement('<a16:creationId xmlns:a16="urn:x" id="1"/>')).toBe(true);
    expect(isWellFormedXmlElement('<zz:custom someAttr="1"/>')).toBe(true);
    expect(isWellFormedXmlElement('<w:drawing><zz:ext val="1"/></w:drawing>')).toBe(true);
  });

  test('the paste boundary additionally requires prefixes to resolve', () => {
    const paste = (xml: string) => isWellFormedXmlElement(xml, { requireBoundPrefixes: true });
    // Nothing declares `zz` at the destination of a paste.
    expect(paste('<zz:custom someAttr="1"/>')).toBe(false);
    expect(paste('<w:drawing><zz:ext val="1"/></w:drawing>')).toBe(false);
    // Self-declared, or a prefix the serializers always emit.
    expect(paste('<zz:custom xmlns:zz="urn:x"/>')).toBe(true);
    expect(paste('<w:fldChar w:fldCharType="begin"/>')).toBe(true);
  });

  test('the paste boundary rejects markup that is not legal run content', () => {
    // Preserved source is written back verbatim INSIDE a `w:r`, so block-level
    // markup there produces a document Word refuses to open.
    const paste = (xml: string) =>
      isWellFormedXmlElement(xml, { requireBoundPrefixes: true, runContentOnly: true });
    expect(paste('<w:p><w:r><w:t>x</w:t></w:r></w:p>')).toBe(false);
    expect(paste('<w:tbl/>')).toBe(false);
    expect(paste('<w:body/>')).toBe(false);
    expect(paste('<w:sectPr/>')).toBe(false);

    // What the parser actually preserves stays accepted.
    expect(paste('<w:fldChar w:fldCharType="separate"/>')).toBe(true);
    expect(paste('<w:instrText xml:space="preserve"> TOC </w:instrText>')).toBe(true);
    expect(
      paste(
        '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
          '<mc:Choice Requires="wpg"><w:drawing/></mc:Choice></mc:AlternateContent>'
      )
    ).toBe(true);
  });
});

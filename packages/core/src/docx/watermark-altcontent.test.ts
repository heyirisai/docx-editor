/**
 * A watermark authored inside `mc:AlternateContent` is reachable from BOTH the
 * header's watermark model and the run parser's preserved-source fallback, and
 * `serializeHeaderFooter` emits both — so one copy has to stand down.
 *
 * It has to stand down in the HEADER only: `extractWatermark` is called from
 * `parseHeader` alone, so dropping the preserved source anywhere else deletes
 * WordArt that nothing else models.
 */
import { describe, expect, test } from 'bun:test';
import { parseHeader, parseFooter } from './headerFooterParser';
import { containsWatermarkShape } from './vmlWatermarkParser';
import { parseXmlDocument, type XmlElement } from './xmlParser';
import type { HeaderFooter, Paragraph } from '../types/document';

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'xmlns:v="urn:schemas-microsoft-com:vml" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"';

const WATERMARK_SHAPE = `<v:shape id="PowerPlusWaterMarkObject1" type="#_x0000_t136" style="rotation:315">
  <v:textpath string="DRAFT" style="font-family:Calibri"/>
</v:shape>`;

const part = (root: 'hdr' | 'ftr', shape: string) =>
  `<?xml version="1.0" encoding="UTF-8"?><w:${root} ${NS}><w:p><w:r><mc:AlternateContent>
    <mc:Choice Requires="wps"><w:drawing/></mc:Choice>
    <mc:Fallback><w:pict>${shape}</w:pict></mc:Fallback>
  </mc:AlternateContent></w:r></w:p></w:${root}>`;

const PLAIN_SHAPE = '<v:shape id="Rectangle 4" style="width:10pt"/>';

function rawXmlCount(hf: HeaderFooter): number {
  let n = 0;
  for (const block of hf.content) {
    if (block.type !== 'paragraph') continue;
    for (const item of (block as Paragraph).content) {
      if (item.type === 'run') n += item.content.filter((c) => c.type === 'rawXml').length;
    }
  }
  return n;
}

describe('a watermark inside mc:AlternateContent', () => {
  test('the shape predicate recognises it', () => {
    expect(containsWatermarkShape(parseXmlDocument(WATERMARK_SHAPE) as XmlElement)).toBe(true);
    expect(containsWatermarkShape(parseXmlDocument(PLAIN_SHAPE) as XmlElement)).toBe(false);
  });

  test('a header models it once and does not also preserve it', () => {
    const hdr = parseHeader(part('hdr', WATERMARK_SHAPE));
    expect(hdr.watermark?.kind).toBe('text');
    expect(rawXmlCount(hdr)).toBe(0);
  });

  test('a FOOTER keeps the preserved source — nothing else models it there', () => {
    const ftr = parseFooter(part('ftr', WATERMARK_SHAPE));
    expect(ftr.watermark).toBeUndefined();
    expect(rawXmlCount(ftr)).toBe(1);
  });

  test('an ordinary shape is preserved in a header too', () => {
    const hdr = parseHeader(part('hdr', PLAIN_SHAPE));
    expect(hdr.watermark).toBeUndefined();
    expect(rawXmlCount(hdr)).toBe(1);
  });
});

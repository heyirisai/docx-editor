/**
 * A watermark authored inside `mc:AlternateContent` is reachable from both the
 * header's watermark model and the run parser's preserved-source fallback.
 * Writing it from both puts two copies in the part.
 */
import { describe, expect, test } from 'bun:test';
import { parseXmlDocument, type XmlElement } from './xmlParser';
import { containsWatermarkShape, extractWatermark } from './vmlWatermarkParser';
import { parseRun } from './runParser';

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'xmlns:v="urn:schemas-microsoft-com:vml" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"';

// Word's WordArt watermark, wrapped so a non-VML consumer gets the shape branch.
const RUN_WITH_WATERMARK = `<w:r ${NS}><mc:AlternateContent>
  <mc:Choice Requires="wps"><w:drawing/></mc:Choice>
  <mc:Fallback><w:pict>
    <v:shape id="PowerPlusWaterMarkObject1" type="#_x0000_t136" style="rotation:315">
      <v:textpath string="DRAFT" style="font-family:Calibri"/>
    </v:shape>
  </w:pict></mc:Fallback>
</mc:AlternateContent></w:r>`;

const RUN_WITH_PLAIN_SHAPE = `<w:r ${NS}><mc:AlternateContent>
  <mc:Fallback><w:pict>
    <v:shape id="Rectangle 4" style="width:10pt"/>
  </w:pict></mc:Fallback>
</mc:AlternateContent></w:r>`;

const el = (xml: string) => parseXmlDocument(xml) as XmlElement;

describe('a watermark inside mc:AlternateContent', () => {
  test('is recognised by both the watermark model and the preserve guard', () => {
    const run = el(RUN_WITH_WATERMARK);
    expect(extractWatermark(run)?.kind).toBe('text');
    expect(containsWatermarkShape(run)).toBe(true);
  });

  test('is not ALSO preserved as raw source', () => {
    const run = parseRun(el(RUN_WITH_WATERMARK), null, null, null, null);
    expect(run.content.filter((c) => c.type === 'rawXml')).toHaveLength(0);
  });

  test('an ordinary shape is still preserved', () => {
    expect(containsWatermarkShape(el(RUN_WITH_PLAIN_SHAPE))).toBe(false);
    const run = parseRun(el(RUN_WITH_PLAIN_SHAPE), null, null, null, null);
    expect(run.content.filter((c) => c.type === 'rawXml').length).toBeGreaterThan(0);
  });
});

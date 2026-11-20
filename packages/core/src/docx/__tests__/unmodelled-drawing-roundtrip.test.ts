/**
 * Drawings with no document model — shapes, groups, drawn VML — must survive
 * a full load → edit-model → save cycle verbatim, while the canvas gets a
 * render-only preview of them. Synthetic in-memory package, no binary fixture.
 */
import { describe, expect, test } from 'bun:test';
import JSZip from 'jszip';
import { parseDocx } from '../parser';
import { repackDocx } from '../rezip';
import { toProseDoc, fromProseDoc } from '../../prosemirror/conversion';
import type { DrawingContent, Paragraph, RawXmlContent } from '../../types/document';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_PR = 'http://schemas.openxmlformats.org/package/2006/relationships';
const NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types';

const ROOT_NS =
  `xmlns:w="${NS_W}" xmlns:r="${NS_R}" ` +
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ' +
  'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" ' +
  'xmlns:v="urn:schemas-microsoft-com:vml" ' +
  'xmlns:o="urn:schemas-microsoft-com:office:office" ' +
  'mc:Ignorable="wps wpg"';

/** A shape whose alt text holds characters that must be escaped on save. */
const SHAPE_WSP =
  `<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="457200"/></a:xfrm>` +
  `<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>` +
  `<a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr><wps:bodyPr/></wps:wsp>`;

function anchoredDrawing(id: number, uri: string, body: string, descr = ''): string {
  return (
    `<w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" ` +
    `behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>` +
    `<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="1828800" cy="457200"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapTopAndBottom/>` +
    `<wp:docPr id="${id}" name="Shape ${id}" descr="${descr}"/><wp:cNvGraphicFramePr/>` +
    `<a:graphic><a:graphicData uri="${uri}">${body}</a:graphicData></a:graphic></wp:anchor></w:drawing>`
  );
}

const WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';

const ALT_CONTENT_SHAPE =
  `<mc:AlternateContent><mc:Choice Requires="wps">` +
  anchoredDrawing(11, WPS, SHAPE_WSP, 'Sales &amp; marketing&#xA;banner') +
  `</mc:Choice><mc:Fallback><w:pict><v:roundrect id="s11" style="position:absolute;width:144pt;height:36pt" fillcolor="#4472c4"/></w:pict></mc:Fallback></mc:AlternateContent>`;

const BARE_SHAPE = anchoredDrawing(12, WPS, SHAPE_WSP.replace('roundRect', 'ellipse'));

const VML_ONLY = `<w:pict><v:rect id="r13" style="position:absolute;margin-left:10pt;margin-top:5pt;width:100pt;height:20pt;z-index:2" fillcolor="#c00000" stroked="f"/></w:pict>`;

/** A text box off the 9525-EMU (1 px) grid, with Word's default 0.1" insets. */
const TEXT_BOX = anchoredDrawing(
  14,
  WPS,
  `<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828801" cy="457211"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></wps:spPr>` +
    `<wps:txbx><w:txbxContent><w:p><w:r><w:t>Delta</w:t></w:r></w:p></w:txbxContent></wps:txbx>` +
    `<wps:bodyPr lIns="91440" tIns="45720" rIns="91440" bIns="45720"/></wps:wsp>`
).replace('<wp:extent cx="1828800" cy="457200"/>', '<wp:extent cx="1828801" cy="457211"/>');

const DOCUMENT_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<w:document ${ROOT_NS}><w:body>` +
  `<w:p><w:r>${ALT_CONTENT_SHAPE}</w:r><w:r><w:t>Alpha</w:t></w:r></w:p>` +
  `<w:p><w:r>${BARE_SHAPE}</w:r><w:r><w:t>Beta</w:t></w:r></w:p>` +
  `<w:p><w:r>${VML_ONLY}</w:r><w:r><w:t>Gamma</w:t></w:r></w:p>` +
  `<w:p><w:r>${TEXT_BOX}</w:r></w:p>` +
  `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr>` +
  `</w:body></w:document>`;

async function buildDocx(): Promise<ArrayBuffer> {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="${NS_CT}">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`
  );
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${NS_PR}">` +
      `<Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="word/document.xml"/></Relationships>`
  );
  zip.file('word/document.xml', DOCUMENT_XML);
  zip.file(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${NS_PR}"></Relationships>`
  );
  return zip.generateAsync({ type: 'arraybuffer' });
}

function runItems(p: Paragraph): Array<DrawingContent | RawXmlContent | { type: string }> {
  return p.content.flatMap((c) => (c.type === 'run' ? c.content : []));
}

async function roundTrip(): Promise<{ xml: string; files: string[] }> {
  const doc = await parseDocx(await buildDocx(), { preloadFonts: false });
  const pm = toProseDoc(doc, { styles: doc.package.styles });
  doc.package.document.content = fromProseDoc(pm, doc).package.document.content;
  const zip = await JSZip.loadAsync(await repackDocx(doc));
  return {
    xml: await zip.file('word/document.xml')!.async('string'),
    files: Object.keys(zip.files),
  };
}

const withoutNs = (s: string) => s.replace(/ xmlns:[A-Za-z0-9]+="[^"]*"/g, '');

describe('unmodelled drawings', () => {
  test('each one parses to preserved source plus a render-only preview at its frame', async () => {
    const doc = await parseDocx(await buildDocx(), { preloadFonts: false });
    const paras = doc.package.document.content as Paragraph[];
    for (const p of paras.slice(0, 3)) {
      const items = runItems(p);
      expect(items.some((c) => c.type === 'rawXml')).toBe(true);
      const previews = items.filter(
        (c): c is DrawingContent => c.type === 'drawing' && !!(c as DrawingContent).image.renderOnly
      );
      expect(previews).toHaveLength(1);
      expect(previews[0].image.src?.startsWith('data:image/svg+xml')).toBe(true);
    }
    const [shape, bare, vml] = paras.map(
      (p) => runItems(p).find((c) => c.type === 'drawing') as DrawingContent
    );
    // The shape's own wrap and extent, so text flows around the real box.
    expect(shape.image.wrap.type).toBe('topAndBottom');
    expect(shape.image.size).toEqual({ width: 1828800, height: 457200 });
    expect(bare.image.wrap.type).toBe('topAndBottom');
    expect(vml.image.size).toEqual({ width: 100 * 12700, height: 20 * 12700 });
    expect(vml.image.wrap.type).toBe('inFront');
  });

  test('save writes every drawing back verbatim and no preview media', async () => {
    const { xml, files } = await roundTrip();
    const body = withoutNs(xml);
    expect(body).toContain(withoutNs(ALT_CONTENT_SHAPE).replace('&#xA;', '&#xA;'));
    expect(body).toContain(BARE_SHAPE);
    expect(body).toContain(VML_ONLY);
    // The previews are paint only: no pictures or media parts were invented.
    expect(xml).not.toContain('<pic:pic');
    expect(files.some((f) => f.startsWith('word/media/'))).toBe(false);
  });

  test('alt text with markup characters and line breaks survives save', async () => {
    const { xml } = await roundTrip();
    expect(xml).toContain('descr="Sales &amp; marketing&#xA;banner"');
  });

  test('a text box keeps its exact extent and insets through save', async () => {
    // Rounding to whole pixels on load moved both onto the 1 px grid on
    // every save (a 0.1" inset came back 0.104").
    const { xml } = await roundTrip();
    expect(xml).toContain('<wp:extent cx="1828801" cy="457211"/>');
    expect(xml).toMatch(/lIns="91440" tIns="45720" rIns="91440" bIns="45720"/);
  });
});

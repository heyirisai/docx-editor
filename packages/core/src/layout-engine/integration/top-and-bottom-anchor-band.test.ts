/**
 * A `wp:anchor` picture with `wrapTopAndBottom` sits at its own anchor offset
 * and breaks the text above and below it (ECMA-376 §20.4.2.20). It does NOT
 * take a block line at its run position: a heading paragraph that anchors a
 * logo grid 24px below its top keeps its text above the grid, and the
 * paragraphs after it — empty ones included — start below the grid.
 *
 * Before the fix the picture was measured as a line of its own where its run
 * sits (first in the paragraph), so the heading text landed BELOW the grid,
 * and empty paragraphs ignored the band entirely.
 *
 * Synthetic fixture: Letter, 1in margins; paragraph 1 = anchored picture
 * (paragraph-relative, posOffset 228600 EMU = 24px, extent 3in × 2in =
 * 288 × 192px) + the text "Heading"; then three empty paragraphs and a text
 * paragraph.
 */

import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import JSZip from 'jszip';
import { installCanvasDocumentStub } from './helpers';
import { layoutDocxHeadless, parseDocxHeadless } from './headlessDocxLayout';
import type { ParagraphMeasure } from '../types';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
const NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const NS_PR = 'http://schemas.openxmlformats.org/package/2006/relationships';

const PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

const BAND_OFFSET_PX = 24;
const PICTURE_HEIGHT_PX = 192;

const picture =
  `<w:r><w:drawing>` +
  `<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" ` +
  `behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
  `<wp:simplePos x="0" y="0"/>` +
  `<wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH>` +
  `<wp:positionV relativeFrom="paragraph"><wp:posOffset>228600</wp:posOffset></wp:positionV>` +
  `<wp:extent cx="2743200" cy="1828800"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
  `<wp:wrapTopAndBottom/><wp:docPr id="1" name="Picture 1"/><wp:cNvGraphicFramePr/>` +
  `<a:graphic xmlns:a="${NS_A}"><a:graphicData uri="${NS_PIC}">` +
  `<pic:pic xmlns:pic="${NS_PIC}"><pic:nvPicPr><pic:cNvPr id="1" name="p.png"/><pic:cNvPicPr/></pic:nvPicPr>` +
  `<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
  `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2743200" cy="1828800"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>` +
  `</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;

const text = (t: string) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;

async function buildDocx(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="${NS_CT}">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Default Extension="png" ContentType="image/png"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `</Types>`
  );
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${NS_PR}">` +
      `<Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="word/document.xml"/></Relationships>`
  );
  zip.file(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${NS_PR}">` +
      `<Relationship Id="rIdImg" Type="${NS_R}/image" Target="media/p.png"/></Relationships>`
  );
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}" xmlns:wp="${NS_WP}"><w:body>` +
      `<w:p>${picture}${text('Heading')}</w:p>` +
      `<w:p/><w:p/><w:p/>` +
      `<w:p>${text('Body text after the grid')}</w:p>` +
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
      `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>` +
      `</w:sectPr></w:body></w:document>`
  );
  zip.file('word/media/p.png', PNG_BASE64, { base64: true });
  return new Uint8Array(await zip.generateAsync({ type: 'arraybuffer' }));
}

describe('topAndBottom anchored picture reserves a band at its own offset (Word parity)', () => {
  let restoreDocument: () => void = () => {};
  let res: Awaited<ReturnType<typeof layoutDocxHeadless>>;
  beforeAll(async () => {
    restoreDocument = installCanvasDocumentStub();
    res = await layoutDocxHeadless(await parseDocxHeadless(await buildDocx()));
  });
  afterAll(() => restoreDocument());

  /** Top of a block's first line on the page (fragment y + any float skip). */
  const firstLineTop = (index: number): number => {
    const block = res.blocks.filter((b) => b.kind === 'paragraph')[index]!;
    const fragment = res.layout.pages[0]!.fragments.find((f) => f.blockId === block.id)!;
    const measure = res.measures[res.blocks.indexOf(block)] as ParagraphMeasure;
    return fragment.y + (measure.lines[0]?.floatSkipBefore ?? 0);
  };

  test('the picture is a positioned float, not a block line', () => {
    const anchor = res.blocks[0];
    expect(anchor?.kind).toBe('paragraph');
    const image = anchor?.kind === 'paragraph' ? anchor.runs.find((r) => r.kind === 'image') : null;
    expect(image && image.kind === 'image' && image.anchoredBand).toBe(true);
    const measure = res.measures[0] as ParagraphMeasure;
    // One text line, sized like text — no 192px image line.
    expect(measure.lines).toHaveLength(1);
    expect(measure.lines[0]!.lineHeight).toBeLessThan(40);
  });

  test('the anchoring heading stays above the band; the paragraphs after start below it', () => {
    const paragraphTop = res.layout.pages[0]!.fragments[0]!.y;
    const bandBottom = paragraphTop + BAND_OFFSET_PX + PICTURE_HEIGHT_PX;
    expect(firstLineTop(0)).toBeCloseTo(paragraphTop, 5);
    for (const index of [1, 2, 3, 4]) {
      expect(firstLineTop(index)).toBeGreaterThanOrEqual(bandBottom - 0.5);
    }
  });
});

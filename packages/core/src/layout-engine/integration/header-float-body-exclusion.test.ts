/**
 * A floating picture in a header/footer that wraps text pushes the body out of
 * its way on every page that uses that header/footer — Word's "the picture in
 * my header pushes my text down". Floats with `wrapNone` (in front of / behind
 * text) displace nothing.
 *
 * Synthetic fixture: A4, `pgMar top=1440 header=720`, and a header holding one
 * empty paragraph with a page-wide anchored picture —
 *   positionH relativeFrom=page posOffset 0,
 *   positionV relativeFrom=paragraph posOffset -552450 EMU (-58px),
 *   extent 8.3in × 1.5in (~797 × 144px),
 * so it runs from just above the page edge down to 134px, while the
 * body's top margin is 1in (96px). Word starts every page's body below the picture; before the
 * fix the body started at the margin and the picture covered its first lines.
 */

import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import JSZip from 'jszip';
import { installCanvasDocumentStub } from './helpers';
import { layoutDocxHeadless, parseDocxHeadless } from './headlessDocxLayout';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
const NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const NS_PR = 'http://schemas.openxmlformats.org/package/2006/relationships';

// 1×1 PNG.
const PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

const EMU_PER_PX = 9525;
const TOP_MARGIN_PX = 1440 / 15;
const HF_DISTANCE_PX = 720 / 15;
const BOTTOM_MARGIN_PX = 1440 / 15;
const A4_HEIGHT_PX = 16838 / 15;
/** Where the page-wide picture ends on the page. */
const PICTURE_BOTTOM_PX = HF_DISTANCE_PX + (-552450 + 1371600) / EMU_PER_PX;

interface Drawing {
  /** `wrapSquare` | `wrapTight` | `wrapThrough` | `wrapTopAndBottom` | `wrapNone`. */
  wrap: string;
  behindDoc?: boolean;
  cx?: number;
  cy?: number;
  posH?: string;
  posV?: string;
  distB?: number;
}

const PAGE_WIDE: Drawing = {
  wrap: 'wrapSquare',
  cx: 7589520,
  cy: 1371600,
  posH: '<wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH>',
  posV: '<wp:positionV relativeFrom="paragraph"><wp:posOffset>-552450</wp:posOffset></wp:positionV>',
};

function anchor(d: Drawing): string {
  const cx = d.cx ?? 7589520;
  const cy = d.cy ?? 1371600;
  const wrapEl =
    d.wrap === 'wrapSquare' || d.wrap === 'wrapTight' || d.wrap === 'wrapThrough'
      ? `<wp:${d.wrap} wrapText="bothSides"/>`
      : `<wp:${d.wrap}/>`;
  return (
    `<w:r><w:drawing>` +
    `<wp:anchor distT="0" distB="${d.distB ?? 0}" distL="114300" distR="114300" simplePos="0" ` +
    `relativeHeight="251658240" behindDoc="${d.behindDoc ? 1 : 0}" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/>` +
    (d.posH ?? PAGE_WIDE.posH) +
    (d.posV ?? PAGE_WIDE.posV) +
    `<wp:extent cx="${cx}" cy="${cy}"/>` +
    `<wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    wrapEl +
    `<wp:docPr id="1" name="Picture 1"/>` +
    `<wp:cNvGraphicFramePr/>` +
    `<a:graphic xmlns:a="${NS_A}"><a:graphicData uri="${NS_PIC}">` +
    `<pic:pic xmlns:pic="${NS_PIC}"><pic:nvPicPr><pic:cNvPr id="1" name="logo.png"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>` +
    `</a:graphicData></a:graphic>` +
    `</wp:anchor></w:drawing></w:r>`
  );
}

function story(tag: 'hdr' | 'ftr', drawing: Drawing | null): string {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:${tag} xmlns:w="${NS_W}" xmlns:r="${NS_R}" xmlns:wp="${NS_WP}">` +
    `<w:p>${drawing ? anchor(drawing) : ''}</w:p>` +
    `</w:${tag}>`
  );
}

function para(text: string): string {
  return `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
}

interface Fixture {
  header?: Drawing | null;
  firstHeader?: Drawing | null;
  footer?: Drawing | null;
  titlePg?: boolean;
  paragraphs?: number;
}

async function buildDocx(f: Fixture): Promise<Uint8Array> {
  const zip = new JSZip();
  const overrides: string[] = [];
  const docRels: string[] = [];
  const refs: string[] = [];
  const addStory = (tag: 'hdr' | 'ftr', name: string, type: string, d: Drawing | null) => {
    const kind = tag === 'hdr' ? 'header' : 'footer';
    overrides.push(
      `<Override PartName="/word/${name}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml"/>`
    );
    docRels.push(`<Relationship Id="rId_${name}" Type="${NS_R}/${kind}" Target="${name}.xml"/>`);
    refs.push(`<w:${kind}Reference w:type="${type}" r:id="rId_${name}"/>`);
    zip.file(`word/${name}.xml`, story(tag, d));
    zip.file(
      `word/_rels/${name}.xml.rels`,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<Relationships xmlns="${NS_PR}">` +
        `<Relationship Id="rIdImg" Type="${NS_R}/image" Target="media/logo.png"/>` +
        `</Relationships>`
    );
  };
  if (f.header !== undefined) addStory('hdr', 'header1', 'default', f.header);
  if (f.firstHeader !== undefined) addStory('hdr', 'header2', 'first', f.firstHeader);
  if (f.footer !== undefined) addStory('ftr', 'footer1', 'default', f.footer);

  const body = Array.from({ length: f.paragraphs ?? 120 }, (_, i) =>
    para(`Question ${i + 1}: describe the control in place.`)
  ).join('');
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Types xmlns="${NS_CT}">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Default Extension="png" ContentType="image/png"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      overrides.join('') +
      `</Types>`
  );
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="${NS_PR}">` +
      `<Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="word/document.xml"/>` +
      `</Relationships>`
  );
  zip.file(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="${NS_PR}">${docRels.join('')}</Relationships>`
  );
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}"><w:body>` +
      para('Synthetic Vendor Form') +
      body +
      `<w:sectPr>${refs.join('')}` +
      `<w:pgSz w:w="11906" w:h="16838"/>` +
      `<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="720" w:footer="720" w:gutter="0"/>` +
      (f.titlePg ? '<w:titlePg/>' : '') +
      `</w:sectPr></w:body></w:document>`
  );
  zip.file('word/media/logo.png', PNG_BASE64, { base64: true });
  return new Uint8Array(await zip.generateAsync({ type: 'arraybuffer' }));
}

async function layoutOf(f: Fixture) {
  const doc = await parseDocxHeadless(await buildDocx(f));
  return (await layoutDocxHeadless(doc)).layout;
}

type Pages = Awaited<ReturnType<typeof layoutOf>>['pages'];

function bodyTops(pages: Pages): number[] {
  return pages.map((p) => Math.min(...p.fragments.map((f) => f.y)));
}

function bodyBottoms(pages: Pages): number[] {
  return pages.map((p) => Math.max(...p.fragments.map((f) => f.y + f.height)));
}

describe('header/footer floats that wrap text displace the body (Word parity)', () => {
  let restoreDocument: () => void = () => {};
  beforeAll(() => {
    restoreDocument = installCanvasDocumentStub();
  });
  afterAll(() => restoreDocument());

  test('baseline: without a header picture every page starts at the top margin', async () => {
    const layout = await layoutOf({ header: null });
    expect(layout.pages.length).toBeGreaterThan(2);
    for (const top of bodyTops(layout.pages)) expect(top).toBeCloseTo(TOP_MARGIN_PX, 5);
  });

  test('a page-wide square-wrapped header picture starts every page below it', async () => {
    const layout = await layoutOf({ header: PAGE_WIDE });
    expect(layout.pages.length).toBeGreaterThan(2);
    for (const top of bodyTops(layout.pages)) expect(top).toBeCloseTo(PICTURE_BOTTOM_PX, 3);
    // The authored margin is untouched; the header still anchors to it.
    expect(layout.pages[0]!.margins.top).toBeCloseTo(TOP_MARGIN_PX, 5);
  });

  test('distB adds to the clearance below the picture', async () => {
    const distB = 10 * EMU_PER_PX;
    const layout = await layoutOf({ header: { ...PAGE_WIDE, distB } });
    expect(bodyTops(layout.pages)[0]).toBeCloseTo(PICTURE_BOTTOM_PX + 10, 3);
  });

  test('tight, through and topAndBottom wraps displace too', async () => {
    for (const wrap of ['wrapTight', 'wrapThrough']) {
      const layout = await layoutOf({ header: { ...PAGE_WIDE, wrap }, paragraphs: 5 });
      expect(bodyTops(layout.pages)[0]).toBeCloseTo(PICTURE_BOTTOM_PX, 3);
    }
    // A topAndBottom picture is also measured as its own line in the header
    // paragraph, which already grows the header band past it; either way the
    // body must start clear of the picture.
    const layout = await layoutOf({ header: { ...PAGE_WIDE, wrap: 'wrapTopAndBottom' } });
    for (const top of bodyTops(layout.pages)) {
      expect(top).toBeGreaterThanOrEqual(PICTURE_BOTTOM_PX - 0.001);
    }
  });

  test('wrapNone pictures (in front of text, behind text) keep the body at the margin', async () => {
    for (const behindDoc of [false, true]) {
      const layout = await layoutOf({
        header: { ...PAGE_WIDE, wrap: 'wrapNone', behindDoc },
        paragraphs: 5,
      });
      expect(bodyTops(layout.pages)[0]).toBeCloseTo(TOP_MARGIN_PX, 5);
    }
  });

  test('a narrow wrapping logo with room beside it does not push the body down', async () => {
    const layout = await layoutOf({
      header: {
        ...PAGE_WIDE,
        cx: 1371600, // 1.5in
        posH: '<wp:positionH relativeFrom="margin"><wp:align>left</wp:align></wp:positionH>',
      },
      paragraphs: 5,
    });
    expect(bodyTops(layout.pages)[0]).toBeCloseTo(TOP_MARGIN_PX, 5);
  });

  test('the smaller body area paginates: content flows onto more pages', async () => {
    const plain = await layoutOf({ header: null, paragraphs: 200 });
    const pushed = await layoutOf({ header: PAGE_WIDE, paragraphs: 200 });
    expect(pushed.pages.length).toBeGreaterThan(plain.pages.length);
    const contentBottom = A4_HEIGHT_PX - BOTTOM_MARGIN_PX;
    for (const bottom of bodyBottoms(pushed.pages)) {
      expect(bottom).toBeLessThanOrEqual(contentBottom + 0.01);
    }
  });

  test('with a different first page, only the pages using the picture header move', async () => {
    const layout = await layoutOf({ header: PAGE_WIDE, firstHeader: null, titlePg: true });
    const tops = bodyTops(layout.pages);
    expect(tops[0]).toBeCloseTo(TOP_MARGIN_PX, 5);
    expect(tops[1]).toBeCloseTo(PICTURE_BOTTOM_PX, 3);
  });

  test('a page-wide square-wrapped footer picture ends the body above it', async () => {
    // Anchored to the footer paragraph and lifted 1.2in, 1in tall: it reaches
    // from ~1in above the bottom margin down into the footer band.
    const lift = 1097280;
    const footer: Drawing = {
      ...PAGE_WIDE,
      cy: 914400,
      posV: `<wp:positionV relativeFrom="paragraph"><wp:posOffset>-${lift}</wp:posOffset></wp:positionV>`,
    };
    const plain = await layoutOf({ footer: null });
    const lifted = await layoutOf({ footer });
    expect(bodyBottoms(lifted.pages)[0]!).toBeLessThan(bodyBottoms(plain.pages)[0]!);
    expect(lifted.pages[0]!.fragments.length).toBeLessThan(plain.pages[0]!.fragments.length);
    // The footer flow starts at most `footer distance` above the page bottom,
    // so the picture's top edge is at most this far down the page.
    const pictureTopAtMost = A4_HEIGHT_PX - HF_DISTANCE_PX - lift / EMU_PER_PX;
    for (const bottom of bodyBottoms(lifted.pages)) {
      expect(bottom).toBeLessThanOrEqual(pictureTopAtMost);
    }
  });
});

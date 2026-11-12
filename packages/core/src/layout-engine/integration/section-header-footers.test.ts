/**
 * Every section paints its OWN header/footer (ECMA-376 §17.10). Before the
 * fix the whole document painted one section's header/footer on every page,
 * so a cover page showed the body sections' letterhead and a page-anchored
 * header picture landed on top of other sections' page-anchored pictures.
 *
 * Synthetic fixture, Letter, four sections, every break `nextPage`:
 *  - S0 (cover): references only a `first` header and has no `w:titlePg`, so
 *    Word shows no header on it; its body holds a page-anchored picture.
 *  - S1: default header A (page-anchored picture, top right) + default footer.
 *  - S2: no references — inherits header A and the footer.
 *  - S3 (final): default header B (eight in-flow lines, taller than the top
 *    margin's header band) and, with `w:titlePg`, first header C; long enough
 *    for two pages.
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

const TOP_MARGIN_PX = 1440 / 15;

/** A picture anchored to the page at (x, y) EMU. */
function pageAnchoredPicture(x: number, y: number, rId: string): string {
  return (
    `<w:r><w:drawing>` +
    `<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" ` +
    `behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/>` +
    `<wp:positionH relativeFrom="page"><wp:posOffset>${x}</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="page"><wp:posOffset>${y}</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="914400" cy="457200"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    `<wp:wrapNone/><wp:docPr id="1" name="Picture 1"/><wp:cNvGraphicFramePr/>` +
    `<a:graphic xmlns:a="${NS_A}"><a:graphicData uri="${NS_PIC}">` +
    `<pic:pic xmlns:pic="${NS_PIC}"><pic:nvPicPr><pic:cNvPr id="1" name="p.png"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>` +
    `</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`
  );
}

function para(text: string, extra = ''): string {
  return `<w:p>${extra}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
}

function story(tag: 'hdr' | 'ftr', body: string): string {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:${tag} xmlns:w="${NS_W}" xmlns:r="${NS_R}" xmlns:wp="${NS_WP}">${body}</w:${tag}>`
  );
}

const PG =
  `<w:pgSz w:w="12240" w:h="15840"/>` +
  `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>`;

/** A paragraph that ends a section (its `w:sectPr` carries the section's own settings). */
function sectionEnd(refs: string): string {
  return `<w:p><w:pPr><w:sectPr>${refs}${PG}</w:sectPr></w:pPr></w:p>`;
}

async function buildDocx(): Promise<Uint8Array> {
  const zip = new JSZip();
  const parts: Array<{ name: string; kind: 'header' | 'footer'; xml: string }> = [
    { name: 'header1', kind: 'header', xml: story('hdr', para('Cover first-page header')) },
    {
      name: 'header2',
      kind: 'header',
      xml: story('hdr', `<w:p>${pageAnchoredPicture(5486400, 228600, 'rIdImg')}</w:p>`),
    },
    {
      name: 'header3',
      kind: 'header',
      xml: story(
        'hdr',
        Array.from({ length: 8 }, (_, i) => para(`Header B line ${i + 1}`)).join('')
      ),
    },
    { name: 'header4', kind: 'header', xml: story('hdr', para('Header C')) },
    { name: 'footer1', kind: 'footer', xml: story('ftr', para('Footer A')) },
  ];
  for (const p of parts) {
    zip.file(`word/${p.name}.xml`, p.xml);
    zip.file(
      `word/_rels/${p.name}.xml.rels`,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<Relationships xmlns="${NS_PR}">` +
        `<Relationship Id="rIdImg" Type="${NS_R}/image" Target="media/p.png"/></Relationships>`
    );
  }
  const ref = (kind: 'header' | 'footer', type: string, name: string) =>
    `<w:${kind}Reference w:type="${type}" r:id="rId_${name}"/>`;

  const body =
    // S0 — cover.
    `<w:p>${pageAnchoredPicture(914400, 4572000, 'rIdBodyImg')}</w:p>` +
    para('Cover title') +
    sectionEnd(ref('header', 'first', 'header1')) +
    // S1.
    `<w:p>${pageAnchoredPicture(914400, 4572000, 'rIdBodyImg')}</w:p>` +
    para('Section one') +
    sectionEnd(ref('header', 'default', 'header2') + ref('footer', 'default', 'footer1')) +
    // S2 — inherits everything.
    para('Section two') +
    sectionEnd('') +
    // S3 (final) — long.
    `<w:p>${pageAnchoredPicture(914400, 4572000, 'rIdBodyImg')}</w:p>` +
    Array.from({ length: 70 }, (_, i) => para(`Final section paragraph ${i + 1}`)).join('');

  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="${NS_CT}">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Default Extension="png" ContentType="image/png"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      parts
        .map(
          (p) =>
            `<Override PartName="/word/${p.name}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${p.kind}+xml"/>`
        )
        .join('') +
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
      parts
        .map(
          (p) =>
            `<Relationship Id="rId_${p.name}" Type="${NS_R}/${p.kind}" Target="${p.name}.xml"/>`
        )
        .join('') +
      `<Relationship Id="rIdBodyImg" Type="${NS_R}/image" Target="media/p.png"/>` +
      `</Relationships>`
  );
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}" xmlns:wp="${NS_WP}"><w:body>` +
      body +
      `<w:sectPr>${ref('header', 'default', 'header3')}${ref('header', 'first', 'header4')}${PG}<w:titlePg/></w:sectPr>` +
      `</w:body></w:document>`
  );
  zip.file('word/media/p.png', PNG_BASE64, { base64: true });
  return new Uint8Array(await zip.generateAsync({ type: 'arraybuffer' }));
}

async function layoutFixture() {
  const doc = await parseDocxHeadless(await buildDocx());
  return layoutDocxHeadless(doc);
}

type Result = Awaited<ReturnType<typeof layoutFixture>>;

/** Text of the first paragraph of a header/footer story, or null. */
function storyText(content: { blocks: unknown[] } | undefined): string | null {
  const first = content?.blocks.find(
    (b): b is { kind: 'paragraph'; runs: Array<{ kind: string; text?: string }> } =>
      (b as { kind?: string }).kind === 'paragraph'
  );
  if (!first) return content ? '' : null;
  return first.runs.map((r) => (r.kind === 'text' ? (r.text ?? '') : '')).join('');
}

describe('each section paints its own header/footer (Word parity)', () => {
  let restoreDocument: () => void = () => {};
  let res: Result;
  beforeAll(async () => {
    restoreDocument = installCanvasDocumentStub();
    res = await layoutFixture();
  });
  afterAll(() => restoreDocument());

  test('every nextPage section starts a page, tagged with its section', () => {
    const pages = res.layout.pages;
    expect(pages.length).toBe(5);
    expect(pages.map((p) => p.sectionIndex)).toEqual([0, 1, 2, 3, 3]);
    expect(pages.map((p) => p.firstPageOfSection === true)).toEqual([
      true,
      true,
      true,
      true,
      false,
    ]);
  });

  test('page-anchored body pictures stay on the page of their section', () => {
    const pageOfImageBlock = res.layout.pages.map((p) =>
      p.fragments.some((f) => {
        const block = res.blocks.find((b) => b.id === f.blockId);
        return (
          block?.kind === 'paragraph' &&
          block.runs.some(
            (r) => r.kind === 'image' && r.position?.horizontal?.relativeTo === 'page'
          )
        );
      })
    );
    expect(pageOfImageBlock).toEqual([true, true, false, true, false]);
  });

  test('pages show their own section part: no `first` header without titlePg, inheritance, titlePg first page', () => {
    const refs = res.layout.pages.map((p) => p.headerFooterRefs);
    expect(refs[0]).toEqual({});
    expect(refs[1]).toEqual({ headerDefault: 'rId_header2', footerDefault: 'rId_footer1' });
    expect(refs[2]).toEqual({ headerDefault: 'rId_header2', footerDefault: 'rId_footer1' });
    expect(refs[3]).toEqual({ headerFirst: 'rId_header4' });
    expect(refs[4]).toEqual({ headerDefault: 'rId_header3', footerDefault: 'rId_footer1' });
  });

  test('the painted stories follow the section', () => {
    const sections = res.sectionHeaderFooters!;
    expect(sections).toHaveLength(4);
    expect(sections[0].header).toBeUndefined();
    expect(sections[0].footer).toBeUndefined();
    expect(storyText(sections[1].header)).toBe('');
    expect(sections[2].header).toBe(sections[1].header);
    expect(storyText(sections[3].header)).toBe('Header B line 1');
    expect(storyText(sections[3].firstPageHeader)).toBe('Header C');
    expect(sections[3].titlePg).toBe(true);
  });

  test("a tall header grows only its own section's top margin", () => {
    const tops = res.layout.pages.map((p) => Math.round(p.margins.top));
    expect(tops.slice(0, 3)).toEqual([TOP_MARGIN_PX, TOP_MARGIN_PX, TOP_MARGIN_PX]);
    expect(tops[4]).toBeGreaterThan(TOP_MARGIN_PX);
  });
});

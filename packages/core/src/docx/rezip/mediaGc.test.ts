/**
 * Media GC (pruneUnreferencedMedia) + the processNewImages "already
 * registered" skip. Together they stop the compounding package bloat where
 * every save re-registered every loaded image (all carry a data: src) and
 * repack preserved every stale copy — a 60s-autosaving collab session grew
 * one proposal to 228MB of media with 9 unique images.
 */

import { describe, test, expect } from 'bun:test';
import JSZip from 'jszip';
import type { BlockContent, Image } from '../../types/content';
import { processNewImages } from './images';
import { pruneUnreferencedMedia } from './mediaGc';

const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const IMAGE_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const HYPERLINK_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';

const PNG_DATA_URL = 'data:image/png;base64,iVBORw0KGgo=';

function rels(entries: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${RELS_NS}">${entries}</Relationships>`;
}

function imageRel(id: string, target: string): string {
  return `<Relationship Id="${id}" Type="${IMAGE_TYPE}" Target="${target}"/>`;
}

async function names(zip: JSZip, prefix: string): Promise<string[]> {
  const out: string[] = [];
  zip.forEach((p, f) => {
    if (!f.dir && p.startsWith(prefix)) out.push(p);
  });
  return out.sort();
}

describe('pruneUnreferencedMedia', () => {
  test('removes media whose only rel is unused in the owning part', async () => {
    const zip = new JSZip();
    zip.file('word/document.xml', '<w:document><w:drawing r:embed="rId1"/></w:document>');
    zip.file(
      'word/_rels/document.xml.rels',
      rels(imageRel('rId1', 'media/image1.png') + imageRel('rId2', 'media/image2.png'))
    );
    zip.file('word/media/image1.png', 'live');
    zip.file('word/media/image2.png', 'stale');

    await pruneUnreferencedMedia(zip, 6);

    expect(await names(zip, 'word/media/')).toEqual(['word/media/image1.png']);
    const relsXml = await zip.file('word/_rels/document.xml.rels')!.async('text');
    expect(relsXml).toContain('rId1');
    expect(relsXml).not.toContain('rId2');
  });

  test('keeps media referenced from headers via VML r:id', async () => {
    const zip = new JSZip();
    zip.file('word/document.xml', '<w:document/>');
    zip.file('word/_rels/document.xml.rels', rels(''));
    zip.file('word/header1.xml', '<w:hdr><v:imagedata r:id="rId7"/></w:hdr>');
    zip.file('word/_rels/header1.xml.rels', rels(imageRel('rId7', 'media/image3.jpeg')));
    zip.file('word/media/image3.jpeg', 'banner');

    await pruneUnreferencedMedia(zip, 6);

    expect(await names(zip, 'word/media/')).toEqual(['word/media/image3.jpeg']);
  });

  test('a rels file with no readable owning part keeps everything it references', async () => {
    const zip = new JSZip();
    zip.file('word/document.xml', '<w:document/>');
    zip.file('word/_rels/document.xml.rels', rels(''));
    // Owner 'word/mystery.xml' does not exist in the zip.
    zip.file('word/_rels/mystery.xml.rels', rels(imageRel('rId1', 'media/image9.png')));
    zip.file('word/media/image9.png', 'keep');

    await pruneUnreferencedMedia(zip, 6);

    expect(await names(zip, 'word/media/')).toEqual(['word/media/image9.png']);
  });

  test('never touches non-image relationships', async () => {
    const zip = new JSZip();
    zip.file('word/document.xml', '<w:document/>');
    zip.file(
      'word/_rels/document.xml.rels',
      rels(
        `<Relationship Id="rId9" Type="${HYPERLINK_TYPE}" Target="https://example.com" TargetMode="External"/>`
      )
    );
    zip.file('word/media/image1.png', 'stale-no-rel');

    await pruneUnreferencedMedia(zip, 6);

    const relsXml = await zip.file('word/_rels/document.xml.rels')!.async('text');
    expect(relsXml).toContain('rId9');
    // Media with no rel at all is dead weight — removed.
    expect(await names(zip, 'word/media/')).toEqual([]);
  });

  test('collapses the compounded-duplicates shape: N copies, one used', async () => {
    const zip = new JSZip();
    zip.file('word/document.xml', '<w:document><w:drawing r:embed="rId5"/></w:document>');
    let entries = imageRel('rId5', 'media/image5.jpeg');
    zip.file('word/media/image5.jpeg', 'used');
    // 20 stale generations of the same banner, each with its own rel.
    for (let i = 0; i < 20; i++) {
      entries += imageRel(`rId${10 + i}`, `media/image${10 + i}.jpeg`);
      zip.file(`word/media/image${10 + i}.jpeg`, `stale-${i}`);
    }
    zip.file('word/_rels/document.xml.rels', rels(entries));

    await pruneUnreferencedMedia(zip, 6);

    expect(await names(zip, 'word/media/')).toEqual(['word/media/image5.jpeg']);
  });
});

describe('processNewImages — already-registered skip', () => {
  function drawingPara(image: Image): BlockContent {
    return {
      type: 'paragraph',
      content: [{ type: 'run', content: [{ type: 'drawing', image }] }],
    } as unknown as BlockContent;
  }

  test('an image whose rId resolves in the part rels is NOT re-added', async () => {
    const zip = new JSZip();
    zip.file('word/_rels/document.xml.rels', rels(imageRel('rId3', 'media/image1.png')));
    zip.file('word/media/image1.png', 'original');
    const image = { type: 'image', rId: 'rId3', src: PNG_DATA_URL } as unknown as Image;

    await processNewImages(
      [{ relsPath: 'word/_rels/document.xml.rels', blocks: [drawingPara(image)] }],
      zip,
      6
    );

    expect(await names(zip, 'word/media/')).toEqual(['word/media/image1.png']);
    expect(image.rId).toBe('rId3');
  });

  test('an image with no rId is still registered', async () => {
    const zip = new JSZip();
    zip.file('word/_rels/document.xml.rels', rels(''));
    zip.file('[Content_Types].xml', '<Types></Types>');
    const image = { type: 'image', rId: '', src: PNG_DATA_URL } as unknown as Image;

    await processNewImages(
      [{ relsPath: 'word/_rels/document.xml.rels', blocks: [drawingPara(image)] }],
      zip,
      6
    );

    expect((await names(zip, 'word/media/')).length).toBe(1);
    expect(image.rId).toMatch(/^rId\d+$/);
  });

  test('two new images with identical bytes share one media file', async () => {
    const zip = new JSZip();
    zip.file('word/_rels/document.xml.rels', rels(''));
    zip.file('[Content_Types].xml', '<Types></Types>');
    const a = { type: 'image', rId: '', src: PNG_DATA_URL } as unknown as Image;
    const b = { type: 'image', rId: '', src: PNG_DATA_URL } as unknown as Image;

    await processNewImages(
      [{ relsPath: 'word/_rels/document.xml.rels', blocks: [drawingPara(a), drawingPara(b)] }],
      zip,
      6
    );

    expect((await names(zip, 'word/media/')).length).toBe(1);
    expect(a.rId).toBe(b.rId);
  });

  test('a stale rId that does not resolve is treated as new', async () => {
    const zip = new JSZip();
    zip.file('word/_rels/document.xml.rels', rels(''));
    zip.file('[Content_Types].xml', '<Types></Types>');
    const image = { type: 'image', rId: 'rId99', src: PNG_DATA_URL } as unknown as Image;

    await processNewImages(
      [{ relsPath: 'word/_rels/document.xml.rels', blocks: [drawingPara(image)] }],
      zip,
      6
    );

    expect((await names(zip, 'word/media/')).length).toBe(1);
    expect(image.rId).not.toBe('rId99');
  });
});

import { describe, expect, test } from 'bun:test';
import JSZip from 'jszip';
import { toProseDoc } from '../../prosemirror/conversion/toProseDoc';
import { fromProseDoc } from '../../prosemirror/conversion/fromProseDoc';
import type { DrawingContent, Paragraph } from '../../types/document';
import { IncompleteExternalMediaManifestError, parseDocx } from '../parser';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
const NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const NS_PR = 'http://schemas.openxmlformats.org/package/2006/relationships';

const PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

const CONTENT_TYPES =
  `<?xml version="1.0" encoding="UTF-8"?>` +
  `<Types xmlns="${NS_CT}">` +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  '<Default Extension="png" ContentType="image/png"/>' +
  '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
  '</Types>';

const PACKAGE_RELS =
  `<?xml version="1.0" encoding="UTF-8"?>` +
  `<Relationships xmlns="${NS_PR}">` +
  `<Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="word/document.xml"/>` +
  '</Relationships>';

const DOCUMENT_XML =
  `<?xml version="1.0" encoding="UTF-8"?>` +
  `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}" xmlns:wp="${NS_WP}" xmlns:a="${NS_A}" xmlns:pic="${NS_PIC}">` +
  '<w:body><w:p><w:r><w:drawing><wp:inline>' +
  '<wp:extent cx="914400" cy="914400"/>' +
  '<wp:docPr id="1" name="Picture 1" descr="Logo"/>' +
  '<a:graphic><a:graphicData>' +
  '<pic:pic><pic:blipFill><a:blip r:embed="rIdImage"/></pic:blipFill>' +
  '<pic:spPr><a:xfrm><a:ext cx="914400" cy="914400"/></a:xfrm></pic:spPr></pic:pic>' +
  '</a:graphicData></a:graphic>' +
  '</wp:inline></w:drawing></w:r></w:p>' +
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>' +
  '</w:body></w:document>';

const DOCUMENT_RELS =
  `<?xml version="1.0" encoding="UTF-8"?>` +
  `<Relationships xmlns="${NS_PR}">` +
  `<Relationship Id="rIdImage" Type="${NS_R}/image" Target="media/logo.png"/>` +
  '</Relationships>';

async function buildDocx(): Promise<ArrayBuffer> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CONTENT_TYPES);
  zip.file('_rels/.rels', PACKAGE_RELS);
  zip.file('word/document.xml', DOCUMENT_XML);
  zip.file('word/_rels/document.xml.rels', DOCUMENT_RELS);
  zip.file('word/media/logo.png', PNG_BASE64, { base64: true });
  return zip.generateAsync({ type: 'arraybuffer' });
}

async function buildDocxWithPartImages(): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(await buildDocx());
  const imageRelationships = (entries: readonly [string, string][]) =>
    `<?xml version="1.0" encoding="UTF-8"?>` +
    `<Relationships xmlns="${NS_PR}">` +
    entries
      .map(([id, target]) => `<Relationship Id="${id}" Type="${NS_R}/image" Target="${target}"/>`)
      .join('') +
    '</Relationships>';

  zip.file(
    'word/_rels/header1.xml.rels',
    imageRelationships([['rIdWatermark', 'media/watermark.png']])
  );
  zip.file('word/_rels/footer1.xml.rels', imageRelationships([['rIdFooter', 'media/footer.png']]));
  zip.file('word/media/watermark.png', PNG_BASE64, { base64: true });
  zip.file('word/media/footer.png', PNG_BASE64, { base64: true });
  return zip.generateAsync({ type: 'arraybuffer' });
}

function firstBodyImage(document: Awaited<ReturnType<typeof parseDocx>>): DrawingContent['image'] {
  const paragraph = document.package.document.content[0] as Paragraph;
  const run = paragraph.content[0];
  if (run.type !== 'run') throw new Error('Expected an image run');
  const drawing = run.content[0] as DrawingContent;
  return drawing.image;
}

describe('external media parsing', () => {
  test('keeps image bytes and data URLs out of the parsed document and ProseMirror JSON', async () => {
    const document = await parseDocx(await buildDocx(), {
      preloadFonts: false,
      externalMedia: {
        entries: [
          {
            assetId: 'asset-logo',
            path: 'word/media/logo.png',
            mimeType: 'image/png',
          },
        ],
      },
    });

    const image = firstBodyImage(document);
    expect(image.assetId).toBe('asset-logo');
    expect(image.src).toBeUndefined();

    const media = document.package.media?.get('word/media/logo.png');
    expect(media?.assetId).toBe('asset-logo');
    expect(media?.data).toBeUndefined();
    expect(media?.dataUrl).toBeUndefined();

    const proseDocument = toProseDoc(document);
    const proseJson = proseDocument.toJSON();
    expect(JSON.stringify(proseJson)).not.toContain('data:image');

    let proseImage: Record<string, unknown> | undefined;
    proseDocument.descendants((node) => {
      if (node.type.name === 'image') proseImage = node.attrs as Record<string, unknown>;
    });
    expect(proseImage?.assetId).toBe('asset-logo');
    expect(proseImage?.src).toBe('');

    const roundTripped = fromProseDoc(proseDocument);
    expect(firstBodyImage(roundTripped).assetId).toBe('asset-logo');
    expect(firstBodyImage(roundTripped).src).toBeUndefined();
  });

  test('preserves the existing eager media behavior when external media is not requested', async () => {
    const document = await parseDocx(await buildDocx(), { preloadFonts: false });
    const image = firstBodyImage(document);

    expect(image.assetId).toBeUndefined();
    expect(image.src).toStartWith('data:image/png;base64,');
    expect(document.package.media?.get('word/media/logo.png')?.data).toBeInstanceOf(ArrayBuffer);
  });

  test('rejects manifest paths that can escape the DOCX media namespace', async () => {
    await expect(
      parseDocx(await buildDocx(), {
        preloadFonts: false,
        externalMedia: {
          entries: [
            {
              assetId: 'asset-logo',
              path: '../word/media/logo.png',
              mimeType: 'image/png',
            },
          ],
        },
      })
    ).rejects.toThrow('Invalid external media path');
  });

  test('rejects one asset ID claiming two different media paths', async () => {
    // Resolver caches and the export asset map key by assetId, so aliasing two
    // images onto one ID would embed the wrong bytes rather than fail loudly.
    await expect(
      parseDocx(await buildDocx(), {
        preloadFonts: false,
        externalMedia: {
          entries: [
            {
              assetId: 'asset-shared',
              path: 'word/media/logo.png',
              mimeType: 'image/png',
            },
            {
              assetId: 'asset-shared',
              path: 'word/media/other.png',
              mimeType: 'image/png',
            },
          ],
        },
      })
    ).rejects.toThrow('Duplicate external media asset ID');
  });

  test('allows an asset ID repeated with an identical path and MIME type', async () => {
    const document = await parseDocx(await buildDocx(), {
      preloadFonts: false,
      externalMedia: {
        entries: [
          {
            assetId: 'asset-logo',
            path: 'word/media/logo.png',
            mimeType: 'image/png',
          },
          {
            assetId: 'asset-logo',
            path: 'word/media/logo.png',
            mimeType: 'image/png',
          },
        ],
      },
    });

    expect(firstBodyImage(document).assetId).toBe('asset-logo');
  });

  test('rejects a manifest missing images referenced by non-body package parts', async () => {
    const parsing = parseDocx(await buildDocxWithPartImages(), {
      preloadFonts: false,
      externalMedia: {
        entries: [
          {
            assetId: 'asset-logo',
            path: 'word/media/logo.png',
            mimeType: 'image/png',
          },
        ],
      },
    });

    await expect(parsing).rejects.toBeInstanceOf(IncompleteExternalMediaManifestError);
    await expect(parsing).rejects.toMatchObject({
      missingPaths: ['word/media/footer.png', 'word/media/watermark.png'],
    });
  });

  test('accepts a complete manifest across body, header, and footer relationships', async () => {
    const document = await parseDocx(await buildDocxWithPartImages(), {
      preloadFonts: false,
      externalMedia: {
        entries: [
          {
            assetId: 'asset-logo',
            path: 'word/media/logo.png',
            mimeType: 'image/png',
          },
          {
            assetId: 'asset-watermark',
            path: 'word/media/watermark.png',
            mimeType: 'image/png',
          },
          {
            assetId: 'asset-footer',
            path: 'word/media/footer.png',
            mimeType: 'image/png',
          },
        ],
      },
    });

    expect(document.package.media?.get('word/media/watermark.png')?.assetId).toBe(
      'asset-watermark'
    );
    expect(document.package.media?.get('word/media/footer.png')?.assetId).toBe('asset-footer');
  });
});

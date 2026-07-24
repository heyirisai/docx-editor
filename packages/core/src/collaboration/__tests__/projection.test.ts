import { describe, expect, test } from 'bun:test';
import { schema } from '../../prosemirror/schema';
import { fromProseDoc } from '../../prosemirror/conversion/fromProseDoc';
import { toProseDoc } from '../../prosemirror/conversion/toProseDoc';
import type { Paragraph } from '../../types/document';
import { projectProseMirrorDocument, rehydrateCollaborationDocument } from '../projection';
import {
  CollaborationExportFidelityError,
  CollaborationIdentityMismatchError,
  exportCollaborationDocument,
  IncompleteCollaborationManifestError,
  MissingCollaborationAssetError,
  parseCollaborationPackage,
} from '../workerApi';
import { createDocx } from '../../docx/rezip';
import { parseDocx } from '../../docx/parser';
import { extractParagraphXml } from '../../docx/selectiveXmlPatch';
import type { Document } from '../../types/document';
import type { Comment } from '../../types/content';
import JSZip from 'jszip';

function imageDocument(src: string, assetId = 'asset-logo', rId = 'rId7') {
  const image = schema.nodes.image.create({
    src,
    assetId,
    rId,
    width: 120,
    height: 40,
  });
  const paragraph = schema.nodes.paragraph.create(
    {
      paraId: 'A1B2C3D4',
      styleId: 'Heading1',
      _originalFormatting: {
        styleId: 'Heading1',
        spacing: { before: 240 },
      },
      _originalRunBoundaries: [
        {
          text: '',
          formatting: { bold: true },
        },
      ],
    },
    [image]
  );
  return schema.nodes.doc.create({}, [paragraph]);
}

async function sha256Hex(buffer: ArrayBuffer): Promise<string> {
  const digest = await crypto.subtle.digest('SHA-256', buffer);
  return [...new Uint8Array(digest)].map((byte) => byte.toString(16).padStart(2, '0')).join('');
}

async function createVmlSourceFixture(): Promise<ArrayBuffer> {
  const initial = await createDocx(
    fromProseDoc(
      schema.nodes.doc.create({}, [
        schema.nodes.paragraph.create({ paraId: '11B2C3D4' }, schema.text('First')),
        schema.nodes.paragraph.create({ paraId: '21B2C3D4' }, schema.text('Second')),
      ])
    )
  );
  const zip = await JSZip.loadAsync(initial);
  const documentFile = zip.file('word/document.xml');
  if (!documentFile) throw new Error('Expected document.xml');
  const documentXml = await documentFile.async('text');
  const firstParagraph = extractParagraphXml(documentXml, '11B2C3D4');
  if (!firstParagraph) throw new Error('Expected stable first paragraph');
  const vmlRun =
    '<w:r><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml" id="source-vml-shape" style="width:12pt;height:12pt"/></w:pict></w:r>';
  const vmlParagraph = firstParagraph.replace('</w:p>', `${vmlRun}</w:p>`);
  zip.file('word/document.xml', documentXml.replace(firstParagraph, vmlParagraph));
  return zip.generateAsync({ type: 'arraybuffer' });
}

async function projectSource(sourceBuffer: ArrayBuffer) {
  const parsedSource = await parseDocx(sourceBuffer, {
    externalMedia: { entries: [] },
  });
  return projectProseMirrorDocument(
    toProseDoc(parsedSource, { styles: parsedSource.package.styles }),
    { baseRevision: 'base-1' }
  );
}

function comment(
  id: number,
  text: string,
  options: { parentId?: number; done?: boolean } = {}
): Comment {
  return {
    id,
    author: 'Reviewer',
    initials: 'R',
    date: '2026-07-23T12:00:00Z',
    content: [
      {
        type: 'paragraph',
        content: [{ type: 'run', content: [{ type: 'text', text }] }],
      },
    ],
    ...options,
  };
}

async function createCommentSourceFixture(): Promise<ArrayBuffer> {
  const commentMark = schema.marks.comment.create({ commentId: 1 });
  const pmDocument = schema.nodes.doc.create({}, [
    schema.nodes.paragraph.create(
      { paraId: '41B2C3D4' },
      schema.text('Commented text', [commentMark])
    ),
  ]);
  const document = fromProseDoc(pmDocument);
  document.package.document.comments = [comment(1, 'Original comment')];
  return createDocx(document);
}

async function createHeaderImageRelationshipFixture(): Promise<ArrayBuffer> {
  const sourceBuffer = await createDocx({
    package: {
      document: {
        content: [{ type: 'paragraph', content: [] }],
      },
    },
  });
  const zip = await JSZip.loadAsync(sourceBuffer);
  zip.file(
    'word/_rels/header1.xml.rels',
    '<?xml version="1.0" encoding="UTF-8"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rIdWatermark" ' +
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" ' +
      'Target="media/watermark.png"/>' +
      '</Relationships>'
  );
  zip.file(
    'word/media/watermark.png',
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
    { base64: true }
  );
  return zip.generateAsync({ type: 'arraybuffer' });
}

describe('collaboration projection', () => {
  test('removes media payloads and fidelity-only attrs while preserving stable asset identity', () => {
    const projected = projectProseMirrorDocument(
      imageDocument(`data:image/png;base64,${'A'.repeat(100_000)}`),
      { baseRevision: 'base-1' }
    );
    const encoded = JSON.stringify(projected.document);
    const paragraph = projected.document.content?.[0];
    const image = paragraph?.content?.[0];

    expect(encoded).not.toContain('data:image');
    expect(encoded).not.toContain('_originalFormatting');
    expect(encoded).not.toContain('_originalRunBoundaries');
    expect(image?.attrs?.assetId).toBe('asset-logo');
    expect(image?.attrs?.src).toBe('');
    expect(image?.attrs?.rId).toBeUndefined();
    expect(typeof paragraph?.attrs?.collaborationId).toBe('string');
    expect(projected.sidecar.entries).not.toEqual({});
    expect(projected.diagnostics).toEqual([]);
  });

  test('keeps distinct fidelity entries for nodes that collide under the legacy 32-bit hash', () => {
    const document = schema.nodes.doc.create({}, [
      schema.nodes.paragraph.create(
        {
          paraId: '182451AA',
          _originalFormatting: { styleId: 'FirstStyle' },
        },
        schema.text('First')
      ),
      schema.nodes.paragraph.create(
        {
          paraId: '0366CC5B',
          _originalFormatting: { styleId: 'SecondStyle' },
        },
        schema.text('Second')
      ),
    ]);

    const projected = projectProseMirrorDocument(document, { baseRevision: 'base-1' });
    const ids = projected.document.content?.map((node) => node.attrs?.collaborationId);
    const hydrated = rehydrateCollaborationDocument(projected.document, projected.sidecar);

    expect(ids?.[0]).not.toBe(ids?.[1]);
    expect(Object.keys(projected.sidecar.entries)).toHaveLength(2);
    expect(hydrated.content?.[0]?.attrs?._originalFormatting).toEqual({
      styleId: 'FirstStyle',
    });
    expect(hydrated.content?.[1]?.attrs?._originalFormatting).toEqual({
      styleId: 'SecondStyle',
    });
  });

  test('rejects duplicate caller-supplied collaboration IDs', () => {
    const document = schema.nodes.doc.create({}, [
      schema.nodes.paragraph.create(
        {
          paraId: '11111111',
          collaborationId: 'duplicate-id',
          _originalFormatting: { styleId: 'FirstStyle' },
        },
        schema.text('First')
      ),
      schema.nodes.paragraph.create(
        {
          paraId: '22222222',
          collaborationId: 'duplicate-id',
          _originalFormatting: { styleId: 'SecondStyle' },
        },
        schema.text('Second')
      ),
    ]);

    expect(() => projectProseMirrorDocument(document, { baseRevision: 'base-1' })).toThrow(
      'Duplicate collaboration ID: duplicate-id'
    );
  });

  test('projected state size is independent of original image byte size', () => {
    const small = projectProseMirrorDocument(
      imageDocument(`data:image/png;base64,${'A'.repeat(100_000)}`),
      { baseRevision: 'base-1' }
    );
    const large = projectProseMirrorDocument(
      imageDocument(`data:image/png;base64,${'B'.repeat(20_000_000)}`),
      { baseRevision: 'base-1' }
    );

    expect(JSON.stringify(small.document)).toBe(JSON.stringify(large.document));
    expect(small.stats.projectedJsonBytes).toBe(large.stats.projectedJsonBytes);
  });

  test('rehydrates fidelity attrs before converting back to the document model', () => {
    const source = imageDocument('data:image/png;base64,AAAA');
    const projected = projectProseMirrorDocument(source, { baseRevision: 'base-1' });
    const hydrated = rehydrateCollaborationDocument(projected.document, projected.sidecar);
    const pmDocument = schema.nodeFromJSON(hydrated);
    const output = fromProseDoc(pmDocument);
    const paragraph = output.package.document.content[0] as Paragraph;
    const drawing = paragraph.content
      .flatMap((item) => (item.type === 'run' ? item.content : []))
      .find((item) => item.type === 'drawing');

    expect(paragraph.formatting?.styleId).toBe('Heading1');
    expect(drawing?.type).toBe('drawing');
    expect(drawing?.type === 'drawing' ? drawing.image.rId : undefined).toBe('rId7');
    expect(drawing?.type === 'drawing' ? drawing.image.assetId : undefined).toBe('asset-logo');
  });

  test('fails eligibility without allowing an inline payload into projected state', () => {
    const projected = projectProseMirrorDocument(imageDocument('data:image/png;base64,AAAA', ''), {
      baseRevision: 'base-1',
    });

    expect(JSON.stringify(projected.document)).not.toContain('data:image');
    expect(projected.diagnostics).toEqual([
      {
        code: 'image_asset_id_missing',
        severity: 'error',
        nodeType: 'image',
      },
    ]);
  });

  test('requires original bytes for a newly inserted asset and exports them when provided', async () => {
    const sourceBuffer = await createDocx({
      package: {
        document: {
          content: [
            {
              type: 'paragraph',
              paraId: '11B2C3D4',
              content: [],
            },
          ],
        },
      },
    });
    const parsedSource = await parseDocx(sourceBuffer, {
      externalMedia: { entries: [] },
    });
    const projected = projectProseMirrorDocument(
      toProseDoc(parsedSource, { styles: parsedSource.package.styles }),
      {
        baseRevision: 'base-1',
      }
    );
    const projectedDocument = structuredClone(projected.document);
    const firstParagraph = projectedDocument.content?.[0];
    if (!firstParagraph) throw new Error('Expected source paragraph');
    firstParagraph.content = [
      schema.nodes.image.create({ src: '', assetId: 'new-asset', rId: '' }).toJSON(),
    ];

    await expect(
      exportCollaborationDocument({
        sourceBuffer,
        sourceSha256: await sha256Hex(sourceBuffer),
        baseRevision: 'base-1',
        manifest: [],
        projectedDocument,
        sidecar: projected.sidecar,
      })
    ).rejects.toBeInstanceOf(MissingCollaborationAssetError);

    const pngBytes = Uint8Array.from(
      atob(
        'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='
      ),
      (character) => character.charCodeAt(0)
    );
    const exported = await exportCollaborationDocument({
      sourceBuffer,
      sourceSha256: await sha256Hex(sourceBuffer),
      baseRevision: 'base-1',
      manifest: [],
      projectedDocument,
      sidecar: projected.sidecar,
      assets: [
        {
          assetId: 'new-asset',
          mimeType: 'image/png',
          bytes: pngBytes,
        },
      ],
    });
    const reparsed = await parseDocx(exported);
    const imageSources: string[] = [];
    toProseDoc(reparsed).descendants((node) => {
      if (node.type.name === 'image') imageSources.push(String(node.attrs.src));
    });

    expect(imageSources).toHaveLength(1);
    expect(imageSources[0]).toContain('data:image/png;base64,');
  });

  test('preserves source media that an unsupported OOXML construct may still reference', async () => {
    const sourceDocument: Document = {
      package: {
        document: {
          content: [{ type: 'paragraph', content: [] }],
        },
      },
    };
    const initial = await createDocx(sourceDocument);
    const sourceZip = await JSZip.loadAsync(initial);
    sourceZip.file('word/media/unsupported.svg', '<svg xmlns="http://www.w3.org/2000/svg"/>');
    const sourceBuffer = await sourceZip.generateAsync({ type: 'arraybuffer' });
    const parsedSource = await parseDocx(sourceBuffer, {
      externalMedia: { entries: [] },
    });
    const projected = projectProseMirrorDocument(
      toProseDoc(parsedSource, { styles: parsedSource.package.styles }),
      {
        baseRevision: 'base-1',
      }
    );

    const exported = await exportCollaborationDocument({
      sourceBuffer,
      sourceSha256: await sha256Hex(sourceBuffer),
      baseRevision: 'base-1',
      manifest: [],
      projectedDocument: projected.document,
      sidecar: projected.sidecar,
    });
    const exportedZip = await JSZip.loadAsync(exported);

    expect(exportedZip.file('word/media/unsupported.svg')).not.toBeNull();
  });

  test('preserves source VML exactly on no-op export and when editing another paragraph', async () => {
    const sourceBuffer = await createVmlSourceFixture();
    const projection = await projectSource(sourceBuffer);
    const sourceSha256 = await sha256Hex(sourceBuffer);

    const noOpExport = await exportCollaborationDocument({
      sourceBuffer,
      sourceSha256,
      baseRevision: 'base-1',
      manifest: [],
      projectedDocument: projection.document,
      sidecar: projection.sidecar,
    });
    expect(new Uint8Array(noOpExport)).toEqual(new Uint8Array(sourceBuffer));

    const editedDocument = structuredClone(projection.document);
    const secondParagraph = editedDocument.content?.find(
      (node) => node.attrs?.paraId === '21B2C3D4'
    );
    if (!secondParagraph) throw new Error('Expected second paragraph');
    secondParagraph.content = [{ type: 'text', text: 'Second edited' }];

    const editedExport = await exportCollaborationDocument({
      sourceBuffer,
      sourceSha256,
      baseRevision: 'base-1',
      manifest: [],
      projectedDocument: editedDocument,
      sidecar: projection.sidecar,
    });
    const [sourceZip, editedZip] = await Promise.all([
      JSZip.loadAsync(sourceBuffer),
      JSZip.loadAsync(editedExport),
    ]);
    const [sourceXml, editedXml] = await Promise.all([
      sourceZip.file('word/document.xml')?.async('text'),
      editedZip.file('word/document.xml')?.async('text'),
    ]);
    if (!sourceXml || !editedXml) throw new Error('Expected document XML');

    expect(extractParagraphXml(editedXml, '11B2C3D4')).toBe(
      extractParagraphXml(sourceXml, '11B2C3D4')
    );
    expect(editedXml.match(/<v:shape\b/g)).toHaveLength(1);
    expect(extractParagraphXml(editedXml, '21B2C3D4')).toContain('Second edited');
  });

  test('exports controlled comment edits, replies, resolution, and deletion', async () => {
    const sourceBuffer = await createCommentSourceFixture();
    const projection = await projectSource(sourceBuffer);
    const sourceSha256 = await sha256Hex(sourceBuffer);
    const baseComments = (await parseDocx(sourceBuffer)).package.document.comments ?? [];

    const noOpExport = await exportCollaborationDocument({
      sourceBuffer,
      sourceSha256,
      baseRevision: 'base-1',
      manifest: [],
      projectedDocument: projection.document,
      sidecar: projection.sidecar,
      comments: baseComments,
    });
    expect(new Uint8Array(noOpExport)).toEqual(new Uint8Array(sourceBuffer));

    const textEditedDocument = structuredClone(projection.document);
    const textNode = textEditedDocument.content?.[0]?.content?.[0];
    if (!textNode) throw new Error('Expected commented text');
    textNode.text = 'Edited body text';
    const textEditedExport = await exportCollaborationDocument({
      sourceBuffer,
      sourceSha256,
      baseRevision: 'base-1',
      manifest: [],
      projectedDocument: textEditedDocument,
      sidecar: projection.sidecar,
      comments: baseComments,
    });
    const [sourceZip, textEditedZip] = await Promise.all([
      JSZip.loadAsync(sourceBuffer),
      JSZip.loadAsync(textEditedExport),
    ]);
    for (const path of [
      'word/comments.xml',
      'word/commentsExtended.xml',
      'word/commentsIds.xml',
      'word/commentsExtensible.xml',
    ]) {
      expect(await textEditedZip.file(path)?.async('text')).toBe(
        await sourceZip.file(path)?.async('text')
      );
    }

    const currentComments = [
      comment(1, 'Edited comment', { done: true }),
      comment(2, 'A reply', { parentId: 1 }),
    ];
    const editedExport = await exportCollaborationDocument({
      sourceBuffer,
      sourceSha256,
      baseRevision: 'base-1',
      manifest: [],
      projectedDocument: projection.document,
      sidecar: projection.sidecar,
      comments: currentComments,
    });
    const reparsed = await parseDocx(editedExport);
    const exportedComments = reparsed.package.document.comments ?? [];
    expect(exportedComments).toHaveLength(2);
    expect(JSON.stringify(exportedComments[0].content)).toContain('Edited comment');
    expect(exportedComments[0].done).toBe(true);
    expect(exportedComments[1].parentId).toBe(1);
    expect(JSON.stringify(exportedComments[1].content)).toContain('A reply');
    const editedZip = await JSZip.loadAsync(editedExport);
    const editedXml = await editedZip.file('word/document.xml')?.async('text');
    expect(editedXml).toContain('<w:commentRangeStart w:id="2"');
    expect(editedXml).toContain('<w:commentRangeEnd w:id="2"');

    const deletedExport = await exportCollaborationDocument({
      sourceBuffer,
      sourceSha256,
      baseRevision: 'base-1',
      manifest: [],
      projectedDocument: projection.document,
      sidecar: projection.sidecar,
      comments: [],
    });
    const deletedZip = await JSZip.loadAsync(deletedExport);
    const deletedDocumentXml = await deletedZip.file('word/document.xml')?.async('text');
    expect(deletedZip.file('word/comments.xml')).toBeNull();
    expect(await deletedZip.file('word/_rels/document.xml.rels')?.async('text')).not.toContain(
      'relationships/comments'
    );
    expect(await deletedZip.file('[Content_Types].xml')?.async('text')).not.toContain(
      '/word/comments'
    );
    expect(deletedDocumentXml).not.toContain('<w:commentRangeStart');
    expect(deletedDocumentXml).not.toContain('<w:commentRangeEnd');
    expect(deletedDocumentXml).not.toContain('<w:commentReference');
  });

  test('fails explicitly instead of replacing changed VML or structural source XML', async () => {
    const sourceBuffer = await createVmlSourceFixture();
    const projection = await projectSource(sourceBuffer);
    const sourceSha256 = await sha256Hex(sourceBuffer);

    const unsafeParagraphEdit = structuredClone(projection.document);
    const firstParagraph = unsafeParagraphEdit.content?.find(
      (node) => node.attrs?.paraId === '11B2C3D4'
    );
    if (!firstParagraph) throw new Error('Expected first paragraph');
    firstParagraph.content = [{ type: 'text', text: 'Unsafe replacement' }];

    const unsafeExport = exportCollaborationDocument({
      sourceBuffer,
      sourceSha256,
      baseRevision: 'base-1',
      manifest: [],
      projectedDocument: unsafeParagraphEdit,
      sidecar: projection.sidecar,
    });
    await expect(unsafeExport).rejects.toBeInstanceOf(CollaborationExportFidelityError);
    await expect(unsafeExport).rejects.toMatchObject({
      reason: 'unsafe_paragraph_markup',
      paragraphId: '11B2C3D4',
    });

    const structuralEdit = structuredClone(projection.document);
    structuralEdit.content?.push({
      type: 'paragraph',
      attrs: { paraId: '31B2C3D4' },
      content: [{ type: 'text', text: 'New paragraph' }],
    });
    const structuralExport = exportCollaborationDocument({
      sourceBuffer,
      sourceSha256,
      baseRevision: 'base-1',
      manifest: [],
      projectedDocument: structuralEdit,
      sidecar: projection.sidecar,
    });
    await expect(structuralExport).rejects.toBeInstanceOf(CollaborationExportFidelityError);
    await expect(structuralExport).rejects.toMatchObject({
      reason: 'structural_change',
    });
  });

  test('rejects a sidecar from a different immutable base revision', async () => {
    const sourceBuffer = await createDocx({
      package: {
        document: {
          content: [{ type: 'paragraph', content: [] }],
        },
      },
    });
    const projected = projectProseMirrorDocument(
      schema.nodes.doc.create({}, [schema.nodes.paragraph.create()]),
      { baseRevision: 'base-1' }
    );

    await expect(
      exportCollaborationDocument({
        sourceBuffer,
        sourceSha256: await sha256Hex(sourceBuffer),
        baseRevision: 'base-2',
        manifest: [],
        projectedDocument: projected.document,
        sidecar: projected.sidecar,
      })
    ).rejects.toBeInstanceOf(CollaborationIdentityMismatchError);
  });

  test('binds a same-revision sidecar to its canonical source projection', async () => {
    const sourceBuffer = await createDocx({
      package: {
        document: {
          content: [
            {
              type: 'paragraph',
              paraId: '51B2C3D4',
              formatting: { spaceBefore: 240 },
              content: [{ type: 'run', content: [{ type: 'text', text: 'First' }] }],
            },
          ],
        },
      },
    });
    const projection = await projectSource(sourceBuffer);
    const sourceSha256 = await sha256Hex(sourceBuffer);
    const sidecar = structuredClone(projection.sidecar);
    const entry = Object.values(sidecar.entries)[0];
    if (!entry) throw new Error('Expected a fidelity sidecar entry');
    entry.attrs = {
      ...entry.attrs,
      _originalFormatting: { spaceBefore: 480 },
    };

    await expect(
      exportCollaborationDocument({
        sourceBuffer,
        sourceSha256,
        baseRevision: 'base-1',
        manifest: [],
        projectedDocument: projection.document,
        sidecar,
      })
    ).rejects.toBeInstanceOf(CollaborationIdentityMismatchError);

    const wrongTypeSidecar = structuredClone(projection.sidecar);
    const wrongTypeEntry = Object.values(wrongTypeSidecar.entries)[0];
    if (!wrongTypeEntry) throw new Error('Expected a fidelity sidecar entry');
    wrongTypeEntry.nodeType = 'table';
    await expect(
      exportCollaborationDocument({
        sourceBuffer,
        sourceSha256,
        baseRevision: 'base-1',
        manifest: [],
        projectedDocument: projection.document,
        sidecar: wrongTypeSidecar,
      })
    ).rejects.toBeInstanceOf(CollaborationIdentityMismatchError);

    const transportedSidecar = JSON.parse(
      JSON.stringify(projection.sidecar)
    ) as typeof projection.sidecar;
    const noOpExport = await exportCollaborationDocument({
      sourceBuffer,
      sourceSha256,
      baseRevision: 'base-1',
      manifest: [],
      projectedDocument: projection.document,
      sidecar: transportedSidecar,
    });
    expect(new Uint8Array(noOpExport)).toEqual(new Uint8Array(sourceBuffer));
  });

  test('rejects duplicate or reassigned collaboration identities', async () => {
    const sourceBuffer = await createDocx({
      package: {
        document: {
          content: [
            {
              type: 'paragraph',
              paraId: '61B2C3D4',
              content: [{ type: 'run', content: [{ type: 'text', text: 'First' }] }],
            },
            {
              type: 'paragraph',
              paraId: '71B2C3D4',
              content: [{ type: 'run', content: [{ type: 'text', text: 'Second' }] }],
            },
          ],
        },
      },
    });
    const projection = await projectSource(sourceBuffer);
    const sourceSha256 = await sha256Hex(sourceBuffer);
    const firstId = projection.document.content?.[0]?.attrs?.collaborationId;
    const secondId = projection.document.content?.[1]?.attrs?.collaborationId;
    if (typeof firstId !== 'string' || typeof secondId !== 'string') {
      throw new Error('Expected source collaboration identities');
    }

    const duplicated = structuredClone(projection.document);
    if (!duplicated.content?.[1]?.attrs) throw new Error('Expected second paragraph attrs');
    duplicated.content[1].attrs.collaborationId = firstId;
    await expect(
      exportCollaborationDocument({
        sourceBuffer,
        sourceSha256,
        baseRevision: 'base-1',
        manifest: [],
        projectedDocument: duplicated,
        sidecar: projection.sidecar,
      })
    ).rejects.toBeInstanceOf(CollaborationIdentityMismatchError);

    const reassigned = structuredClone(projection.document);
    if (!reassigned.content?.[0]?.attrs || !reassigned.content[1]?.attrs) {
      throw new Error('Expected paragraph attrs');
    }
    reassigned.content[0].attrs.collaborationId = secondId;
    reassigned.content[1].attrs.collaborationId = firstId;
    await expect(
      exportCollaborationDocument({
        sourceBuffer,
        sourceSha256,
        baseRevision: 'base-1',
        manifest: [],
        projectedDocument: reassigned,
        sidecar: projection.sidecar,
      })
    ).rejects.toBeInstanceOf(CollaborationIdentityMismatchError);
  });

  test('rejects source bytes that do not match the immutable source digest', async () => {
    const sourceBuffer = await createDocx({
      package: {
        document: {
          content: [{ type: 'paragraph', content: [] }],
        },
      },
    });
    const projected = projectProseMirrorDocument(
      schema.nodes.doc.create({}, [schema.nodes.paragraph.create()]),
      { baseRevision: 'base-1' }
    );

    await expect(
      exportCollaborationDocument({
        sourceBuffer,
        sourceSha256: '0'.repeat(64),
        baseRevision: 'base-1',
        manifest: [],
        projectedDocument: projected.document,
        sidecar: projected.sidecar,
      })
    ).rejects.toBeInstanceOf(CollaborationIdentityMismatchError);
  });

  test('rejects a package whose external-media manifest omits an image', async () => {
    const packageBuffer = await createDocx(
      fromProseDoc(imageDocument('data:image/png;base64,AAAA'))
    );

    await expect(
      parseCollaborationPackage({
        packageBuffer,
        manifest: [],
        baseRevision: 'base-1',
      })
    ).rejects.toBeInstanceOf(IncompleteCollaborationManifestError);
  });

  test('rejects a package whose manifest omits a non-body image relationship', async () => {
    const parsing = parseCollaborationPackage({
      packageBuffer: await createHeaderImageRelationshipFixture(),
      manifest: [],
      baseRevision: 'base-1',
    });

    await expect(parsing).rejects.toBeInstanceOf(IncompleteCollaborationManifestError);
    await expect(parsing).rejects.toMatchObject({
      code: 'incomplete_collaboration_manifest',
      missingPaths: ['word/media/watermark.png'],
    });
  });
});

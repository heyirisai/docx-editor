/**
 * Header/Footer & Comment Packaging
 *
 * Serialize modified headers, footers, and comments back into the ZIP and
 * ensure each new part is registered in `[Content_Types].xml` and the
 * document's rels file. Without this step Word silently drops parts that
 * the editor inserted into a previously-blank document (#274).
 */

import type JSZip from 'jszip';
import type { Document } from '../../types/document';
import type { HeaderFooter } from '../../types/content';
import { serializeHeaderFooter } from '../serializer/headerFooterSerializer';
import {
  serializeCommentsWithInfo,
  serializeCommentsExtended,
  serializeCommentsIds,
  serializeCommentsExtensible,
} from '../serializer/commentSerializer';
import { RELATIONSHIP_TYPES } from '../relsParser';
import { findMaxRId, readRelsOrStub, headerFooterFilename } from './parts';

const HEADER_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml';

const FOOTER_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml';

export const COMMENTS_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml';

export const COMMENTS_EXTENDED_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml';

export const COMMENTS_IDS_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml';

export const COMMENTS_EXTENSIBLE_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml';

/**
 * Ensure every header/footer in `doc.package.relationships` is wired up in
 * `[Content_Types].xml` and `word/_rels/document.xml.rels`. For blank documents
 * where the user adds a header/footer for the first time, these files don't
 * know about the new part yet — without this Word silently drops them (#274).
 */
export async function ensureHeaderFooterParts(
  doc: Document,
  zip: JSZip,
  compressionLevel: number
): Promise<void> {
  const rels = doc.package.relationships;
  if (!rels) return;

  const parts: Array<{ rId: string; target: string; contentType: string; relType: string }> = [];
  for (const [rId, rel] of rels) {
    if (!rel.target) continue;
    const contentType =
      rel.type === RELATIONSHIP_TYPES.header
        ? HEADER_CONTENT_TYPE
        : rel.type === RELATIONSHIP_TYPES.footer
          ? FOOTER_CONTENT_TYPE
          : null;
    if (!contentType) continue;
    parts.push({
      rId,
      target: rel.target.replace(/^(\/?word\/)/, ''),
      contentType,
      relType: rel.type,
    });
  }
  if (parts.length === 0) return;

  const ctFile = zip.file('[Content_Types].xml');
  if (ctFile) {
    let ctXml = await ctFile.async('text');
    let changed = false;
    for (const { target, contentType } of parts) {
      const partName = `/word/${target}`;
      if (!ctXml.includes(`PartName="${partName}"`)) {
        ctXml = ctXml.replace(
          '</Types>',
          `<Override PartName="${partName}" ContentType="${contentType}"/></Types>`
        );
        changed = true;
      }
    }
    if (changed) {
      zip.file('[Content_Types].xml', ctXml, {
        compression: 'DEFLATE',
        compressionOptions: { level: compressionLevel },
      });
    }
  }

  const relsPath = 'word/_rels/document.xml.rels';
  let relsXml = await readRelsOrStub(zip, relsPath);
  let relsChanged = false;
  for (const { rId, relType, target } of parts) {
    if (!relsXml.includes(`Id="${rId}"`)) {
      relsXml = relsXml.replace(
        '</Relationships>',
        `<Relationship Id="${rId}" Type="${relType}" Target="${target}"/></Relationships>`
      );
      relsChanged = true;
    }
  }
  if (relsChanged) {
    zip.file(relsPath, relsXml, {
      compression: 'DEFLATE',
      compressionOptions: { level: compressionLevel },
    });
  }
}

/**
 * Ensure content types and relationships exist for all comment parts.
 * Reads each shared file once, applies all modifications, writes once.
 */
async function ensureAllCommentParts(zip: JSZip, compressionLevel: number): Promise<void> {
  const COMMENT_PARTS = [
    {
      partName: '/word/comments.xml',
      contentType: COMMENTS_CONTENT_TYPE,
      target: 'comments.xml',
      relType: RELATIONSHIP_TYPES.comments,
    },
    {
      partName: '/word/commentsExtended.xml',
      contentType: COMMENTS_EXTENDED_CONTENT_TYPE,
      target: 'commentsExtended.xml',
      relType: RELATIONSHIP_TYPES.commentsExtended,
    },
    {
      partName: '/word/commentsIds.xml',
      contentType: COMMENTS_IDS_CONTENT_TYPE,
      target: 'commentsIds.xml',
      relType: RELATIONSHIP_TYPES.commentsIds,
    },
    {
      partName: '/word/commentsExtensible.xml',
      contentType: COMMENTS_EXTENSIBLE_CONTENT_TYPE,
      target: 'commentsExtensible.xml',
      relType: RELATIONSHIP_TYPES.commentsExtensible,
    },
  ];

  // Content types — single read/write
  const ctFile = zip.file('[Content_Types].xml');
  if (ctFile) {
    let ctXml = await ctFile.async('text');
    let changed = false;
    for (const { partName, contentType } of COMMENT_PARTS) {
      if (!ctXml.includes(partName)) {
        ctXml = ctXml.replace(
          '</Types>',
          `<Override PartName="${partName}" ContentType="${contentType}"/></Types>`
        );
        changed = true;
      }
    }
    if (changed) {
      zip.file('[Content_Types].xml', ctXml, {
        compression: 'DEFLATE',
        compressionOptions: { level: compressionLevel },
      });
    }
  }

  // Relationships — single read/write
  const relsPath = 'word/_rels/document.xml.rels';
  const relsFile = zip.file(relsPath);
  if (relsFile) {
    let relsXml = await relsFile.async('text');
    let changed = false;
    for (const { target, relType } of COMMENT_PARTS) {
      if (!relsXml.includes(target)) {
        const newRId = `rId${findMaxRId(relsXml) + 1}`;
        relsXml = relsXml.replace(
          '</Relationships>',
          `<Relationship Id="${newRId}" Type="${relType}" Target="${target}"/></Relationships>`
        );
        changed = true;
      }
    }
    if (changed) {
      zip.file(relsPath, relsXml, {
        compression: 'DEFLATE',
        compressionOptions: { level: compressionLevel },
      });
    }
  }
}

/**
 * Serialize all comment-side XML parts (comments + extended + ids + extensible)
 * into the ZIP and register them in content types and rels.
 */
export async function serializeCommentsToZip(
  doc: Document,
  zip: JSZip,
  compressionLevel: number
): Promise<void> {
  const comments = doc.package.document.comments;
  if (!comments || comments.length === 0) return;

  const { xml: commentsXml, paraInfos } = serializeCommentsWithInfo(comments);
  zip.file('word/comments.xml', commentsXml, {
    compression: 'DEFLATE',
    compressionOptions: { level: compressionLevel },
  });

  // Write commentsExtended.xml for reply threading (Word/Google Docs interop)
  const extendedXml = serializeCommentsExtended(paraInfos);
  if (extendedXml) {
    zip.file('word/commentsExtended.xml', extendedXml, {
      compression: 'DEFLATE',
      compressionOptions: { level: compressionLevel },
    });
  }

  // Write commentsIds.xml for stable IDs (Word Online needs this for replies)
  const idsXml = serializeCommentsIds(paraInfos);
  if (idsXml) {
    zip.file('word/commentsIds.xml', idsXml, {
      compression: 'DEFLATE',
      compressionOptions: { level: compressionLevel },
    });
  }

  // Write commentsExtensible.xml for UTC dates (Pages, Word 2016+)
  const extensibleXml = serializeCommentsExtensible(paraInfos, comments);
  if (extensibleXml) {
    zip.file('word/commentsExtensible.xml', extensibleXml, {
      compression: 'DEFLATE',
      compressionOptions: { level: compressionLevel },
    });
  }

  await ensureAllCommentParts(zip, compressionLevel);
}

/**
 * Collect serialized header/footer XML updates from the document model.
 * Uses the relationship map to resolve rId → filename.
 */
export function collectHeaderFooterUpdates(doc: Document): Map<string, string> {
  const updates = new Map<string, string>();
  const rels = doc.package.relationships;
  if (!rels) return updates;

  const parts: Array<{
    map: Map<string, HeaderFooter> | undefined;
    type: string;
  }> = [
    { map: doc.package.headers, type: RELATIONSHIP_TYPES.header },
    { map: doc.package.footers, type: RELATIONSHIP_TYPES.footer },
  ];

  for (const { map, type } of parts) {
    if (!map) continue;
    for (const [rId, headerFooter] of map.entries()) {
      const rel = rels.get(rId);
      if (rel && rel.type === type && rel.target) {
        updates.set(headerFooterFilename(rel.target), serializeHeaderFooter(headerFooter));
      }
    }
  }

  return updates;
}

/**
 * Serialize modified headers and footers into the ZIP.
 */
export function serializeHeadersFootersToZip(
  doc: Document,
  zip: JSZip,
  compressionLevel: number
): void {
  const compressionOptions = { level: compressionLevel };
  for (const [filename, xml] of collectHeaderFooterUpdates(doc)) {
    zip.file(filename, xml, { compression: 'DEFLATE', compressionOptions });
  }
}

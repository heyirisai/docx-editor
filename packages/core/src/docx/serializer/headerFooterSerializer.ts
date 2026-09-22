/**
 * Header/Footer Serializer - Serialize headers/footers to OOXML XML
 *
 * Converts HeaderFooter objects back to valid header*.xml / footer*.xml format.
 * Reuses paragraph and table serializers for content.
 *
 * OOXML Reference:
 * - Header root: w:hdr
 * - Footer root: w:ftr
 * - Content: w:p, w:tbl (same as document body)
 */

import type { BlockContent, HeaderFooter } from '../../types/document';
import { buildMcIgnorable } from './xmlUtils';
import { CAPTURING_ROOT_PREFIXES, buildRootNamespaces } from './rootNamespaces';
import { serializeParagraph } from './paragraphSerializer';
import { serializeTable } from './tableSerializer';
import { serializeBlockSdt } from './sdtSerializer';
import { serializeWatermark } from './vmlWatermarkSerializer';

/**
 * Serialize a block content item (paragraph, table, or block SDT) for
 * header/footer content.
 */
function serializeBlock(block: BlockContent): string {
  if (block.type === 'paragraph') {
    return serializeParagraph(block);
  } else if (block.type === 'table') {
    return serializeTable(block);
  } else if (block.type === 'blockSdt') {
    return serializeBlockSdt(block, serializeBlock);
  }
  return '';
}

/**
 * Serialize a HeaderFooter object to valid OOXML XML
 *
 * @param hf - HeaderFooter object to serialize
 * @returns Complete XML string for header*.xml or footer*.xml
 */
export function serializeHeaderFooter(hf: HeaderFooter): string {
  const rootTag = hf.type === 'header' ? 'w:hdr' : 'w:ftr';
  const ns = buildRootNamespaces(CAPTURING_ROOT_PREFIXES, hf.rootNamespaces);
  // A header is exactly where this PR's preserved `mc:AlternateContent` lands,
  // and an extension element is only skippable if its prefix is listed here.
  const ignorable = buildMcIgnorable(ns.prefixes, hf.rootIgnorable);

  // Serialize content blocks
  let contentXml = hf.content.map((block) => serializeBlock(block)).join('');

  // Prepend the watermark VML (Word stores it as the first run in the header)
  // so it paints behind the body content.
  if (hf.watermark) {
    contentXml = serializeWatermark(hf.watermark) + contentXml;
  }

  // Ensure at least one empty paragraph (required by OOXML spec)
  if (!contentXml) {
    contentXml = '<w:p><w:pPr/></w:p>';
  }

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<${rootTag} ${ns.decl}${ignorable}>${contentXml}</${rootTag}>`;
}

/**
 * New Hyperlink Registration
 *
 * On save, scan all parts for hyperlinks that have an href but no rId
 * (created in the editor), assign rIds, and write the relationship entries
 * to the owning part's rels file.
 */

import type JSZip from 'jszip';
import type { BlockContent, Hyperlink } from '../../types/content';
import { RELATIONSHIP_TYPES } from '../relsParser';
import { escapeXml } from '../serializer/xmlUtils';
import { findMaxRId, readRelsOrStub, type Part } from './parts';

/**
 * Collect all hyperlinks that have an href but no rId from block content.
 * These are newly created hyperlinks that need relationship entries.
 */
function collectHyperlinksWithoutRId(blocks: BlockContent[]): Hyperlink[] {
  const hyperlinks: Hyperlink[] = [];

  for (const block of blocks) {
    if (block.type === 'paragraph') {
      for (const item of block.content) {
        if (item.type === 'hyperlink' && item.href && !item.rId && !item.anchor) {
          hyperlinks.push(item);
        }
      }
    } else if (block.type === 'table') {
      for (const row of block.rows) {
        for (const cell of row.cells) {
          hyperlinks.push(...collectHyperlinksWithoutRId(cell.content));
        }
      }
    }
  }

  return hyperlinks;
}

/**
 * Process newly created hyperlinks across all parts (body, headers, footers):
 * assign rIds and add relationship entries to the owning part's rels file.
 *
 * Mutates each hyperlink's rId in-place.
 */
export async function processNewHyperlinks(
  parts: Part[],
  zip: JSZip,
  compressionLevel: number
): Promise<void> {
  for (const { relsPath, blocks } of parts) {
    const hyperlinks = collectHyperlinksWithoutRId(blocks);
    if (hyperlinks.length === 0) continue;

    const relsXml = await readRelsOrStub(zip, relsPath);
    let maxId = findMaxRId(relsXml);
    const relEntries: string[] = [];

    for (const hyperlink of hyperlinks) {
      maxId++;
      const newRId = `rId${maxId}`;

      relEntries.push(
        `<Relationship Id="${newRId}" Type="${RELATIONSHIP_TYPES.hyperlink}" Target="${escapeXml(hyperlink.href!)}" TargetMode="External"/>`
      );

      hyperlink.rId = newRId;
    }

    const updatedRelsXml = relsXml.replace(
      '</Relationships>',
      relEntries.join('') + '</Relationships>'
    );
    zip.file(relsPath, updatedRelsXml, {
      compression: 'DEFLATE',
      compressionOptions: { level: compressionLevel },
    });
  }
}

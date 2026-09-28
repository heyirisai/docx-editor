/**
 * Paragraph-id uniqueness.
 *
 * `w14:paraId` is a per-document paragraph identity used by Word for
 * co-authoring and revision tracking. Some foreign exporters emit the same
 * value on multiple paragraphs, which Word flags. Individual values are already
 * range-normalized at parse (see {@link normalizeLongHexId}); this pass runs
 * afterwards and regenerates any *duplicate* paraId across the main document
 * stories so every paragraph carries a unique id.
 *
 * Comment-content paragraphs are intentionally excluded — their paraIds are
 * managed against comments.xml / commentsExtended.xml on save.
 */

import type { BlockContent, Document, Paragraph } from '../types/document';
import { generateHexId, isValidLongHexId } from '../utils/hexId';

/**
 * A fresh `w14:paraId` not in `seen` (valid `ST_LongHexNumber`, below
 * `0x80000000`). The caller records it in `seen` if later ids must avoid it.
 */
export function uniqueParagraphId(seen: ReadonlySet<string>): string {
  let id = generateHexId();
  while (seen.has(id) || !isValidLongHexId(id)) {
    id = generateHexId();
  }
  return id;
}

function dedupeParagraph(paragraph: Paragraph, seen: Set<string>): void {
  if (paragraph.paraId == null) return;
  if (seen.has(paragraph.paraId)) {
    paragraph.paraId = uniqueParagraphId(seen);
  }
  seen.add(paragraph.paraId);
}

function walkBlocks(
  blocks: BlockContent[] | undefined,
  seen: Set<string>,
  visit: (paragraph: Paragraph, seen: Set<string>) => void
): void {
  if (!blocks) return;

  for (const block of blocks) {
    if (block.type === 'paragraph') {
      visit(block, seen);
    } else if (block.type === 'table') {
      for (const row of block.rows) {
        for (const cell of row.cells) {
          walkBlocks(cell.content, seen, visit);
        }
      }
    } else if (block.type === 'blockSdt') {
      walkBlocks(block.content, seen, visit);
    }
  }
}

/**
 * Ensure every paragraph in the main document stories (body, headers, footers,
 * footnotes, endnotes, separators) has a unique `paraId`, regenerating any
 * later duplicate.
 */
export function dedupeParagraphIds(doc: Document): void {
  walkStories(doc, new Set<string>(), dedupeParagraph);
}

/**
 * Every `w14:paraId` in the main document stories (the same scope
 * {@link dedupeParagraphIds} keeps unique), for minting a fresh id that
 * collides with none of them (see {@link uniqueParagraphId}).
 */
export function collectParagraphIds(doc: Document): Set<string> {
  const seen = new Set<string>();
  walkStories(doc, seen, (paragraph, ids) => {
    if (paragraph.paraId != null) ids.add(paragraph.paraId);
  });
  return seen;
}

function walkStories(
  doc: Document,
  seen: Set<string>,
  visit: (paragraph: Paragraph, seen: Set<string>) => void
): void {
  const pkg = doc.package;
  walkBlocks(pkg.document?.content, seen, visit);

  for (const header of pkg.headers?.values() ?? []) walkBlocks(header.content, seen, visit);
  for (const footer of pkg.footers?.values() ?? []) walkBlocks(footer.content, seen, visit);
  for (const note of pkg.footnotes ?? []) walkBlocks(note.content, seen, visit);
  for (const note of pkg.endnotes ?? []) walkBlocks(note.content, seen, visit);
  for (const note of pkg.footnoteSeparators ?? []) walkBlocks(note.content, seen, visit);
  for (const note of pkg.endnoteSeparators ?? []) walkBlocks(note.content, seen, visit);
}

/**
 * Paragraph identity across a content-control content rewrite.
 *
 * Answering a control replaces its paragraphs, but the paragraphs themselves
 * are the same ones Word had: Word keeps a paragraph's `w14:paraId` (and its
 * `w:pPr`) when the text in it is edited. Hosts key answer anchors by that
 * paraId, so a rewrite that drops it makes the slot unfindable after a save.
 *
 * Contract, shared by the headless and editor (PM) paths:
 * - the replacement's Nth paragraph takes over the control's Nth original
 *   paragraph — paraId, textId, paragraph properties — with the new content;
 * - a paragraph beyond the original count continues the last original one's
 *   paragraph properties (as pressing Enter in Word does) under a fresh paraId
 *   unique in the document (and a fresh textId when the source had one);
 * - original paragraphs beyond the replacement's count go away with their ids.
 */

import type { BlockContent, Document, Paragraph } from '../types/document';
import { collectParagraphIds, uniqueParagraphId } from '../docx/paragraphIdIntegrity';
import { generateHexId } from '../utils/hexId';

/** Mints paraIds unique in one document, across all the writes of one mutation. */
export interface ParaIdPool {
  fresh(): string;
}

/** A pool over the ids `collect()` returns, collected lazily on the first mint. */
export function createParaIdPool(collect: () => Set<string>): ParaIdPool {
  let seen: Set<string> | undefined;
  return {
    fresh() {
      seen ??= collect();
      const id = uniqueParagraphId(seen);
      seen.add(id);
      return id;
    },
  };
}

/** A pool over every paraId in `doc`'s main stories. */
export function documentParaIdPool(doc: Document): ParaIdPool {
  return createParaIdPool(() => collectParagraphIds(doc));
}

/** A paragraph's identity + properties, minus content-bound state. */
function carriedFrom(source: Paragraph): Omit<Paragraph, 'content'> {
  // `renderedPageBreakBefore` is Word's cached layout of the old content.
  const { content: _content, renderedPageBreakBefore: _rendered, ...rest } = source;
  return structuredClone(rest);
}

/** Properties a following paragraph continues from `source`, with a new identity. */
function continuedFrom(source: Paragraph, pool: ParaIdPool): Omit<Paragraph, 'content'> {
  const next: Omit<Paragraph, 'content'> = { type: 'paragraph', paraId: pool.fresh() };
  if (source.textId) next.textId = generateHexId();
  if (source.formatting) next.formatting = structuredClone(source.formatting);
  if (source.listRendering) next.listRendering = structuredClone(source.listRendering);
  return next;
}

/**
 * Apply the identity contract to `replacement`, the new content of a control
 * whose content was `original`. Only the direct paragraphs of each list take
 * part. With `explicit` (caller-built blocks) a replacement paragraph keeps its
 * own properties and only gains the missing paraId/textId.
 */
export function carryParagraphIdentity(
  original: readonly BlockContent[],
  replacement: BlockContent[],
  pool: ParaIdPool,
  options: { explicit?: boolean } = {}
): BlockContent[] {
  const sources = original.filter((b): b is Paragraph => b.type === 'paragraph');
  let index = 0;
  return replacement.map((block) => {
    if (block.type !== 'paragraph') return block;
    const source = sources[index] as Paragraph | undefined;
    const last = sources[sources.length - 1];
    index += 1;
    if (options.explicit) {
      if (block.paraId) return block;
      const next: Paragraph = { ...block, paraId: source?.paraId ?? pool.fresh() };
      if (!block.textId && source?.textId) next.textId = source.textId;
      return next;
    }
    if (source) return { ...carriedFrom(source), content: block.content };
    if (last) return { ...continuedFrom(last, pool), content: block.content };
    return { ...block, paraId: pool.fresh() };
  });
}

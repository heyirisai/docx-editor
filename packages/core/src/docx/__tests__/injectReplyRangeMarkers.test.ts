/**
 * Reply range-marker injection must reach comment anchors ANYWHERE in the
 * body — including inside block-level SDTs (content controls). Real
 * proposals anchor comments inside a TOC `buildingBlockGallery` SDT; a
 * reply without markers in document.xml is silently DISCARDED by Word on
 * open, so a skipped container means lost replies (found via a live-collab
 * export where every reply of an SDT-anchored thread vanished in Word).
 */

import { describe, expect, test } from 'bun:test';
import { injectReplyRangeMarkers, injectTCReplyRangeMarkers } from '../injectReplyRangeMarkers';
import type { BlockContent, Comment, ParagraphContent } from '../../types/content';

function comment(id: number, parentId?: number): Comment {
  return { id, parentId, author: 'x', date: '2026-07-08T00:00:00Z', content: [] };
}

function paragraphWithRange(commentId: number): BlockContent {
  return {
    type: 'paragraph',
    content: [
      { type: 'commentRangeStart', id: commentId },
      { type: 'run', content: [{ type: 'text', text: 'anchor' }] },
      { type: 'commentRangeEnd', id: commentId },
    ] as ParagraphContent[],
  } as BlockContent;
}

function markerIds(block: BlockContent, kind: 'commentRangeStart' | 'commentRangeEnd'): number[] {
  const para = block as { content: ParagraphContent[] };
  return para.content
    .filter((i): i is { type: typeof kind; id: number } => i.type === kind)
    .map((i) => i.id);
}

describe('injectReplyRangeMarkers', () => {
  test('injects reply markers next to a parent in a plain paragraph', () => {
    const para = paragraphWithRange(7);
    injectReplyRangeMarkers([para], [comment(7), comment(8, 7), comment(9, 7)]);
    expect(markerIds(para, 'commentRangeStart')).toEqual([7, 8, 9]);
    expect(markerIds(para, 'commentRangeEnd')).toEqual([7, 8, 9]);
  });

  test('injects reply markers for a parent anchored inside a block SDT (TOC gallery)', () => {
    const para = paragraphWithRange(7);
    const sdt = { type: 'blockSdt', properties: {}, content: [para] } as unknown as BlockContent;
    injectReplyRangeMarkers([sdt], [comment(7), comment(8, 7)]);
    expect(markerIds(para, 'commentRangeStart')).toEqual([7, 8]);
    expect(markerIds(para, 'commentRangeEnd')).toEqual([7, 8]);
  });

  test('injects reply markers for a parent inside an SDT nested in a table cell', () => {
    const para = paragraphWithRange(7);
    const sdt = { type: 'blockSdt', properties: {}, content: [para] } as unknown as BlockContent;
    const table = {
      type: 'table',
      rows: [{ cells: [{ content: [sdt] }] }],
    } as unknown as BlockContent;
    injectReplyRangeMarkers([table], [comment(7), comment(8, 7)]);
    expect(markerIds(para, 'commentRangeStart')).toEqual([7, 8]);
  });

  test('injects reply markers through nested block SDTs', () => {
    const para = paragraphWithRange(7);
    const inner = { type: 'blockSdt', properties: {}, content: [para] } as unknown as BlockContent;
    const outer = { type: 'blockSdt', properties: {}, content: [inner] } as unknown as BlockContent;
    injectReplyRangeMarkers([outer], [comment(7), comment(8, 7)]);
    expect(markerIds(para, 'commentRangeStart')).toEqual([7, 8]);
  });
});

describe('injectTCReplyRangeMarkers', () => {
  test('wraps a tracked change inside a block SDT with reply markers', () => {
    const para = {
      type: 'paragraph',
      content: [
        {
          type: 'insertion',
          info: { id: 41, author: 'x', date: '2026-07-08T00:00:00Z' },
          content: [],
        },
      ],
    } as unknown as BlockContent;
    const sdt = { type: 'blockSdt', properties: {}, content: [para] } as unknown as BlockContent;
    // Reply 50 whose parent (41) is a revision id, not a comment id.
    injectTCReplyRangeMarkers([sdt], [comment(50, 41)]);
    expect(markerIds(para, 'commentRangeStart')).toEqual([50]);
    expect(markerIds(para, 'commentRangeEnd')).toEqual([50]);
  });
});

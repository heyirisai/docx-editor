import { describe, test, expect } from 'bun:test';
import type { Comment } from '../../types/content';
import {
  serializeComments,
  serializeCommentsWithInfo,
  serializeCommentsExtended,
  serializeCommentsIds,
  serializeCommentsExtensible,
} from './commentSerializer';

function makeComment(overrides: Partial<Comment> = {}): Comment {
  return {
    id: 1,
    author: 'Alice',
    date: '2024-01-01T00:00:00Z',
    content: [
      { type: 'paragraph', content: [{ type: 'run', content: [{ type: 'text', text: 'Test' }] }] },
    ],
    ...overrides,
  };
}

function makeMultiParagraphComment(overrides: Partial<Comment> = {}): Comment {
  return {
    id: 1,
    author: 'Alice',
    date: '2024-01-01T00:00:00Z',
    content: [
      { type: 'paragraph', content: [{ type: 'run', content: [{ type: 'text', text: 'First' }] }] },
      {
        type: 'paragraph',
        content: [{ type: 'run', content: [{ type: 'text', text: 'Second' }] }],
      },
      { type: 'paragraph', content: [{ type: 'run', content: [{ type: 'text', text: 'Third' }] }] },
    ],
    ...overrides,
  };
}

describe('commentSerializer', () => {
  describe('w:done (resolved state)', () => {
    test('does NOT put w:done on w:comment (resolved state is in commentsExtended only)', () => {
      const xml = serializeComments([makeComment({ done: true })]);
      expect(xml).not.toContain('w:done');
    });

    test('resolved state appears in commentsExtended.xml as w15:done="1"', () => {
      const { paraInfos } = serializeCommentsWithInfo([makeComment({ done: true })]);
      const extXml = serializeCommentsExtended(paraInfos);
      expect(extXml).toContain('w15:done="1"');
    });

    test('unresolved state appears in commentsExtended.xml as w15:done="0"', () => {
      const { paraInfos } = serializeCommentsWithInfo([makeComment({ done: false })]);
      const extXml = serializeCommentsExtended(paraInfos);
      expect(extXml).toContain('w15:done="0"');
    });
  });

  describe('paraId placement (last paragraph only)', () => {
    test('single-paragraph: w14:paraId appears exactly once', () => {
      const xml = serializeComments([makeComment()]);
      const matches = xml.match(/w14:paraId=/g);
      expect(matches).toHaveLength(1);
    });

    test('single-paragraph: no w14:textId', () => {
      const xml = serializeComments([makeComment()]);
      expect(xml).not.toContain('w14:textId');
    });

    test('multi-paragraph: w14:paraId only on LAST w:p', () => {
      const xml = serializeComments([makeMultiParagraphComment()]);
      const paraIdMatches = xml.match(/w14:paraId=/g);
      expect(paraIdMatches).toHaveLength(1);

      // paraId should be on the paragraph containing "Third" (last)
      const paraIdPos = xml.indexOf('w14:paraId=');
      const thirdPos = xml.indexOf('Third');
      const firstPos = xml.indexOf('First');
      expect(paraIdPos).toBeGreaterThan(firstPos);
      // paraId is on the <w:p> tag, which comes before the text content
      expect(paraIdPos).toBeLessThan(thirdPos);
    });

    test('multi-paragraph: no w14:textId anywhere', () => {
      const xml = serializeComments([makeMultiParagraphComment()]);
      expect(xml).not.toContain('w14:textId');
    });

    test('empty comment: w14:paraId present, no w14:textId', () => {
      const xml = serializeComments([makeComment({ content: [] })]);
      const paraIdMatches = xml.match(/w14:paraId=/g);
      expect(paraIdMatches).toHaveLength(1);
      expect(xml).not.toContain('w14:textId');
    });
  });

  describe('durableId independence', () => {
    test('durableId differs from paraId for every comment', () => {
      const { paraInfos } = serializeCommentsWithInfo([
        makeComment({ id: 1 }),
        makeComment({ id: 2 }),
      ]);
      for (const info of paraInfos) {
        expect(info.durableId).not.toBe(info.lastParaId);
      }
    });

    test('durableId is 8-char hex', () => {
      const { paraInfos } = serializeCommentsWithInfo([makeComment()]);
      expect(paraInfos[0].durableId).toMatch(/^[0-9A-F]{8}$/);
    });
  });

  describe('cross-file consistency', () => {
    test('paraId matches across comments.xml, commentsExtended, and commentsIds', () => {
      const comments = [makeComment({ id: 1 }), makeComment({ id: 2, parentId: 1 })];
      const { xml, paraInfos } = serializeCommentsWithInfo(comments);
      const extXml = serializeCommentsExtended(paraInfos);
      const idsXml = serializeCommentsIds(paraInfos);

      for (const info of paraInfos) {
        // paraId should appear in comments.xml as w14:paraId
        expect(xml).toContain(`w14:paraId="${info.lastParaId}"`);
        // Same paraId in commentsExtended.xml as w15:paraId
        expect(extXml).toContain(`w15:paraId="${info.lastParaId}"`);
        // Same paraId in commentsIds.xml as w16cid:paraId
        expect(idsXml).toContain(`w16cid:paraId="${info.lastParaId}"`);
      }
    });

    test('reply threading: parent paraId appears as paraIdParent', () => {
      const comments = [makeComment({ id: 1 }), makeComment({ id: 2, parentId: 1 })];
      const { paraInfos } = serializeCommentsWithInfo(comments);
      const extXml = serializeCommentsExtended(paraInfos);

      const parentInfo = paraInfos.find((p) => p.commentId === 1)!;
      expect(extXml).toContain(`w15:paraIdParent="${parentInfo.lastParaId}"`);
    });

    test('durableId matches between commentsIds and commentsExtensible', () => {
      const comments = [makeComment({ id: 1 })];
      const { paraInfos } = serializeCommentsWithInfo(comments);
      const idsXml = serializeCommentsIds(paraInfos);
      const extXml = serializeCommentsExtensible(paraInfos, comments);

      expect(idsXml).toContain(`w16cid:durableId="${paraInfos[0].durableId}"`);
      expect(extXml).toContain(`w16cex:durableId="${paraInfos[0].durableId}"`);
    });
  });

  describe('reply threading', () => {
    test('does NOT put w16cid:parentId on w:comment', () => {
      const xml = serializeComments([makeComment({ id: 1 }), makeComment({ id: 2, parentId: 1 })]);
      expect(xml).not.toContain('w16cid:parentId');
    });

    test('serializes top-level comments before replies', () => {
      const xml = serializeComments([makeComment({ id: 2, parentId: 1 }), makeComment({ id: 1 })]);
      const topPos = xml.indexOf('w:id="1"');
      const replyPos = xml.indexOf('w:id="2"');
      expect(topPos).toBeLessThan(replyPos);
    });

    test('commentsExtended.xml links reply to parent via paraIdParent', () => {
      const { paraInfos } = serializeCommentsWithInfo([
        makeComment({ id: 1 }),
        makeComment({ id: 2, parentId: 1 }),
      ]);
      const extXml = serializeCommentsExtended(paraInfos);
      expect(extXml).toContain('w15:paraIdParent=');
    });
  });

  describe('collab namespaced IDs (issue #257)', () => {
    // In live collaboration each client mints comment IDs inside a private
    // 2^21-wide block keyed by its Yjs clientID, so IDs can be up to 2^53−1.
    // A parent from one client and a reply from another must both serialize,
    // with threading intact and the large IDs emitted verbatim.
    const parentId = 0x9abcdef1 * 2 ** 21 + 1; // client A's block
    const replyId = 0x12345678 * 2 ** 21 + 1; // client B's block

    test('parent and reply with cross-client IDs both serialize', () => {
      const parent = makeComment({ id: parentId });
      const reply = makeComment({ id: replyId, author: 'Bob', parentId: parentId });
      const { xml } = serializeCommentsWithInfo([parent, reply]);
      expect(xml.match(/<w:comment /g)?.length).toBe(2);
      expect(xml).toContain(`w:id="${parentId}"`);
      expect(xml).toContain(`w:id="${replyId}"`);
    });

    test('reply threading survives cross-client IDs', () => {
      const parent = makeComment({ id: parentId });
      const reply = makeComment({ id: replyId, author: 'Bob', parentId: parentId });
      const { paraInfos } = serializeCommentsWithInfo([parent, reply]);
      const extended = serializeCommentsExtended(paraInfos);
      const parentInfo = paraInfos.find((p) => p.commentId === parentId);
      expect(extended).toContain(`w15:paraIdParent="${parentInfo?.lastParaId}"`);
    });

    test('the largest possible namespaced ID round-trips through serialization', () => {
      const maxId = Number.MAX_SAFE_INTEGER; // 2^53 − 1
      const { xml } = serializeCommentsWithInfo([makeComment({ id: maxId })]);
      expect(xml).toContain(`w:id="${maxId}"`);
      expect(String(maxId)).toBe('9007199254740991'); // no precision loss
    });
  });

  describe('date formatting', () => {
    test('strips milliseconds from w:date', () => {
      const xml = serializeComments([makeComment({ date: '2024-01-01T12:30:45.123Z' })]);
      expect(xml).toContain('w:date="2024-01-01T12:30:45Z"');
      expect(xml).not.toContain('.123');
    });
  });
});

/**
 * Comment + tracked-change ID allocation.
 *
 * Comments (`w:comment` ids) and tracked changes (`w:ins`/`w:del` revision ids)
 * share a single OOXML ID space — a duplicate ID between the two corrupts the
 * round-trip. Allocation is therefore one monotonic, no-reuse counter, exposed
 * as an **instance-scoped** factory rather than module-global state so two
 * editor instances on one page never share (or collide on) a counter.
 *
 * Kept separate from the comment/tracked-change transaction builders
 * (`commentOps.ts`) so the allocator can be owned independently — the editor
 * engine seeds and threads it without dragging in the PM-text-lookup graph.
 */

import type { EditorView } from 'prosemirror-view';
import type { Comment } from '../types/content';

/** Sentinel ID for a comment that hasn't been persisted yet (anchored to selection). */
export const PENDING_COMMENT_ID = -1;

export interface CommentIdAllocator {
  /** Allocate the next ID and advance the counter. */
  next(): number;
  /**
   * Bump the counter above an ID observed in the document (loaded comments,
   * tracked-change marks, or comments arriving from a collaboration peer) so
   * subsequent allocations don't collide with it. Namespaced allocators
   * ignore IDs outside their own block — those can never collide with the
   * IDs this allocator mints.
   */
  seedAbove(observedId: number): void;
}

/**
 * Size of each namespace block: a namespaced allocator mints IDs in
 * `[namespace * STRIDE + 1, (namespace + 1) * STRIDE)`. 2^21 IDs per
 * namespace is far beyond any real session; with a 32-bit namespace the
 * largest possible ID is 2^53 − 1, which is still a safe JS integer.
 * (Large `w:id` values are already in the wild: tracked-change revision
 * ids are minted from `Date.now()` — see plugins/revisionIds.ts.)
 */
export const COMMENT_ID_NAMESPACE_STRIDE = 2 ** 21;

/**
 * Create an instance-scoped monotonic comment/revision ID allocator. IDs are
 * never reused (deleting a comment does not free its ID), and the counter is
 * private to this allocator — multiple editors get independent ID spaces.
 *
 * `namespace` makes the allocator collaboration-safe: pass a per-client
 * unique 32-bit integer (e.g. the Yjs `doc.clientID`) and every mint lands
 * in that client's private block, so two peers editing the same document can
 * never mint the same ID even before they sync (issue #257). Without a
 * namespace the allocator behaves exactly as before (counter starts at 1,
 * seeds above every observed ID).
 */
export function createCommentIdAllocator(namespace?: number): CommentIdAllocator {
  // >>> 0 coerces to uint32 so a negative/fractional namespace can't produce
  // IDs that collide with another client's block or with legacy small IDs.
  const base = namespace != null ? (namespace >>> 0) * COMMENT_ID_NAMESPACE_STRIDE : 0;
  let nextId = base + 1;
  return {
    next: () => nextId++,
    seedAbove(observedId: number) {
      // A namespaced allocator only ever mints inside its own block, so only
      // in-block IDs can collide with future mints. Seeding above a foreign
      // block's (much larger) IDs would spill the counter out of our block.
      if (
        namespace != null &&
        (observedId < base || observedId >= base + COMMENT_ID_NAMESPACE_STRIDE)
      ) {
        return;
      }
      if (observedId >= nextId) nextId = observedId + 1;
    },
  };
}

/**
 * Seed an allocator above every comment/revision ID currently in the document
 * — comment objects (including replies, which carry no mark) plus
 * tracked-change `revisionId` marks. Because `seedAbove` only ever raises the
 * counter, this is safe to call on load (React), before each allocation
 * (Vue), or whenever remote comments arrive (collab): new IDs never collide
 * with or reuse an existing one, and the comment and revision ID spaces stay
 * unified. Each ID is fed to `seedAbove` individually (not a global max) so
 * a namespaced allocator can filter foreign-block IDs per observation.
 */
export function seedCommentAllocator(
  allocator: CommentIdAllocator,
  comments: Comment[] | undefined,
  view: EditorView | null
): void {
  for (const comment of comments ?? []) allocator.seedAbove(comment.id);
  if (view) {
    view.state.doc.descendants((node) => {
      for (const mark of node.marks) {
        if (mark.attrs.revisionId != null) allocator.seedAbove(mark.attrs.revisionId as number);
      }
    });
  }
}

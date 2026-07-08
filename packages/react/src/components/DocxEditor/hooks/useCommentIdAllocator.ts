import { useEffect, useRef } from 'react';
import type { Comment } from '@eigenpal/docx-editor-core/types/content';
import {
  createCommentIdAllocator,
  seedCommentAllocator,
  type CommentIdAllocator,
} from '@eigenpal/docx-editor-core/prosemirror/commentIdAllocator';

/**
 * Owns the editor instance's comment/revision ID allocator (monotonic, no
 * reuse). Seeded above the loaded doc's max ID on load (useDocumentLoader);
 * shared by every comment/tracked-change allocation in DocxEditor and its
 * hooks.
 *
 * When `commentIdNamespace` is set (collab), the allocator mints inside that
 * client's private ID block, and is recreated if the namespace ever changes
 * (e.g. a new collab session with a new clientID).
 *
 * The re-seed effect runs whenever the comments array changes so an ID that
 * arrived from a collaboration peer (controlled `comments` prop) can never
 * be minted again locally. Cheap: O(comments); namespaced allocators ignore
 * foreign-block IDs internally. This alone fixes the confirmed cross-user
 * overwrite even without a namespace: replying to a remote thread re-seeds
 * above its ID before the reply mints.
 */
export function useCommentIdAllocator(
  commentIdNamespace: number | undefined,
  comments: Comment[]
): React.RefObject<CommentIdAllocator> {
  const commentIdAllocatorRef = useRef(createCommentIdAllocator(commentIdNamespace));
  const commentIdNamespaceRef = useRef(commentIdNamespace);
  if (commentIdNamespaceRef.current !== commentIdNamespace) {
    commentIdNamespaceRef.current = commentIdNamespace;
    commentIdAllocatorRef.current = createCommentIdAllocator(commentIdNamespace);
  }

  useEffect(() => {
    seedCommentAllocator(commentIdAllocatorRef.current, comments, null);
  }, [comments]);

  return commentIdAllocatorRef;
}

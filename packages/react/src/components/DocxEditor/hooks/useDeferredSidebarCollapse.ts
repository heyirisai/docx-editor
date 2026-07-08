import { useCallback, useRef } from 'react';

/**
 * Cursor-driven expand/collapse policy for the comment-sidebar thread:
 * the cursor rule may only collapse what the CURSOR expanded.
 *
 * The expanded thread is re-derived from the PM cursor on every selection
 * tick. Unowned collapses conflict with direct sidebar interaction: card
 * clicks TOGGLE the thread (UnifiedSidebar.toggleExpand) and do NOT move
 * the PM cursor or blur the PM view — so a tick deriving `null` right after
 * a card click would instantly re-close the thread the user just opened
 * (flicker), and in collaborative sessions remote transactions tick
 * constantly while the user reads or replies. Tracking whether the CURSOR
 * drove the current expansion makes repeat derivations idempotent:
 *
 * - item derived and unchanged → no-op (does not fight the card toggle)
 * - item derived and new      → expand + open the sidebar
 * - null derived, cursor owns the expansion → collapse (cursor moved away)
 * - null derived, user opened it via the card → leave it alone
 */
export function useDeferredSidebarCollapse({
  setShowCommentsSidebar,
  setExpandedSidebarItem,
}: {
  setShowCommentsSidebar: React.Dispatch<React.SetStateAction<boolean>>;
  setExpandedSidebarItem: React.Dispatch<React.SetStateAction<string | null>>;
}): {
  /** Feed the cursor-derived sidebar item (or null) on each selection tick. */
  applyCursorSidebarItem: (item: string | null) => void;
} {
  const lastItemRef = useRef<string | null>(null);

  const applyCursorSidebarItem = useCallback(
    (item: string | null) => {
      if (item) {
        if (lastItemRef.current !== item) {
          setShowCommentsSidebar(true);
          setExpandedSidebarItem(item);
        }
        lastItemRef.current = item;
        return;
      }
      if (lastItemRef.current === null) return;
      lastItemRef.current = null;
      setExpandedSidebarItem(null);
    },
    [setShowCommentsSidebar, setExpandedSidebarItem]
  );

  return { applyCursorSidebarItem };
}

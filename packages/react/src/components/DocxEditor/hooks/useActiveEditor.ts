import { useCallback } from 'react';
import type { PagedEditorRef } from '../PagedEditor';
import type { InlineHeaderFooterEditorRef } from '../../InlineHeaderFooterEditor';
import type { HistoryOverride } from '../types';

/**
 * Bundles the four "active editor" routing helpers — every callback
 * that needs to dispatch into PM checks whether the inline header/footer
 * editor is open and forwards to that view instead of the body's
 * PagedEditor. Keeps the routing rule in one place so callers don't
 * have to repeat the `hfEditPosition && hfEditorRef.current ? hf : body`
 * check.
 *
 * With a `historyOverride` (collab mode), BODY undo/redo run the override
 * commands; header/footer editors — separate non-collaborative PM docs
 * with their own built-in history — keep the default routing.
 */
export function useActiveEditor({
  hfEditPosition,
  hfEditorRef,
  pagedEditorRef,
  historyOverride,
}: {
  hfEditPosition: 'header' | 'footer' | null;
  hfEditorRef: React.RefObject<InlineHeaderFooterEditorRef | null>;
  pagedEditorRef: React.RefObject<PagedEditorRef | null>;
  historyOverride?: HistoryOverride;
}) {
  const getActiveEditorView = useCallback(() => {
    if (hfEditPosition && hfEditorRef.current) {
      return hfEditorRef.current.getView();
    }
    return pagedEditorRef.current?.getView();
  }, [hfEditPosition, hfEditorRef, pagedEditorRef]);

  const focusActiveEditor = useCallback(() => {
    if (hfEditPosition && hfEditorRef.current) {
      hfEditorRef.current.focus();
    } else {
      pagedEditorRef.current?.focus();
    }
  }, [hfEditPosition, hfEditorRef, pagedEditorRef]);

  const undoActiveEditor = useCallback(() => {
    if (hfEditPosition && hfEditorRef.current) {
      hfEditorRef.current.undo();
    } else if (historyOverride) {
      const view = pagedEditorRef.current?.getView();
      if (view) historyOverride.undo(view.state, view.dispatch, view);
    } else {
      pagedEditorRef.current?.undo();
    }
  }, [hfEditPosition, hfEditorRef, pagedEditorRef, historyOverride]);

  const redoActiveEditor = useCallback(() => {
    if (hfEditPosition && hfEditorRef.current) {
      hfEditorRef.current.redo();
    } else if (historyOverride) {
      const view = pagedEditorRef.current?.getView();
      if (view) historyOverride.redo(view.state, view.dispatch, view);
    } else {
      pagedEditorRef.current?.redo();
    }
  }, [hfEditPosition, hfEditorRef, pagedEditorRef, historyOverride]);

  return { getActiveEditorView, focusActiveEditor, undoActiveEditor, redoActiveEditor };
}

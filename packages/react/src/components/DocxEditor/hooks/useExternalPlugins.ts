import { useMemo, useState } from 'react';
import { keymap } from 'prosemirror-keymap';
import type { Plugin } from 'prosemirror-state';
import { createSuggestionModePlugin } from '@eigenpal/docx-editor-core/prosemirror/plugins';
import type { HistoryOverride } from '../types';
import type { EditorMode } from '../internals/editing-modes';

/**
 * Assembles the external plugin list handed to the hidden body EditorView:
 * the suggestion-mode plugin first, then the caller's `externalPlugins`
 * (e.g. ySyncPlugin — kept first among caller plugins so they can intercept
 * before extension keymaps), then the history-override keymap.
 *
 * Also latches the `historyOverride` prop on mount: the ExtensionManager is
 * built exactly once, so switching between built-in and override history
 * after mount is not supported. With the `history` extension disabled its
 * Mod-Z/Mod-Y bindings are gone — the injected keymap drives the override
 * commands instead.
 */
export function useExternalPlugins({
  historyOverride,
  editingMode,
  author,
  externalPlugins,
}: {
  historyOverride: HistoryOverride | undefined;
  editingMode: EditorMode;
  author: string;
  externalPlugins: Plugin[] | undefined;
}): {
  latchedHistoryOverride: HistoryOverride | undefined;
  allExternalPlugins: Plugin[];
} {
  const [latchedHistoryOverride] = useState(() => historyOverride);

  // Suggestion mode plugin — latched at mount (same as before extraction);
  // mode changes are applied via setSuggestionMode transactions.
  const suggestionPlugin = useMemo(
    () => createSuggestionModePlugin(editingMode === 'suggesting', author),
    [] // eslint-disable-line react-hooks/exhaustive-deps
  );

  const historyOverrideKeymap = useMemo(
    () =>
      latchedHistoryOverride
        ? keymap({
            'Mod-z': latchedHistoryOverride.undo,
            'Mod-y': latchedHistoryOverride.redo,
            'Mod-Shift-z': latchedHistoryOverride.redo,
          })
        : null,
    [latchedHistoryOverride]
  );

  const allExternalPlugins = useMemo(
    () => [
      suggestionPlugin,
      ...(externalPlugins ?? []),
      ...(historyOverrideKeymap ? [historyOverrideKeymap] : []),
    ],
    [suggestionPlugin, externalPlugins, historyOverrideKeymap]
  );

  return { latchedHistoryOverride, allExternalPlugins };
}

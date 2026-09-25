/**
 * Suggestion-mode plugin state + shared meta keys.
 *
 * Kept in its own module to avoid an import cycle: every handler file needs
 * the plugin key + meta constants but should not transitively pull in the
 * plugin factory.
 */

import { PluginKey } from 'prosemirror-state';

export interface SuggestionModeState {
  active: boolean;
  author: string;
}

export interface MarkAttrs {
  revisionId: number;
  author: string;
  date: string;
}

export const suggestionModeKey = new PluginKey<SuggestionModeState>('suggestionMode');

/** Set on transactions the plugin authored so it ignores them in `appendTransaction`. */
export const SUGGESTION_META = 'suggestionModeApplied';

/** Set by accept/reject commands to bypass suggesting-mode interception. */
export const SUGGESTION_BYPASS_META = 'suggestionModeBypass';

/**
 * Meta key y-prosemirror sets on the transactions its sync plugin dispatches
 * (`new PluginKey('y-sync')` → key string `y-sync$`). Such a transaction
 * replays a collaborator's already-committed edit; suggesting mode must not
 * re-wrap it as a tracked insertion by the local author.
 */
export const YJS_SYNC_META = 'y-sync$';

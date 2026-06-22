/**
 * @heyirisai/docx-editor-vue
 *
 * Curated root entry for the documented Vue 3 editor API. Advanced surfaces
 * stay public through explicit subpaths:
 * - `@heyirisai/docx-editor-vue/ui`
 * - `@heyirisai/docx-editor-vue/dialogs`
 * - `@heyirisai/docx-editor-vue/composables`
 * - `@heyirisai/docx-editor-vue/plugin-api`
 *
 * Framework-agnostic document utilities live in `@heyirisai/docx-editor-core`.
 * Agent/MCP surfaces live in `@heyirisai/docx-editor-agents`.
 *
 * @packageDocumentation
 * @public
 */

export const VERSION = '0.0.2';

// Main editor contract
export { default as DocxEditor } from './components/DocxEditor.vue';
export type { DocxEditorProps, EditorMode } from './components/DocxEditor/types';

// Document factory helpers — re-exported from `@heyirisai/docx-editor-core` so
// the common "spawn a blank editor" affordance is available without forcing
// consumers to add `-core` to their dependency tree alongside `-vue`.
export {
  createEmptyDocument,
  createDocumentWithText,
  type CreateEmptyDocumentOptions,
} from '@heyirisai/docx-editor-core';

// i18n contract — runtime only. Locale string types (LocaleStrings,
// Translations, PartialLocaleStrings, TranslationKey) live in
// `@heyirisai/docx-editor-i18n`; import them from there.
export { useTranslation, provideLocale, i18nPlugin, defaultLocale } from './i18n';

// renderAsync
export { renderAsync } from './renderAsync';
export type { DocxEditorHandle, RenderAsyncOptions } from './renderAsync';

// Public ref shape (typecheck contract with EditorRefLike — Decision 10).
export type { DocxEditorRef } from './components/DocxEditor/types';

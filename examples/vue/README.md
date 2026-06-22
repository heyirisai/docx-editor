# Vue example

`@heyirisai/docx-editor-vue` in a plain Vue 3 + Vite SPA. Same editor and same
surface as the React adapter, with Vue components and refs. No SSR, so the
editor mounts directly.

## Run it

From the repo root:

```bash
bun install
bun run dev:vue        # http://localhost:5174
```

Or from this directory: `bun run dev`.

## Files

| File             | What it does                                           |
| ---------------- | ------------------------------------------------------ |
| `src/App.vue`    | The editor: open `.docx`, edit, agent panel            |
| `src/main.ts`    | Vue app root + `@heyirisai/docx-editor-vue/styles.css` |
| `index.html`     | Loads the Material Symbols font for toolbar icons      |
| `vite.config.ts` | Aliases `@heyirisai/*` to workspace source in dev      |

## Minimal integration

```vue
<script setup lang="ts">
import { DocxEditor } from '@heyirisai/docx-editor-vue';
import '@heyirisai/docx-editor-vue/styles.css';
import { createEmptyDocument } from '@heyirisai/docx-editor-core';

const doc = createEmptyDocument();
</script>

<template>
  <DocxEditor :document="doc" :show-toolbar="true" />
</template>
```

To open a real file, read it as an `ArrayBuffer` and pass it as
`:document-buffer` instead of `:document`.

## Use it in your own Vue app

```bash
npm install @heyirisai/docx-editor-vue @heyirisai/docx-editor-core
```

Unlike the React adapter, the Vue adapter ships a stylesheet you must import
once: `@heyirisai/docx-editor-vue/styles.css`. Add the Material Symbols font
to `index.html`:

```html
<link
  rel="stylesheet"
  href="https://fonts.googleapis.com/css2?family=Material+Symbols+Outlined:opsz,wght,FILL,GRAD@20..48,100..700,0..1,-50..200&display=block"
/>
```

Docs: https://www.docx-editor.dev/docs/1.x/vue

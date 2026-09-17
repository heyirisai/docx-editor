import { defineConfig } from 'tsup';

// Tsup builds the framework-agnostic + React entries. Vue SFCs are built by
// `vite.config.ts` because tsup/esbuild can't compile `.vue` files. The
// dedicated `tsconfig.tsup.json` excludes vue/* so the d.ts pass doesn't
// trip on the SFC shim.
export default defineConfig({
  entry: {
    index: 'src/index.ts',
    bridge: 'src/bridge.ts',
    server: 'src/server.ts',
    react: 'src/react.ts',
    mcp: 'src/mcp/index.ts',
    'ai-sdk/server': 'src/ai-sdk/server.ts',
    'ai-sdk/react': 'src/ai-sdk/react.ts',
  },
  format: ['cjs', 'esm'],
  dts: { resolve: true },
  tsconfig: 'tsconfig.tsup.json',
  splitting: true,
  sourcemap: false,
  clean: true,
  treeshake: {
    preset: 'smallest',
  },
  minify: true,
  noExternal: ['@eigenpal/docx-editor-core'],
  // Every prosemirror package must stay external, NOT just the ones this
  // package imports directly: `noExternal` inlines core, and core imports the
  // rest. Bundling a copy of prosemirror-tables re-runs its module-level
  // `Selection.jsonID('cell', CellSelection)` beside the host app's copy, which
  // throws "Duplicate use of selection JSON ID cell" at import time. Mirrors
  // the `/^prosemirror-/` external in vite.config.ts.
  external: [/^prosemirror-/, 'react', 'ai'],
});

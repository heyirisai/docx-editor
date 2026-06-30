#!/usr/bin/env node
/**
 * Publish-time brand rename: @eigenpal/* -> @heyirisai/*
 *
 * WHY: this fork tracks upstream `eigenpal/docx-editor` on `main` with the
 * source kept byte-for-byte under the `@eigenpal/*` npm scope, so upstream
 * commits merge in without conflicts. We still want to ship our own build to
 * npm under `@heyirisai/*`. Doing the rename IN SOURCE (the abandoned
 * `heyiris-release` branch) means every upstream sync fights ~491 renamed
 * import lines. Instead we rename ONLY at publish time, in the ephemeral CI
 * workspace, right after `build:packages` and right before `changeset publish`
 * — the rename is never committed.
 *
 * WHAT: swaps the npm scope `@eigenpal/` -> `@heyirisai/` in:
 *   - each publishable package's package.json (its `name` + any internal
 *     `@eigenpal/*` dependency specifiers)
 *   - everything under each package's `dist/` (the built cross-package import
 *     specifiers, e.g. `require('@eigenpal/docx-editor-core')`, and the CSS
 *     subpath imports)
 *   - `.changeset/config.json` (the `fixed` group package names)
 *
 * It deliberately matches the scope prefix `@eigenpal/` (with the leading `@`
 * and trailing `/`) so it only touches npm package specifiers — NOT the
 * upstream `github.com/eigenpal` repo URLs or author emails, which must stay.
 *
 * Safe to run more than once (idempotent): a second pass finds no `@eigenpal/`.
 *
 * Usage:
 *   node scripts/brand-rename.mjs            # rewrite in place
 *   node scripts/brand-rename.mjs --dry-run  # report what would change, write nothing
 */

import { readFileSync, writeFileSync, existsSync, readdirSync, statSync } from 'node:fs';
import { join, dirname, extname } from 'node:path';
import { fileURLToPath } from 'node:url';

const OLD_SCOPE = '@eigenpal/';
const NEW_SCOPE = '@heyirisai/';

const repoRoot = join(dirname(fileURLToPath(import.meta.url)), '..');
const DRY_RUN = process.argv.includes('--dry-run');

// The publishable workspace packages (must match root `build:packages`).
const PACKAGE_DIRS = ['core', 'react', 'agents', 'i18n', 'vue', 'nuxt'].map((p) =>
  join(repoRoot, 'packages', p)
);

// Only rewrite text files we know carry module specifiers. Skip binaries
// (fonts/images) that may live under dist/.
const TEXT_EXTENSIONS = new Set([
  '.js',
  '.mjs',
  '.cjs',
  '.ts',
  '.mts',
  '.cts',
  '.json',
  '.css',
  '.map',
  '.d.ts',
]);

let filesChanged = 0;
let occurrences = 0;

/** Replace OLD_SCOPE in a single file; returns true if it changed. */
function rewriteFile(path) {
  const before = readFileSync(path, 'utf8');
  if (!before.includes(OLD_SCOPE)) return false;
  const count = before.split(OLD_SCOPE).length - 1;
  const after = before.split(OLD_SCOPE).join(NEW_SCOPE);
  occurrences += count;
  filesChanged += 1;
  if (DRY_RUN) {
    console.log(`  would rewrite ${count.toString().padStart(4)}x  ${path.replace(`${repoRoot}/`, '')}`);
  } else {
    writeFileSync(path, after);
  }
  return true;
}

/** Recursively rewrite every text file under `dir`. */
function rewriteDir(dir) {
  if (!existsSync(dir)) return;
  for (const entry of readdirSync(dir)) {
    const full = join(dir, entry);
    const st = statSync(full);
    if (st.isDirectory()) {
      rewriteDir(full);
    } else if (TEXT_EXTENSIONS.has(extname(full)) || full.endsWith('.d.mts')) {
      rewriteFile(full);
    }
  }
}

console.log(
  `${DRY_RUN ? '[dry-run] ' : ''}Brand rename ${OLD_SCOPE}* -> ${NEW_SCOPE}*`
);

for (const pkgDir of PACKAGE_DIRS) {
  const pkgJson = join(pkgDir, 'package.json');
  if (existsSync(pkgJson)) rewriteFile(pkgJson);
  rewriteDir(join(pkgDir, 'dist'));
}

// Keep changeset's fixed-group package names consistent with the renamed
// package.json names so `changeset publish` resolves the workspace cleanly.
const changesetConfig = join(repoRoot, '.changeset', 'config.json');
if (existsSync(changesetConfig)) rewriteFile(changesetConfig);

console.log(
  `${DRY_RUN ? '[dry-run] ' : ''}Done: ${occurrences} occurrence(s) across ${filesChanged} file(s).`
);

if (!DRY_RUN && filesChanged === 0) {
  console.warn(
    'WARNING: no @eigenpal/ occurrences found. Did build:packages run first, ' +
      'or has the rename already been applied?'
  );
}

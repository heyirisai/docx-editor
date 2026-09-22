/**
 * The agents bundle inlines core (`noExternal`) but must NOT inline core's
 * peer libraries. Two failure modes this guards:
 *
 *  - an inlined `prosemirror-tables` re-runs its module-level
 *    `Selection.jsonID('cell', CellSelection)` beside the host app's copy and
 *    throws "Duplicate use of selection JSON ID cell" at import time;
 *  - an externalized import that `package.json` never declares cannot be
 *    resolved by a strict consumer.
 *
 * So the tsup `external` list and the declared peers have to stay in step.
 */
import { describe, expect, test } from 'bun:test';
import { existsSync, readdirSync, readFileSync, statSync } from 'node:fs';
import { join } from 'node:path';
import pkg from '../../package.json';

const DIST = join(import.meta.dir, '../../dist');

function distFiles(dir: string): string[] {
  if (!existsSync(dir)) return [];
  return readdirSync(dir).flatMap((entry) => {
    const full = join(dir, entry);
    if (statSync(full).isDirectory()) return distFiles(full);
    return /\.(mjs|js|cjs)$/.test(entry) ? [full] : [];
  });
}

const files = distFiles(DIST);
const bundle = files.map((f) => readFileSync(f, 'utf8')).join('\n');

describe.skipIf(files.length === 0)('agents bundle dependency hygiene', () => {
  test('no prosemirror package is inlined', () => {
    // `Selection.jsonID("cell", ...)` only exists inside prosemirror-tables.
    expect(bundle).not.toMatch(/jsonID\(\s*["']cell["']/);
  });

  test('every prosemirror package it imports is declared', () => {
    const imported = [...bundle.matchAll(/["'](prosemirror-[a-z-]+)["']/g)].map((m) => m[1]);
    const declared = new Set([
      ...Object.keys(pkg.peerDependencies ?? {}),
      ...Object.keys((pkg as { dependencies?: Record<string, string> }).dependencies ?? {}),
    ]);
    expect([...new Set(imported)].filter((name) => !declared.has(name))).toEqual([]);
  });
});

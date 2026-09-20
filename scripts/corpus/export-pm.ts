/**
 * Export a `.docx` through the editor's REAL save path, headlessly.
 *
 * The app saves PM state, not the parsed model: `parser` → `toProseDoc` →
 * (edits) → `fromProseDoc` → `serializer` → `rezip`. A plain
 * `parse → repack` skips the two conversions and misses everything they
 * introduce — the missing `<wps:bodyPr/>` that made Word refuse a COMET export
 * only appears on this path, because `fromProseDoc` is what rebuilds the shape.
 *
 *   bun scripts/corpus/export-pm.ts <in.docx> <out.docx>
 *
 * Pair it with `tier3-word-opens.sh`, which feeds the results to real Word.
 */

import { readFileSync, writeFileSync } from 'node:fs';
import { fullParseDocx } from '../../packages/core/src/docx/parser';
import { toProseDoc } from '../../packages/core/src/prosemirror/conversion/toProseDoc';
import { fromProseDoc } from '../../packages/core/src/prosemirror/conversion/fromProseDoc';
import { repackDocx } from '../../packages/core/src/docx/rezip';

const [input, output] = process.argv.slice(2);
if (!input || !output) {
  console.error('usage: bun scripts/corpus/export-pm.ts <in.docx> <out.docx>');
  process.exit(2);
}

const raw = new Uint8Array(readFileSync(input)).buffer as ArrayBuffer;
const doc = await fullParseDocx(raw);
const packed = await repackDocx(fromProseDoc(toProseDoc(doc), doc));
writeFileSync(output, Buffer.from(packed));
console.log(`${output} — ${Buffer.from(packed).length} bytes`);

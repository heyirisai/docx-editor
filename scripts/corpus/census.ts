/**
 * Feature census — what does a real corpus actually USE, and do we keep it?
 *
 * "Render 95% of what Word can render" is unmeasurable as stated: ECMA-376 has
 * thousands of elements and a typical document touches a few hundred. The
 * denominator that matters is not the spec, it is USAGE — an element that
 * appears 40,000 times across the corpus matters more than one that appears
 * twice, and the spec weights them equally.
 *
 * So this counts every element INSTANCE in every XML part of every input, then
 * exports each document through the REAL save path (parser → toProseDoc →
 * fromProseDoc → serializer → rezip) and counts again. Coverage is the share of
 * input element instances that survive.
 *
 *   bun scripts/corpus/census.ts <dir>              # ranked report
 *   bun scripts/corpus/census.ts <dir> --json o.json
 *   bun scripts/corpus/census.ts <dir> --top 40     # how many rows to print
 *
 * Survival is a LOWER bound on fidelity, not a proof of it: an element we keep
 * verbatim in preserved XML counts as survived even if the painter ignores it.
 * That gap is exactly what Tier 2 is for. What this CAN prove is the opposite —
 * anything that does not survive is definitively not rendered.
 */

import { readFileSync, writeFileSync, statSync, readdirSync } from 'node:fs';
import { join, extname, basename } from 'node:path';
import { fullParseDocx } from '../../packages/core/src/docx/parser';
import { toProseDoc } from '../../packages/core/src/prosemirror/conversion/toProseDoc';
import { fromProseDoc } from '../../packages/core/src/prosemirror/conversion/fromProseDoc';
import { repackDocx } from '../../packages/core/src/docx/rezip';

/** Parts whose elements describe rendered content. */
const COUNTED_PART =
  /^word\/(document|header\d*|footer\d*|footnotes|endnotes|numbering|styles|theme\/theme\d*)\.xml$/;

/**
 * Elements that are SUPPOSED to disappear, and must not count as loss.
 *
 * `mc:Fallback` holds a legacy VML twin of a DrawingML shape; Word renders the
 * `mc:Choice` and ignores the fallback, and so do we. `w:proofErr` is spell
 * checker state. The `w:rsid*` family is Word's merge-tracking bookkeeping and
 * carries no formatting. Counting these as dropped features measures the
 * runner, not the engine — the same trap Tier 1's text check fell into.
 */
const EXPECTED_TO_DROP = new Set([
  'mc:Fallback',
  'mc:Choice',
  'mc:AlternateContent',
  'w:proofErr',
  'w:rsid',
  'w:rsids',
  'w:noProof',
  'w:lastRenderedPageBreak',
  'w:bookmarkStart',
  'w:bookmarkEnd',
]);

type Census = Map<string, number>;

/**
 * Count element instances by qualified name.
 *
 * Deliberately a scanner and not a DOM parse: this must run over attacker-
 * controlled XML from an untrusted `.docx`, and a regex that only ever reads
 * tag names cannot resolve an entity, fetch a DTD or allocate from a
 * file-supplied number. It never builds a tree, so nesting depth is irrelevant.
 */
function censusOf(xml: string, into: Census): void {
  const tag = /<([a-zA-Z_][\w.-]*(?::[a-zA-Z_][\w.-]*)?)[\s/>]/g;
  let m: RegExpExecArray | null;
  while ((m = tag.exec(xml)) !== null) {
    const name = m[1];
    into.set(name, (into.get(name) ?? 0) + 1);
  }
}

async function censusOfDocx(buf: ArrayBuffer): Promise<Census> {
  const { default: JSZip } = await import('jszip');
  const zip = await JSZip.loadAsync(buf);
  const out: Census = new Map();
  for (const path of Object.keys(zip.files)) {
    if (!COUNTED_PART.test(path)) continue;
    const xml = await zip.files[path].async('string');
    censusOf(xml, out);
  }
  return out;
}

function listDocx(target: string): string[] {
  const st = statSync(target);
  if (!st.isDirectory()) return [target];
  return readdirSync(target)
    .filter((f) => extname(f).toLowerCase() === '.docx' && !f.startsWith('~$'))
    .map((f) => join(target, f))
    .sort();
}

// ---------------------------------------------------------------------------

const args = process.argv.slice(2);
const target = args.find((a) => !a.startsWith('--'));
const jsonAt = args.indexOf('--json');
const topAt = args.indexOf('--top');
const TOP = topAt >= 0 ? Number(args[topAt + 1]) : 25;

if (!target) {
  console.error('usage: bun scripts/corpus/census.ts <dir-or-file> [--json out.json] [--top N]');
  process.exit(2);
}

const files = listDocx(target);
const before: Census = new Map();
const after: Census = new Map();
/** Elements seen in the input, and in how many DISTINCT documents. */
const docsUsing: Map<string, number> = new Map();
const failed: string[] = [];

for (const file of files) {
  const raw = new Uint8Array(readFileSync(file)).buffer as ArrayBuffer;
  let inC: Census;
  let outC: Census;
  try {
    inC = await censusOfDocx(raw);
    const doc = await fullParseDocx(raw);
    const packed = await repackDocx(fromProseDoc(toProseDoc(doc), doc));
    outC = await censusOfDocx(new Uint8Array(packed).buffer as ArrayBuffer);
  } catch (err) {
    failed.push(`${basename(file)}: ${(err as Error).message}`);
    continue;
  }
  for (const [k, v] of inC) {
    before.set(k, (before.get(k) ?? 0) + v);
    docsUsing.set(k, (docsUsing.get(k) ?? 0) + 1);
  }
  for (const [k, v] of outC) after.set(k, (after.get(k) ?? 0) + v);
  process.stderr.write(`. ${basename(file)}\n`);
}

interface Row {
  name: string;
  in: number;
  out: number;
  docs: number;
  kept: number;
  expected: boolean;
}

const rows: Row[] = [...before.entries()]
  .map(([name, n]) => {
    const out = after.get(name) ?? 0;
    return {
      name,
      in: n,
      out,
      docs: docsUsing.get(name) ?? 0,
      kept: Math.min(1, out / n),
      expected: EXPECTED_TO_DROP.has(name),
    };
  })
  .sort((a, b) => b.in - a.in);

const counted = rows.filter((r) => !r.expected);
const totalIn = counted.reduce((s, r) => s + r.in, 0);
const totalKept = counted.reduce((s, r) => s + Math.min(r.in, r.out), 0);
const lost = counted.filter((r) => r.out === 0);
const partial = counted.filter((r) => r.out > 0 && r.out < r.in * 0.98);

console.log(
  `\n${files.length - failed.length}/${files.length} documents, ${counted.length} distinct elements in use\n`
);
console.log(
  `instance coverage   ${((totalKept / totalIn) * 100).toFixed(2)}%  (${totalKept.toLocaleString()} / ${totalIn.toLocaleString()} element instances survive the save path)`
);
console.log(
  `element coverage    ${(((counted.length - lost.length) / counted.length) * 100).toFixed(2)}%  (${counted.length - lost.length} / ${counted.length} distinct elements survive at all)\n`
);

if (lost.length) {
  console.log(`DROPPED ENTIRELY — ranked by how much the corpus uses them\n`);
  console.log(`  ${'element'.padEnd(28)} ${'uses'.padStart(7)} ${'docs'.padStart(5)}`);
  for (const r of lost.slice(0, TOP)) {
    console.log(`  ${r.name.padEnd(28)} ${String(r.in).padStart(7)} ${String(r.docs).padStart(5)}`);
  }
  if (lost.length > TOP) console.log(`  ... and ${lost.length - TOP} more`);
  console.log('');
}

if (partial.length) {
  console.log(`PARTIALLY LOST\n`);
  console.log(
    `  ${'element'.padEnd(28)} ${'in'.padStart(7)} ${'out'.padStart(7)} ${'kept'.padStart(6)}`
  );
  for (const r of partial.slice(0, TOP)) {
    console.log(
      `  ${r.name.padEnd(28)} ${String(r.in).padStart(7)} ${String(r.out).padStart(7)} ${(r.kept * 100).toFixed(0).padStart(5)}%`
    );
  }
  console.log('');
}

if (failed.length) {
  console.log(`FAILED TO PROCESS\n${failed.map((f) => '  ' + f).join('\n')}\n`);
}

if (jsonAt >= 0 && args[jsonAt + 1]) {
  writeFileSync(args[jsonAt + 1], JSON.stringify({ rows, totalIn, totalKept, failed }, null, 2));
  console.log(`wrote ${args[jsonAt + 1]}`);
}

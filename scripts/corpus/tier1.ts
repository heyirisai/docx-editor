/**
 * Tier 1 corpus runner — oracle-free conformance checks over arbitrary `.docx`.
 *
 * The template suite tells us how close we are to Word on thirteen documents we
 * chose. It cannot tell us how the engine behaves on the documents we have
 * never seen, and those are the ones that break in front of a user. This runner
 * answers the second question: point it at any pile of `.docx` files and it
 * asserts the things that must hold for EVERY document, with no reference
 * render, no browser and no human judgement.
 *
 *   bun run corpus:tier1 -- <dir-or-glob>            # run and print a report
 *   bun run corpus:tier1 -- <dir> --json out.json    # machine-readable report
 *   bun run corpus:tier1 -- <dir> --baseline b.json  # fail on regression
 *
 * What it deliberately does NOT check: anything about where pixels land. Layout
 * needs text measurement, which needs a browser — that is Tier 2. Everything
 * here runs headless in Bun at roughly a dozen documents a second, which is
 * what makes a corpus of thousands practical.
 */

import { readFileSync, writeFileSync, statSync, readdirSync } from 'node:fs';
import { join, extname, basename } from 'node:path';
import { fullParseDocx } from '../../packages/core/src/docx/parser';
import { serializeDocument } from '../../packages/core/src/docx/serializer/documentSerializer';
import { repackDocx } from '../../packages/core/src/docx/rezip';
import type { Document } from '../../packages/core/src/types/document';

/** One assertion about one document. `null` detail means it passed. */
interface CheckResult {
  name: string;
  ok: boolean;
  detail?: string;
}

interface DocReport {
  file: string;
  bytes: number;
  ms: number;
  checks: CheckResult[];
  /** Fraction of the source's visible text still present after a round trip. */
  textRetained: number | null;
  /** The actual source substrings that went missing — the actionable half. */
  textMissing: string[];
  /** Parser warnings — not failures, but the coverage signal worth watching. */
  warnings: string[];
}

const TEXT_RETENTION_FLOOR = 0.999;

// ---------------------------------------------------------------------------
// Text extraction
// ---------------------------------------------------------------------------

/**
 * Drop every `<mc:Fallback>` subtree. A cover-page text box is authored twice —
 * once as `mc:Choice` (DrawingML) and once as `mc:Fallback` (legacy VML) — and
 * the two copies hold the SAME words. Word renders the Choice and ignores the
 * Fallback; so do we, so counting the source's fallback text as "lost on save"
 * is measuring the runner, not the engine. (Higher Education scored 86.7% for
 * exactly this reason: its whole cover page is duplicated in a fallback.)
 *
 * Nesting is legal (an AlternateContent inside a fallback), so this tracks
 * depth rather than pairing with a non-greedy regex.
 */
function stripAlternateFallbacks(xml: string): string {
  const OPEN = '<mc:Fallback>';
  const CLOSE = '</mc:Fallback>';
  let out = '';
  let i = 0;
  for (;;) {
    const start = xml.indexOf(OPEN, i);
    if (start < 0) return out + xml.slice(i);
    out += xml.slice(i, start);
    let depth = 1;
    let j = start + OPEN.length;
    while (depth > 0) {
      const nextOpen = xml.indexOf(OPEN, j);
      const nextClose = xml.indexOf(CLOSE, j);
      if (nextClose < 0) return out; // unbalanced — drop the rest
      if (nextOpen >= 0 && nextOpen < nextClose) {
        depth++;
        j = nextOpen + OPEN.length;
      } else {
        depth--;
        j = nextClose + CLOSE.length;
      }
    }
    i = j;
  }
}

/**
 * Concatenate every `<w:t>` run in a `document.xml`, normalised so that
 * whitespace and entity differences don't register as data loss. This is a
 * deliberately dumb comparison: it asks "are the characters still there", not
 * "are they in the same runs", because re-running splits is legal and reshaping
 * them is not data loss.
 */
function extractText(documentXml: string): string {
  const out: string[] = [];
  const re = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g;
  let m: RegExpExecArray | null;
  const scoped = stripAlternateFallbacks(documentXml);
  while ((m = re.exec(scoped)) !== null) out.push(m[1] ?? '');
  return out
    .join('')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
    .replace(/\s+/g, ' ')
    .trim();
}

/** Pull `word/document.xml` out of a .docx buffer without a full parse. */
async function readDocumentXml(buffer: ArrayBuffer): Promise<string> {
  const { default: JSZip } = await import('jszip');
  const zip = await JSZip.loadAsync(buffer);
  const entry = zip.file('word/document.xml');
  if (!entry) throw new Error('no word/document.xml');
  return entry.async('string');
}

/**
 * What fraction of the source's text survives a round trip, sampled in chunks
 * so that re-splitting runs or reordering attributes doesn't read as loss.
 *
 * The chunk size scales with the document: a fixed 40 characters gave a cover
 * page with 200 characters of body text only five samples, so a single miss
 * scored it 80% and the metric looked like data loss when it was arithmetic.
 * Aim for ~100 samples, floored at 12 characters so a chunk stays specific
 * enough that matching it means something.
 *
 * A bare percentage is not actionable — "97.3% retained" tells you something
 * broke but not what, and a chunk boundary that lands mid-word can miss for
 * reasons that are not loss at all. So every miss is narrowed to the LONGEST
 * run of characters in it that genuinely isn't in the output, and reported.
 * That narrowing is what makes the number trustworthy: if the narrowed string
 * is empty the chunk was a sampling artefact, not data loss.
 */
interface RetentionReport {
  ratio: number;
  /** Deduped source substrings that are genuinely absent from the output. */
  missing: string[];
}

/**
 * Shrink a missing chunk to the smallest substring that is still absent. A
 * chunk can miss because the round trip re-split a run mid-chunk while keeping
 * every character; walking in from both ends finds the real gap, or proves
 * there isn't one.
 */
function narrowMiss(chunk: string, after: string): string {
  let lo = 0;
  let hi = chunk.length;
  // Trim from the left while the remainder is still missing.
  while (lo < hi && !after.includes(chunk.slice(lo + 1, hi))) lo++;
  // Then from the right.
  while (hi > lo + 1 && !after.includes(chunk.slice(lo, hi - 1))) hi--;
  const core = chunk.slice(lo, hi);
  return after.includes(core) ? '' : core;
}

function retention(before: string, after: string): RetentionReport {
  if (before.length === 0) return { ratio: 1, missing: [] };
  if (after.length === 0) return { ratio: 0, missing: [before.slice(0, 80)] };
  const chunk = Math.max(12, Math.min(48, Math.floor(before.length / 100)));
  if (before.length <= chunk) {
    return after.includes(before) ? { ratio: 1, missing: [] } : { ratio: 0, missing: [before] };
  }
  let found = 0;
  let total = 0;
  const missing = new Set<string>();
  for (let i = 0; i + chunk <= before.length; i += chunk) {
    total++;
    const slice = before.slice(i, i + chunk);
    if (after.includes(slice)) {
      found++;
      continue;
    }
    const core = narrowMiss(slice, after);
    // An empty core means the chunk boundary, not the content, was the
    // problem — every character in it is still present. Count it as retained.
    if (core === '') found++;
    else missing.add(core);
  }
  return { ratio: total === 0 ? 1 : found / total, missing: [...missing] };
}

// ---------------------------------------------------------------------------
// The checks
// ---------------------------------------------------------------------------

async function checkDocument(path: string): Promise<DocReport> {
  const t0 = Date.now();
  const bytes = statSync(path).size;
  const checks: CheckResult[] = [];
  const warnings: string[] = [];
  let textRetained: number | null = null;
  let textMissing: string[] = [];

  const pass = (name: string) => checks.push({ name, ok: true });
  const fail = (name: string, detail: string) => checks.push({ name, ok: false, detail });
  const short = (e: unknown) => String(e instanceof Error ? e.message : e).slice(0, 160);

  const raw = new Uint8Array(readFileSync(path)).buffer as ArrayBuffer;

  // 1. Parse. Everything else depends on this, so bail out if it throws.
  let doc: Document;
  try {
    doc = await fullParseDocx(raw);
    pass('parse');
  } catch (e) {
    fail('parse', short(e));
    return { file: path, bytes, ms: Date.now() - t0, checks, textRetained, textMissing, warnings };
  }
  if (doc.warnings) warnings.push(...doc.warnings);

  // 2. Serialize the body.
  let xml1: string;
  try {
    xml1 = serializeDocument(doc);
    pass('serialize');
  } catch (e) {
    fail('serialize', short(e));
    return { file: path, bytes, ms: Date.now() - t0, checks, textRetained, textMissing, warnings };
  }

  // 3. Serializing the same model twice must give the same bytes. A failure
  //    here means the serializer reads mutable state — the kind of bug that
  //    makes a save depend on how many times you looked at the document.
  try {
    if (serializeDocument(doc) === xml1) pass('serialize-deterministic');
    else fail('serialize-deterministic', 'two serializations of one model differ');
  } catch (e) {
    fail('serialize-deterministic', short(e));
  }

  // 4. Repack to a real .docx and parse it back. This is the save path a user
  //    actually takes, so it has to survive its own output.
  let doc2: Document | null = null;
  try {
    const repacked = await repackDocx(doc);
    pass('repack');
    try {
      doc2 = await fullParseDocx(repacked);
      pass('reparse');
    } catch (e) {
      fail('reparse', short(e));
    }

    // 5. Text survives the round trip. Compared against the SOURCE's own
    //    document.xml, so this catches content we silently drop on the way in
    //    as well as on the way out.
    try {
      const srcText = extractText(await readDocumentXml(raw));
      const outText = extractText(await readDocumentXml(repacked));
      const report = retention(srcText, outText);
      textRetained = report.ratio;
      textMissing = report.missing;
      if (report.ratio >= TEXT_RETENTION_FLOOR) pass('text-preserved');
      else {
        const sample = report.missing
          .slice(0, 5)
          .map((m) => JSON.stringify(m))
          .join(', ');
        fail(
          'text-preserved',
          `${(report.ratio * 100).toFixed(1)}% of ${srcText.length} source chars found; ` +
            `${report.missing.length} missing string(s): ${sample}` +
            (report.missing.length > 5 ? ' …' : '')
        );
      }
    } catch (e) {
      fail('text-preserved', short(e));
    }
  } catch (e) {
    fail('repack', short(e));
  }

  // 6. Round-trip stability: the model from the repacked file must serialize to
  //    what the first pass produced. If pass 2 differs from pass 1, every save
  //    keeps mutating the file — the drift that turns a template into garbage
  //    after a dozen edits.
  if (doc2) {
    try {
      const xml2 = serializeDocument(doc2);
      if (xml2 === xml1) pass('roundtrip-stable');
      else {
        let at = 0;
        while (at < xml1.length && xml1[at] === xml2[at]) at++;
        fail(
          'roundtrip-stable',
          `diverges at char ${at}: ${JSON.stringify(xml1.slice(at, at + 60))} vs ${JSON.stringify(xml2.slice(at, at + 60))}`
        );
      }
    } catch (e) {
      fail('roundtrip-stable', short(e));
    }
  }

  return { file: path, bytes, ms: Date.now() - t0, checks, textRetained, textMissing, warnings };
}

// ---------------------------------------------------------------------------
// Runner
// ---------------------------------------------------------------------------

function collectDocx(target: string): string[] {
  const st = statSync(target);
  if (!st.isDirectory()) return [target];
  const out: string[] = [];
  const walk = (dir: string) => {
    for (const entry of readdirSync(dir, { withFileTypes: true })) {
      const p = join(dir, entry.name);
      if (entry.isDirectory()) walk(p);
      // Word's lock files (`~$name.docx`) are not documents.
      else if (extname(entry.name).toLowerCase() === '.docx' && !entry.name.startsWith('~$'))
        out.push(p);
    }
  };
  walk(target);
  return out.sort();
}

function argValue(flag: string): string | undefined {
  const i = process.argv.indexOf(flag);
  return i >= 0 ? process.argv[i + 1] : undefined;
}

async function main(): Promise<void> {
  const target = process.argv.slice(2).find((a) => !a.startsWith('--'));
  if (!target) {
    console.error('usage: bun run corpus:tier1 -- <dir-or-file> [--json out] [--baseline in]');
    process.exit(2);
  }

  const files = collectDocx(target);
  if (files.length === 0) {
    console.error(`no .docx found under ${target}`);
    process.exit(2);
  }

  const t0 = Date.now();
  const reports: DocReport[] = [];
  for (const f of files) {
    try {
      reports.push(await checkDocument(f));
    } catch (e) {
      // A throw out here is the runner's bug, not the document's — record it
      // rather than losing the whole run.
      reports.push({
        file: f,
        bytes: 0,
        ms: 0,
        checks: [{ name: 'runner', ok: false, detail: String(e).slice(0, 160) }],
        textRetained: null,
        textMissing: [],
        warnings: [],
      });
    }
  }
  const elapsed = (Date.now() - t0) / 1000;

  // Per-check tally.
  const names = [...new Set(reports.flatMap((r) => r.checks.map((c) => c.name)))];
  const failed = reports.filter((r) => r.checks.some((c) => !c.ok));

  console.log(`\n${reports.length} documents in ${elapsed.toFixed(1)}s`);
  console.log(`${(reports.length / Math.max(elapsed, 0.001)).toFixed(1)} docs/sec\n`);
  for (const n of names) {
    const ran = reports.filter((r) => r.checks.some((c) => c.name === n));
    const ok = ran.filter((r) => r.checks.find((c) => c.name === n)?.ok);
    const bar = ran.length ? Math.round((ok.length / ran.length) * 20) : 0;
    console.log(
      `  ${n.padEnd(24)} ${String(ok.length).padStart(5)}/${String(ran.length).padEnd(5)} ` +
        `${'#'.repeat(bar)}${'.'.repeat(20 - bar)}`
    );
  }

  if (failed.length > 0) {
    console.log(`\n${failed.length} document(s) with failures:\n`);
    for (const r of failed.slice(0, 40)) {
      console.log(`  ${basename(r.file)}`);
      for (const c of r.checks.filter((x) => !x.ok)) console.log(`      ${c.name}: ${c.detail}`);
    }
    if (failed.length > 40) console.log(`  … and ${failed.length - 40} more`);
  }

  const warned = reports.filter((r) => r.warnings.length > 0);
  if (warned.length > 0) {
    const tally = new Map<string, number>();
    for (const r of warned) for (const w of r.warnings) tally.set(w, (tally.get(w) ?? 0) + 1);
    console.log(`\nparser warnings (${warned.length} documents):`);
    for (const [w, n] of [...tally].sort((a, b) => b[1] - a[1]).slice(0, 15)) {
      console.log(`  ${String(n).padStart(5)}  ${w.slice(0, 90)}`);
    }
  }

  const jsonPath = argValue('--json');
  if (jsonPath) {
    writeFileSync(jsonPath, JSON.stringify({ generated: Date.now(), reports }, null, 2));
    console.log(`\nreport → ${jsonPath}`);
  }

  // Regression gate. The absolute pass rate will never be 100% on a real
  // corpus, so CI compares against a stored baseline and fails only when a
  // document that used to pass a check stops passing it.
  const baselinePath = argValue('--baseline');
  if (baselinePath) {
    const base = JSON.parse(readFileSync(baselinePath, 'utf8')) as { reports: DocReport[] };
    const was = new Map<string, boolean>();
    for (const r of base.reports)
      for (const c of r.checks) was.set(`${basename(r.file)}::${c.name}`, c.ok);

    const regressions: string[] = [];
    for (const r of reports) {
      for (const c of r.checks) {
        const key = `${basename(r.file)}::${c.name}`;
        if (was.get(key) === true && !c.ok) regressions.push(`${key} — ${c.detail}`);
      }
    }
    if (regressions.length > 0) {
      console.log(`\nREGRESSIONS vs baseline (${regressions.length}):`);
      for (const r of regressions.slice(0, 40)) console.log(`  ${r}`);
      process.exit(1);
    }
    console.log('\nno regressions vs baseline.');
  }
}

await main();

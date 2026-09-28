/**
 * Words and spaces that sit in different runs lay out like Word.
 *
 * PDF-converted questionnaires split every word and every space into its own run, each carrying
 * character spacing (`w:spacing`) and horizontal scale (`w:w`):
 *
 *   <w:r><w:rPr><w:w w:val="110"/></w:rPr><w:t>Customer</w:t></w:r>
 *   <w:r><w:rPr><w:spacing w:val="6"/><w:w w:val="110"/></w:rPr><w:t xml:space="preserve"> </w:t></w:r>
 *   <w:r><w:rPr><w:w w:val="110"/></w:rPr><w:t>Records</w:t></w:r>
 *
 * and split words mid-word (`<w:t xml:space="preserve"> r</w:t>` +
 * `<w:t>etained?</w:t>`).
 *
 * Word's truth:
 *  - w:w scales each glyph advance, so the following run starts after the
 *    SCALED width. The painter used a paint-only `transform: scaleX()`, whose
 *    layout box keeps the unscaled width: a 110% word's glyphs overhung the
 *    next run and swallowed the single-space run ("CustomerRecords").
 *  - w:spacing is added after every character, including a lone space.
 *  - a run boundary is not a line-break opportunity: "r" + "etained?" wraps
 *    as one word (the breaker used to wrap between the runs: "r / etained?").
 *
 * A real document with this run shape can be pointed at via
 * `DOCX_FIXTURE_RUN_SPACING` for an extra whole-document check (skipped when
 * unset).
 */

import { describe, test, expect, beforeAll, afterAll } from 'bun:test';
import fs from 'node:fs';
import JSZip from 'jszip';
import { installCanvasDocumentStub } from './helpers';
import { layoutDocxHeadless, parseDocxHeadless } from './headlessDocxLayout';
import { measureParagraph } from '../../layout-bridge/measuring/measureParagraph';
import {
  measureTextWidth,
  resetCanvasContext,
} from '../../layout-bridge/measuring/measureContainer';
import type { FlowBlock, Measure, ParagraphBlock, ParagraphMeasure, Run, TextRun } from '../types';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** Stub canvas: every glyph advances 0.5em; 11pt = 14.667px → 7.333px/char. */
const CHAR_PX = ((11 * 96) / 72) * 0.5;
const twipsPx = (twips: number) => twips / 15;

function r(text: string, rPr = ''): string {
  const space = /^\s|\s$/.test(text) ? ' xml:space="preserve"' : '';
  return `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t${space}>${text}</w:t></w:r>`;
}
const w = (pct: number) => `<w:w w:val="${pct}"/>`;
const sp = (twips: number) => `<w:spacing w:val="${twips}"/>`;
const p = (runs: string) => `<w:p>${runs}</w:p>`;

// Run patterns of the shape PDF converters emit (synthetic text).
const CUSTOMER_RECORDS = p(
  r('Customer', w(110)) +
    r(' ', sp(6) + w(110)) +
    r('Records', w(110)) +
    r(' ', sp(9) + w(110)) +
    r('Management', sp(-2) + w(110))
);
const VENDOR_DATA = p(
  r('Vendor', w(105)) +
    r(' ', sp(4) + w(110)) +
    r('Data', sp(-2) + w(110)) +
    r(' ', sp(-2) + w(110)) +
    r('AI', sp(2)) +
    r(' ', sp(37)) +
    r('Governance', sp(2))
);
const ACCESS_REVIEW = p(
  r('Access', w(110)) + r(' ', sp(-10) + w(110)) + r('Review', sp(-4) + w(110))
);
const MFA = p(r('Multi-Factor Tokens ', w(105)) + r('(MFT)', sp(-2) + w(105)));
/**
 * The mid-word split, in a cell narrow enough that the line fills up right at
 * "r": the 41 characters before it plus "r" take ~307px of the 330px cell, but
 * "retained?" (66px) does not fit after them, so Word moves the whole word to
 * the next line.
 */
const SPLIT_TEXT = ['How long are the audit log copies', ' ', '(ALCs)', ' r', 'etained?'];
const SPLIT_WORD = p(
  r('How long are the audit log copies', w(100)) +
    r(' ', sp(-6)) +
    r('(ALCs)', sp(-2)) +
    r(' r') +
    r('etained?', sp(-2))
);
const SPLIT_CELL_TWIPS = 330 * 15; // 330px wide cell

const DOCUMENT_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}"><w:body>` +
  CUSTOMER_RECORDS +
  VENDOR_DATA +
  ACCESS_REVIEW +
  MFA +
  `<w:tbl><w:tblPr><w:tblW w:w="${SPLIT_CELL_TWIPS}" w:type="dxa"/><w:tblLayout w:type="fixed"/>` +
  `<w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>` +
  `<w:tblGrid><w:gridCol w:w="${SPLIT_CELL_TWIPS}"/></w:tblGrid>` +
  `<w:tr><w:tc><w:tcPr><w:tcW w:w="${SPLIT_CELL_TWIPS}" w:type="dxa"/></w:tcPr>${SPLIT_WORD}</w:tc></w:tr></w:tbl>` +
  p('') +
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '</w:sectPr></w:body></w:document>';

async function buildDocx(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `</Types>`
  );
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
      `</Relationships>`
  );
  zip.file('word/document.xml', DOCUMENT_XML);
  return new Uint8Array(await zip.generateAsync({ type: 'arraybuffer' }));
}

// ---------------------------------------------------------------------------
// Layout helpers
// ---------------------------------------------------------------------------

type Para = { block: ParagraphBlock; measure: ParagraphMeasure };

/** Every paragraph (body + table cells, recursively) with its measure. */
function paragraphs(blocks: FlowBlock[], measures: Measure[]): Para[] {
  const out: Para[] = [];
  blocks.forEach((block, i) => {
    const measure = measures[i];
    if (!measure) return;
    if (block.kind === 'paragraph' && measure.kind === 'paragraph') {
      out.push({ block, measure });
    } else if (block.kind === 'table' && measure.kind === 'table') {
      block.rows.forEach((row, ri) =>
        row.cells.forEach((cell, ci) => {
          const cm = measure.rows[ri]?.cells[ci];
          if (cm) out.push(...paragraphs(cell.blocks, cm.blocks));
        })
      );
    }
  });
  return out;
}

const textOf = (runs: Run[]) =>
  runs.map((run) => (run.kind === 'text' ? run.text : run.kind === 'tab' ? '\t' : '')).join('');

/** The text of each measured line. */
function lineTexts({ block, measure }: Para): string[] {
  return measure.lines.map((line) => {
    let s = '';
    for (let i = line.fromRun; i <= line.toRun; i++) {
      const run = block.runs[i];
      if (!run) continue;
      if (run.kind === 'text') {
        const from = i === line.fromRun ? line.fromChar : 0;
        const to = i === line.toRun ? line.toChar : run.text.length;
        s += run.text.slice(from, to);
      } else if (run.kind === 'tab') {
        s += '\t';
      }
    }
    return s;
  });
}

const isBreakChar = (c: string | undefined) => c === ' ' || c === '-' || c === '\t';

/**
 * Line boundaries that fall between two word characters — i.e. a word broken
 * mid-word. A word too long for the line legitimately hard-breaks, so those
 * are excluded (the whole word must fit on an empty line to count).
 */
function midWordBreaks(para: Para, lineWidthPx: number): string[] {
  const lines = lineTexts(para);
  const bad: string[] = [];
  for (let i = 0; i + 1 < lines.length; i++) {
    const before = lines[i]!;
    const after = lines[i + 1]!;
    if (!before || !after) continue;
    if (isBreakChar(before[before.length - 1]) || isBreakChar(after[0])) continue;
    const tail = before.split(/[ \t-]/).pop()!;
    const head = after.split(/[ \t-]/)[0]!;
    // Stub metrics: word width ≈ chars × 0.5em of the largest font seen.
    if ((tail.length + head.length) * CHAR_PX * 1.2 < lineWidthPx) bad.push(`${tail}|${head}`);
  }
  return bad;
}

function findPara(paras: Para[], phrase: string): Para {
  const found = paras.find((x) => textOf(x.block.runs).includes(phrase));
  if (!found) throw new Error(`paragraph containing "${phrase}" not found`);
  return found;
}

// ---------------------------------------------------------------------------
// Synthetic fixtures
// ---------------------------------------------------------------------------

describe('cross-run words and spaces measure like Word', () => {
  let restoreDocument: () => void = () => {};
  let paras: Para[];
  beforeAll(async () => {
    restoreDocument = installCanvasDocumentStub();
    resetCanvasContext();
    const doc = await parseDocxHeadless(await buildDocx());
    const res = await layoutDocxHeadless(doc);
    paras = paragraphs(res.blocks, res.measures);
  });
  afterAll(() => {
    resetCanvasContext();
    restoreDocument();
  });

  test('w:w and w:spacing reach the flow runs of space-only runs', () => {
    const runs = findPara(paras, 'Customer').block.runs as TextRun[];
    expect(runs.map((x) => x.text)).toEqual(['Customer', ' ', 'Records', ' ', 'Management']);
    expect(runs.map((x) => x.horizontalScale)).toEqual([110, 110, 110, 110, 110]);
    expect(runs[1]!.letterSpacing).toBeCloseTo(twipsPx(6), 6);
    expect(runs[4]!.letterSpacing).toBeCloseTo(twipsPx(-2), 6);
  });

  test('a scaled run measures at its scaled width; spacing follows every character', () => {
    const { block, measure } = findPara(paras, 'Customer');
    expect(measure.lines).toHaveLength(1);
    // Customer(8) + ' '(1) + Records(7) + ' '(1) + Management(10) at 110%,
    // plus w:spacing after every character (a lone space included).
    const glyphs = (8 + 1 + 7 + 1 + 10) * CHAR_PX * 1.1;
    const tracking = twipsPx(6) * 1 + twipsPx(9) * 1 + twipsPx(-2) * 10;
    expect(measure.lines[0]!.width).toBeCloseTo(glyphs + tracking, 3);

    // A single space with negative spacing keeps a positive advance.
    const space = findPara(paras, 'Access').block.runs[1] as TextRun;
    const spaceWidth = measureTextWidth(space.text, {
      fontSize: 11,
      letterSpacing: space.letterSpacing,
      horizontalScale: space.horizontalScale,
    });
    expect(spaceWidth).toBeCloseTo(CHAR_PX * 1.1 + twipsPx(-10), 6);
    expect(spaceWidth).toBeGreaterThan(0);
    expect(block.runs).toHaveLength(5);
  });

  test('the vendor / MFA paragraphs keep their single-space runs', () => {
    expect(lineTexts(findPara(paras, 'Vendor'))).toEqual(['Vendor Data AI Governance']);
    expect(lineTexts(findPara(paras, 'Multi-Factor'))).toEqual(['Multi-Factor Tokens (MFT)']);
  });

  test('a word split across runs wraps as one word ("r" + "etained?")', () => {
    const para = findPara(paras, 'etained?');
    expect(textOf(para.block.runs)).toBe(SPLIT_TEXT.join(''));
    const lines = lineTexts(para);
    expect(lines).toEqual(['How long are the audit log copies (ALCs) ', 'retained?']);
    expect(midWordBreaks(para, 330)).toEqual([]);
  });
});

describe('measureParagraph: run boundaries are not break opportunities', () => {
  let restoreDocument: () => void = () => {};
  beforeAll(() => {
    restoreDocument = installCanvasDocumentStub();
    resetCanvasContext();
  });
  afterAll(() => {
    resetCanvasContext();
    restoreDocument();
  });

  let pm = 0;
  const text = (t: string, extra: Partial<TextRun> = {}): TextRun => {
    const run: TextRun = { kind: 'text', text: t, pmStart: pm, pmEnd: pm + t.length, ...extra };
    pm += t.length;
    return run;
  };
  const block = (runs: Run[]): ParagraphBlock => ({
    kind: 'paragraph',
    id: 1,
    runs,
    attrs: {},
    pmStart: 0,
    pmEnd: pm + 1,
  });
  const lines = (runs: Run[], width: number) => {
    const b = block(runs);
    return lineTexts({ block: b, measure: measureParagraph(b, width) });
  };

  test('the whole word moves to the next line', () => {
    // "aaaa " = 36.7px, "r" + "etained?" = 66px: 102.7 > 80.
    expect(lines([text('aaaa r'), text('etained?')], 80)).toEqual(['aaaa ', 'retained?']);
  });

  test('a word spanning three runs stays whole', () => {
    expect(lines([text('aaaa b'), text('cd', { bold: true }), text('ef gg')], 80)).toEqual([
      'aaaa ',
      'bcdef gg',
    ]);
  });

  test('a break char at the run boundary still allows the wrap there', () => {
    expect(lines([text('aaaa bb '), text('ccccccc')], 80)).toEqual(['aaaa bb ', 'ccccccc']);
    expect(lines([text('aaaa bb-'), text('ccccccc')], 80)).toEqual(['aaaa bb-', 'ccccccc']);
  });

  test('a tab ends the word', () => {
    const tab: Run = { kind: 'tab', pmStart: 100, pmEnd: 101 } as Run;
    const out = lines([text('aaaa u'), tab, text('b')], 400);
    expect(out).toHaveLength(1);
  });

  test('a word longer than the line still hard-breaks by characters', () => {
    const out = lines([text('aaaaaa'), text('bbbbbbbbbbbb')], 60);
    expect(out.join('')).toBe('aaaaaabbbbbbbbbbbb');
    expect(out.length).toBeGreaterThan(1);
  });
});

// ---------------------------------------------------------------------------
// Optional: a real document with this run shape (not in the repo).
// ---------------------------------------------------------------------------

const FIXTURE = process.env.DOCX_FIXTURE_RUN_SPACING;
const available = !!FIXTURE && fs.existsSync(FIXTURE);
if (!available) {
  console.warn(
    '[cross-run-word-layout] real-file check SKIPPED — set DOCX_FIXTURE_RUN_SPACING to the path ' +
      'of a PDF-converted .docx with per-word runs to run it.'
  );
}
const runFixture = available ? test : test.skip;

describe('real PDF-converted questionnaire (fixture) lays out like Word', () => {
  let restoreDocument: () => void = () => {};
  beforeAll(() => {
    if (available) {
      restoreDocument = installCanvasDocumentStub();
      resetCanvasContext();
    }
  });
  afterAll(() => {
    resetCanvasContext();
    restoreDocument();
  });

  runFixture(
    'space-only runs keep a width and no word breaks mid-word',
    async () => {
      const doc = await parseDocxHeadless(new Uint8Array(fs.readFileSync(FIXTURE!)));
      const res = await layoutDocxHeadless(doc);
      const paras = paragraphs(res.blocks, res.measures);
      expect(paras.length).toBeGreaterThan(0);

      for (const para of paras) {
        const runs = para.block.runs.filter((x): x is TextRun => x.kind === 'text');
        // Every scaled run is measured at its scaled width, so the next run
        // (the space) starts after the scaled glyphs, not under them.
        for (const run of runs) {
          const width = measureTextWidth(run.text, {
            fontFamily: run.fontFamily,
            fontSize: run.fontSize,
            letterSpacing: run.letterSpacing,
            horizontalScale: run.horizontalScale,
          });
          if (run.text !== '' && run.text.trim() === '') expect(width).toBeGreaterThan(0);
        }
      }

      // Whole document: no line boundary falls between two runs of one word.
      const broken = paras.flatMap((para) => {
        const width = para.measure.lines.reduce((m, l) => Math.max(m, l.width), 0);
        return midWordBreaks(para, Math.max(width, 1));
      });
      expect(broken).toEqual([]);
    },
    60_000
  );
});

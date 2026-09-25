/**
 * Create the three synthetic DOCX fixtures for repeating-table-header
 * pagination (`w:trPr/w:tblHeader`). Content is generic sample text only.
 *
 * Word's rules these fixtures exercise:
 *  - a header row never sits alone at the bottom of a page (and is never
 *    split across the boundary); if header + first body row does not fit,
 *    the whole table moves to the next page;
 *  - a repeated header on a continuation page renders at full height with
 *    the first body row beneath it.
 *
 * Two corpora, because the editor serves both:
 *  - RFP / questionnaire: wide 7-column matrices with many short body rows;
 *  - proposal: a 2-column table with a repeating header inside long prose.
 *
 * Run: bun scripts/create-table-header-pagination-fixtures.mjs
 */

import JSZip from 'jszip';
import * as fs from 'node:fs';
import * as path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const FIXTURES = path.join(ROOT, 'e2e/fixtures');
const ZIP_DATE = new Date('2026-01-01T00:00:00Z');

const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>`;

const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
</Relationships>`;

const documentRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
    <w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr>
  </w:style>
</w:styles>`;

function coreXml(title) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties
  xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:dcterms="http://purl.org/dc/terms/"
  xmlns:dcmitype="http://purl.org/dc/dcmitype/"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>${title}</dc:title>
  <dc:creator>docx-editor fixture generator</dc:creator>
  <cp:lastModifiedBy>docx-editor fixture generator</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:modified>
</cp:coreProperties>`;
}

const SECT_PR = `<w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
      <w:cols w:space="720"/>
    </w:sectPr>`;

function p(text, options = {}) {
  const bold = options.bold ? '<w:b/>' : '';
  const size = options.size ?? 22;
  return `<w:p>
    <w:pPr><w:spacing w:after="${options.after ?? 0}" w:line="240" w:lineRule="auto"/></w:pPr>
    <w:r><w:rPr>${bold}<w:sz w:val="${size}"/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>
  </w:p>`;
}

function tc(text, width, options = {}) {
  return `<w:tc>
    <w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${
      options.shaded ? '<w:shd w:val="clear" w:color="auto" w:fill="EDEDED"/>' : ''
    }</w:tcPr>
    ${p(text, { after: 0, bold: options.bold })}
  </w:tc>`;
}

function tr(cellTexts, width, options = {}) {
  const trPr = [
    options.header ? '<w:tblHeader/>' : '',
    options.cantSplit ? '<w:cantSplit/>' : '',
  ].join('');
  return `<w:tr>
    ${trPr ? `<w:trPr>${trPr}</w:trPr>` : ''}
    ${cellTexts.map((t) => tc(t, width, { bold: options.header, shaded: options.header })).join('\n')}
  </w:tr>`;
}

function tbl(columnCount, rows, width = 9360) {
  const colWidth = Math.floor(width / columnCount);
  return `<w:tbl>
      <w:tblPr>
        <w:tblW w:w="${width}" w:type="dxa"/>
        <w:tblBorders>
          <w:top w:val="single" w:sz="8" w:space="0" w:color="666666"/>
          <w:left w:val="single" w:sz="8" w:space="0" w:color="666666"/>
          <w:bottom w:val="single" w:sz="8" w:space="0" w:color="666666"/>
          <w:right w:val="single" w:sz="8" w:space="0" w:color="666666"/>
          <w:insideH w:val="single" w:sz="8" w:space="0" w:color="999999"/>
          <w:insideV w:val="single" w:sz="8" w:space="0" w:color="999999"/>
        </w:tblBorders>
      </w:tblPr>
      <w:tblGrid>${Array.from({ length: columnCount }, () => `<w:gridCol w:w="${colWidth}"/>`).join('')}</w:tblGrid>
      ${rows.map((row) => tr(row.cells, colWidth, row)).join('\n')}
    </w:tbl>`;
}

/**
 * Filler prose, one single line per paragraph, so the count tunes the table's
 * start position to the line. Sized (see the call sites) so the table's header
 * row lands in the last line of space on the page — the case Word resolves by
 * moving the whole table to the next page.
 */
function fillerParagraphs(count, label) {
  return Array.from({ length: count }, (_, i) => p(`${label} note ${i + 1}.`)).join('\n');
}

const FILLER_RFP = Number(process.env.FILLER_RFP ?? 44);
const FILLER_PROPOSAL = Number(process.env.FILLER_PROPOSAL ?? 44);

const SENTENCE =
  'Sample response text describing the capability, its scope, and the supporting evidence.';

/** 15 body rows of 2-4 lines each. */
function bodyRows(columnCount, cantSplit) {
  return Array.from({ length: 15 }, (_, i) => ({
    cantSplit,
    cells: Array.from(
      { length: columnCount },
      (_, c) => `Row ${i + 1} / col ${c + 1}. ${SENTENCE.repeat(1 + (i % 3))}`
    ),
  }));
}

function headerRow(columnCount, labels) {
  return {
    header: true,
    cells: Array.from(
      { length: columnCount },
      (_, c) => labels[c] ?? `Evaluation criterion column ${c + 1}`
    ),
  };
}

async function write(filename, title, bodyXml) {
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    ${bodyXml}
    ${SECT_PR}
  </w:body>
</w:document>`;

  const zip = new JSZip();
  const opts = { date: ZIP_DATE, createFolders: false };
  zip.file('[Content_Types].xml', contentTypesXml, opts);
  zip.file('_rels/.rels', relsXml, opts);
  zip.file('word/_rels/document.xml.rels', documentRelsXml, opts);
  zip.file('word/document.xml', documentXml, opts);
  zip.file('word/styles.xml', stylesXml, opts);
  zip.file('docProps/core.xml', coreXml(title), opts);

  const buffer = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 9 },
  });
  const out = path.join(FIXTURES, filename);
  fs.writeFileSync(out, buffer);
  console.log(`Created ${out}`);
}

// 1. RFP / questionnaire: 7-column table with a repeating header, rows may split.
await write(
  'table-header-orphan-rfp.docx',
  'Repeating Table Header Orphan — RFP Matrix Fixture',
  [
    p('Requirements Matrix Fixture', { bold: true, size: 32, after: 240 }),
    fillerParagraphs(FILLER_RFP, 'Requirements overview'),
    tbl(7, [
      headerRow(7, [
        'Requirement ID',
        'Requirement description',
        'Priority level',
        'Vendor response',
        'Supporting evidence',
        'Owner',
        'Notes',
      ]),
      ...bodyRows(7, false),
    ]),
    p('Closing generated paragraph after the requirements matrix.'),
  ].join('\n')
);

// 2. Same matrix with w:cantSplit on every row.
await write(
  'table-header-orphan-cantsplit.docx',
  'Repeating Table Header Orphan — cantSplit Fixture',
  [
    p('Requirements Matrix Fixture (rows may not split)', { bold: true, size: 32, after: 240 }),
    fillerParagraphs(FILLER_RFP, 'Requirements overview'),
    tbl(7, [
      {
        ...headerRow(7, [
          'Requirement ID',
          'Requirement description',
          'Priority level',
          'Vendor response',
          'Supporting evidence',
          'Owner',
          'Notes',
        ]),
        cantSplit: true,
      },
      ...bodyRows(7, true),
    ]),
    p('Closing generated paragraph after the requirements matrix.'),
  ].join('\n')
);

// 3. Proposal-style: 2-column table with a repeating header inside long prose.
await write(
  'table-header-orphan-proposal.docx',
  'Repeating Table Header Orphan — Proposal Fixture',
  [
    p('Solution Proposal Fixture', { bold: true, size: 32, after: 240 }),
    fillerParagraphs(FILLER_PROPOSAL, 'Executive summary'),
    tbl(2, [
      headerRow(2, [
        'Capability area and evaluation criterion covered by this response',
        'How the proposed solution delivers the capability, including scope and evidence',
      ]),
      ...bodyRows(2, false),
    ]),
    fillerParagraphs(12, 'Implementation approach'),
  ].join('\n')
);

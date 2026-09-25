/**
 * Create the DOCX fixtures for a table BODY ROW that splits across a page
 * boundary (Word's "allow row to break across pages" — no `w:cantSplit`).
 * Content is generic sample text only.
 *
 * Word's rules these fixtures exercise:
 *  - the remainder of a split row resumes at the very top of the printable
 *    area on the next page — nothing of the table is painted inside the top
 *    margin, where the running header lives;
 *  - a `w:tblHeader` row repeats above that remainder, because it repeats on
 *    "each new page on which part of this table is displayed" (ECMA-376
 *    §17.4.78) and a page showing the tail of a split row displays part of
 *    the table;
 *  - the split loses and duplicates nothing: the continuation starts at the
 *    line where the previous page stopped.
 *
 * Two corpora, because the editor serves both:
 *  - RFP / questionnaire: a wide matrix with a repeating header whose first
 *    body row is far taller than a page;
 *  - proposal: a headerless 2-column table with a tall row inside long prose.
 *
 * Run: bun scripts/create-table-row-split-fixtures.mjs
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

function tc(paragraphs, width, options = {}) {
  return `<w:tc>
    <w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${
      options.shaded ? '<w:shd w:val="clear" w:color="auto" w:fill="EDEDED"/>' : ''
    }</w:tcPr>
    ${paragraphs.map((text) => p(text, { after: 0, bold: options.bold })).join('\n')}
  </w:tc>`;
}

/** A row's cells are arrays of paragraphs, so a cell can be paragraphs tall. */
function tr(cells, width, options = {}) {
  const trPr = options.header ? '<w:trPr><w:tblHeader/></w:trPr>' : '';
  return `<w:tr>
    ${trPr}
    ${cells
      .map((paragraphs) => tc(paragraphs, width, { bold: options.header, shaded: options.header }))
      .join('\n')}
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

const SENTENCE =
  'Sample response text describing the capability, its scope, and the supporting evidence for the evaluation committee.';

/** A cell of `count` numbered paragraphs — tall enough to force a row split. */
function tallCell(label, count) {
  return Array.from({ length: count }, (_, i) => `${label} paragraph ${i + 1}. ${SENTENCE}`);
}

/** Prose filler, one paragraph per line, to tune where the split lands. */
function fillerParagraphs(count, label) {
  return Array.from({ length: count }, (_, i) => p(`${label} note ${i + 1}.`)).join('\n');
}

const FILLER_SPLIT = Number(process.env.FILLER_SPLIT ?? 12);

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

// 1. RFP / questionnaire: repeating header + a body row taller than a page, so
//    the row splits and the continuation page must repeat the header above it.
await write(
  'table-row-split-header.docx',
  'Split Table Row — Repeating Header Fixture',
  [
    p('Requirements Matrix — Split Row Fixture', { bold: true, size: 32, after: 240 }),
    fillerParagraphs(FILLER_SPLIT, 'Requirements overview'),
    tbl(4, [
      {
        header: true,
        cells: [
          ['Requirement ID'],
          ['Requirement description'],
          ['Priority level'],
          ['Vendor response'],
        ],
      },
      {
        cells: [
          ['PS-01'],
          tallCell('Requirement detail', 26),
          ['Must-Have'],
          tallCell('Vendor response', 26),
        ],
      },
      {
        cells: [['PS-02'], ['Short follow-up requirement.'], ['Nice-to-Have'], ['Supported.']],
      },
    ]),
    p('Closing generated paragraph after the requirements matrix.'),
  ].join('\n')
);

// 2. Proposal: headerless 2-column table with a tall row inside narrative prose.
await write(
  'table-row-split-noheader.docx',
  'Split Table Row — Headerless Proposal Fixture',
  [
    p('Implementation Approach — Split Row Fixture', { bold: true, size: 32, after: 240 }),
    fillerParagraphs(FILLER_SPLIT, 'Approach narrative'),
    tbl(2, [
      { cells: [['Phase 1 — Discovery'], ['Two week discovery and alignment workshop.']] },
      { cells: [['Phase 2 — Delivery'], tallCell('Delivery detail', 30)] },
      { cells: [['Phase 3 — Hypercare'], ['Thirty days of post go-live support.']] },
    ]),
    fillerParagraphs(6, 'Approach narrative continued'),
  ].join('\n')
);

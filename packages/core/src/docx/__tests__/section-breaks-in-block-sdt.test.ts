import { describe, expect, test } from 'bun:test';
import { parseDocumentBody } from '../documentParser';

/**
 * A `w:p` that carries `w:pPr/w:sectPr` ends a section wherever it sits —
 * including inside a block-level `w:sdt`, which ECMA-376 treats as a
 * transparent container. `toFlowBlocks` already descends into block SDTs when
 * it emits `sectionBreak` blocks, so when `buildSections` did not, the two
 * disagreed: the paginator numbered pages from the flow's section breaks while
 * `sections` was short, and `resolveSectionHeaderFooters` clamped the overflow
 * to the LAST section — painting the closing section's full-page cover artwork
 * over ordinary body pages.
 */

const sectPr = (headerRid: string) => `
      <w:pPr>
        <w:sectPr>
          <w:headerReference w:type="default" r:id="${headerRid}"/>
          <w:type w:val="nextPage"/>
          <w:pgSz w:w="12240" w:h="15840"/>
        </w:sectPr>
      </w:pPr>`;

const doc = (body: string) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>${body}</w:body>
</w:document>`;

const parse = (xml: string) => parseDocumentBody(xml, null, null, null, null, null);

describe('section breaks inside a block SDT', () => {
  test('a sectPr paragraph nested in w:sdt still ends a section', () => {
    const body = parse(
      doc(`
    <w:p><w:r><w:t>cover</w:t></w:r>${sectPr('rId1')}</w:p>
    <w:sdt>
      <w:sdtPr><w:id w:val="1"/></w:sdtPr>
      <w:sdtContent>
        <w:p><w:r><w:t>body</w:t></w:r></w:p>
        <w:p>${sectPr('rId2')}</w:p>
        <w:p><w:r><w:t>back cover</w:t></w:r></w:p>
      </w:sdtContent>
    </w:sdt>
    <w:sectPr>
      <w:headerReference w:type="default" r:id="rId3"/>
      <w:pgSz w:w="12240" w:h="15840"/>
    </w:sectPr>`)
    );

    // 2 inline sectPr + the body-level one = 3 sections.
    expect(body.sections?.length).toBe(3);
    expect(body.sections?.map((s) => s.properties.headerReferences?.[0]?.rId)).toEqual([
      'rId1',
      'rId2',
      'rId3',
    ]);
  });

  test('the body-level sectPr always closes a final section', () => {
    // The last inline break sits inside the SDT with no top-level content after
    // it, so the "remaining content" heuristic dropped the final section.
    const body = parse(
      doc(`
    <w:sdt>
      <w:sdtPr><w:id w:val="1"/></w:sdtPr>
      <w:sdtContent>
        <w:p><w:r><w:t>body</w:t></w:r></w:p>
        <w:p>${sectPr('rId1')}</w:p>
      </w:sdtContent>
    </w:sdt>
    <w:sectPr>
      <w:headerReference w:type="default" r:id="rId9"/>
      <w:pgSz w:w="12240" w:h="15840"/>
    </w:sectPr>`)
    );

    expect(body.sections?.length).toBe(2);
    expect(body.sections?.[1].properties.headerReferences?.[0]?.rId).toBe('rId9');
    // `finalSectionProperties` is re-pointed at the last section so header
    // resolution reads the inheritance-resolved copy, not a stale one.
    expect(body.finalSectionProperties).toBe(body.sections?.[1].properties);
  });

  test('a document with no section breaks still has exactly one section', () => {
    const body = parse(
      doc(`
    <w:p><w:r><w:t>only</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`)
    );
    expect(body.sections?.length).toBe(1);
  });
});

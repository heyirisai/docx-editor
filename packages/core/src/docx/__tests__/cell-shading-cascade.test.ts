/**
 * Cell shading follows the OOXML cascade: a cell's own `w:shd` wins over the
 * table style's conditional (firstRow / band) shading, and `w:shd` is read as
 * the two-colour pattern it is (§17.3.5).
 *
 * Two real-template failures motivated this:
 *  - A one-row table styled `firstRow` navy, whose cells declare
 *    `w:shd w:fill="auto"` ("no shading"), painted navy over its own text.
 *    The resolver compared `shading.fill` rather than the shading itself, so
 *    an explicit "none" read as absent and fell through to the style.
 *  - Header rows shaded `w:val="solid" w:color="2E5090" w:fill="auto"`
 *    painted nothing, taking their white header text with them.
 */

import { describe, expect, test } from 'bun:test';
import { parseDocumentBody } from '../documentParser';
import { parseStyles } from '../styleParser';
import { toProseDoc } from '../../prosemirror/conversion/toProseDoc';
import type { Document } from '../../types/document';

const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="table" w:customStyle="1" w:styleId="NavyHeader">
    <w:name w:val="Navy Header"/>
    <w:tblStylePr w:type="firstRow">
      <w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="001B49"/></w:tcPr>
    </w:tblStylePr>
  </w:style>
</w:styles>`;

function docXml(cellShd: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tblPr>
        <w:tblStyle w:val="NavyHeader"/>
        <w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1"
          w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>
      </w:tblPr>
      <w:tr><w:tc><w:tcPr>${cellShd}</w:tcPr>
        <w:p><w:r><w:t>Row one</w:t></w:r></w:p>
      </w:tc></w:tr>
    </w:tbl>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
  </w:body>
</w:document>`;
}

/** backgroundColor of the table's single cell, as the painter would read it. */
function cellBackground(cellShd: string): string | null {
  const styles = parseStyles(STYLES_XML, null);
  const body = parseDocumentBody(docXml(cellShd), styles);
  // `parseStyles` yields a Map; the PM converter wants StyleDefinitions.
  const styleDefinitions = { styles: [...styles.values()] };
  const doc = { package: { document: body, styles: styleDefinitions } } as unknown as Document;
  const pm = toProseDoc(doc, { styles: styleDefinitions });
  let bg: string | null | undefined;
  pm.descendants((node) => {
    if ((node.type.name === 'tableCell' || node.type.name === 'tableHeader') && bg === undefined) {
      bg = (node.attrs as { backgroundColor?: string | null }).backgroundColor ?? null;
    }
    return true;
  });
  return bg ?? null;
}

describe('cell shading cascade', () => {
  test("the table style's firstRow shading applies when the cell declares none", () => {
    expect(cellBackground('')).toBe('001B49');
  });

  test('a cell’s explicit "no shading" beats the style’s conditional shading', () => {
    // `w:fill="auto"` is an explicit automatic (transparent) background, not
    // an absent value — Word paints no navy here.
    expect(cellBackground('<w:shd w:val="clear" w:color="auto" w:fill="auto"/>')).toBe(null);
  });

  test('a solid pattern paints its colour, not its fill', () => {
    expect(cellBackground('<w:shd w:val="solid" w:color="2E5090" w:fill="auto"/>')).toBe('2E5090');
  });

  test('a cell fill still wins over the style', () => {
    expect(cellBackground('<w:shd w:val="clear" w:color="auto" w:fill="FFE599"/>')).toBe('FFE599');
  });
});

/**
 * Two things a `wps:wsp` carries that the model used to drop:
 *
 *  - TABLES inside `w:txbxContent`. The element is `EG_BlockLevelElts`, so a
 *    text box can hold a laid-out panel; the parser skipped tables and the
 *    box painted as an empty frame (the Iris proposal template's "PROOF
 *    POINT" block).
 *
 *  - The `a:prstGeom` PRESET. Every shape was drawn as a rectangle, so an
 *    `ellipse` badge (COMET's "15 years of experience" roundel) and a
 *    `roundRect` pill button (the Iris back cover's "EXPLORE OUR PLATFORM")
 *    both came out hard-cornered.
 */

import { describe, expect, test } from 'bun:test';
import { parseDocumentBody } from '../documentParser';
import { serializeDocumentBody } from '../serializer/documentSerializer';
import type { Paragraph, Run, Shape, ShapeContent, Table } from '../../types/document';

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"';

function documentWith(drawing: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document ${NS}>
      <w:body>
        <w:p><w:r>${drawing}</w:r></w:p>
        <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
      </w:body>
    </w:document>`;
}

function shapesIn(xml: string): Shape[] {
  const out: Shape[] = [];
  for (const block of parseDocumentBody(xml).content) {
    if (block.type !== 'paragraph') continue;
    for (const item of (block as Paragraph).content) {
      if (item.type !== 'run') continue;
      for (const rc of (item as Run).content) {
        if (rc.type === 'shape') out.push((rc as ShapeContent).shape);
      }
    }
  }
  return out;
}

function textBoxDrawing(geometry: string, body: string): string {
  return `<w:drawing>
    <wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0"
               relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
      <wp:simplePos x="0" y="0"/>
      <wp:positionH relativeFrom="margin"><wp:posOffset>0</wp:posOffset></wp:positionH>
      <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
      <wp:extent cx="2569464" cy="420624"/>
      <wp:wrapSquare wrapText="bothSides"/>
      <wp:docPr id="7" name="Box 7"/>
      <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
        <wps:wsp>
          <wps:cNvSpPr txBox="1"/>
          <wps:spPr>
            <a:xfrm><a:off x="0" y="0"/><a:ext cx="2569464" cy="420624"/></a:xfrm>
            ${geometry}
            <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
          </wps:spPr>
          <wps:txbx><w:txbxContent>${body}</w:txbxContent></wps:txbx>
          <wps:bodyPr/>
        </wps:wsp>
      </a:graphicData></a:graphic>
    </wp:anchor>
  </w:drawing>`;
}

const RECT = '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>';

const PANEL_TABLE =
  '<w:tbl><w:tblPr><w:tblW w:w="9628" w:type="dxa"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="2019"/><w:gridCol w:w="7609"/></w:tblGrid>' +
  '<w:tr>' +
  '<w:tc><w:tcPr><w:tcW w:w="2019" w:type="dxa"/></w:tcPr>' +
  '<w:p><w:r><w:t>PROOF</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:tcPr><w:tcW w:w="7609" w:type="dxa"/></w:tcPr>' +
  '<w:p><w:r><w:t>&lt;&lt;intro&gt;&gt;</w:t></w:r></w:p></w:tc>' +
  '</w:tr></w:tbl>';

describe('a table inside w:txbxContent', () => {
  test('parses as table content, not a dropped block', () => {
    const [shape] = shapesIn(documentWith(textBoxDrawing(RECT, PANEL_TABLE)));
    const content = shape?.textBody?.content ?? [];
    expect(content.map((c) => c.type)).toEqual(['table']);
    const table = content[0] as Table;
    expect(table.rows).toHaveLength(1);
    expect(table.rows[0].cells).toHaveLength(2);
  });

  test('keeps paragraphs and tables in source order', () => {
    const body = `<w:p><w:r><w:t>Before</w:t></w:r></w:p>${PANEL_TABLE}<w:p><w:r><w:t>After</w:t></w:r></w:p>`;
    const [shape] = shapesIn(documentWith(textBoxDrawing(RECT, body)));
    expect((shape?.textBody?.content ?? []).map((c) => c.type)).toEqual([
      'paragraph',
      'table',
      'paragraph',
    ]);
  });

  test('round-trips the table back into the text box', () => {
    const parsed = parseDocumentBody(documentWith(textBoxDrawing(RECT, PANEL_TABLE)));
    const xml = serializeDocumentBody(parsed);
    expect(xml).toContain('<w:txbxContent>');
    expect(xml).toContain('<w:tbl>');
    expect(xml).toContain('<w:t>PROOF</w:t>');
    // The table must land INSIDE the box, not after it.
    expect(xml.indexOf('<w:tbl>')).toBeGreaterThan(xml.indexOf('<w:txbxContent>'));
    expect(xml.indexOf('<w:tbl>')).toBeLessThan(xml.indexOf('</w:txbxContent>'));
  });
});

describe('a:prstGeom preset geometry', () => {
  const para = '<w:p><w:r><w:t>Label</w:t></w:r></w:p>';

  test('ellipse is modelled and written back', () => {
    const geom = '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>';
    const parsed = parseDocumentBody(documentWith(textBoxDrawing(geom, para)));
    const [shape] = shapesIn(documentWith(textBoxDrawing(geom, para)));
    expect(shape?.geometry).toBe('ellipse');
    expect(shape?.cornerAdj).toBeUndefined();
    expect(serializeDocumentBody(parsed)).toContain('<a:prstGeom prst="ellipse">');
  });

  test('a flow-chart connector is an ellipse too', () => {
    const geom = '<a:prstGeom prst="flowChartConnector"><a:avLst/></a:prstGeom>';
    const [shape] = shapesIn(documentWith(textBoxDrawing(geom, para)));
    expect(shape?.geometry).toBe('ellipse');
  });

  test('roundRect carries its adjust as a fraction of the short side', () => {
    const geom =
      '<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 50000"/></a:avLst></a:prstGeom>';
    const parsed = parseDocumentBody(documentWith(textBoxDrawing(geom, para)));
    const [shape] = shapesIn(documentWith(textBoxDrawing(geom, para)));
    expect(shape?.geometry).toBe('roundRect');
    expect(shape?.cornerAdj).toBe(0.5);
    expect(serializeDocumentBody(parsed)).toContain('<a:gd name="adj" fmla="val 50000"/>');
  });

  test('roundRect with no avLst falls back to Word’s default adjust', () => {
    const geom = '<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>';
    const [shape] = shapesIn(documentWith(textBoxDrawing(geom, para)));
    expect(shape?.cornerAdj).toBeCloseTo(0.16667, 5);
  });

  test('a file-supplied adjust is clamped into the legal range', () => {
    const geom =
      '<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 9999999"/></a:avLst></a:prstGeom>';
    const [shape] = shapesIn(documentWith(textBoxDrawing(geom, para)));
    expect(shape?.cornerAdj).toBe(0.5);
  });

  test('a plain rect stays square', () => {
    const [shape] = shapesIn(documentWith(textBoxDrawing(RECT, para)));
    expect(shape?.geometry).toBeUndefined();
  });
});

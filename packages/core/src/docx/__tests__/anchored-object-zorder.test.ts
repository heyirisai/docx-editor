/**
 * Regression — OOXML `wp:anchor relativeHeight` is the z-order among
 * overlapping anchored objects and must survive parse so render can stack
 * them the way Word does. A cover built as banner image (low
 * relativeHeight) + title text boxes (higher) was painting the banner OVER
 * the titles: the front floating-image layer used a flat z-index of 10
 * while anchored text boxes rendered at z-index 1.
 */

import { describe, expect, test } from 'bun:test';
import { parseDocumentBody } from '../documentParser';

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"';

function anchoredTextBoxXml(relativeHeight: number, text: string): string {
  return `<w:r>
    <mc:AlternateContent>
      <mc:Choice Requires="wps">
        <w:drawing>
          <wp:anchor distT="0" distB="0" distL="114300" distR="114300"
            simplePos="0" relativeHeight="${relativeHeight}" behindDoc="0" locked="0"
            layoutInCell="1" allowOverlap="1">
            <wp:simplePos x="0" y="0"/>
            <wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>
            <wp:positionV relativeFrom="paragraph"><wp:posOffset>100330</wp:posOffset></wp:positionV>
            <wp:extent cx="1390650" cy="278130"/>
            <wp:effectExtent l="0" t="0" r="0" b="0"/>
            <wp:wrapSquare wrapText="bothSides"/>
            <wp:docPr id="7" name="Title"/>
            <wp:cNvGraphicFramePr/>
            <a:graphic>
              <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                <wps:wsp>
                  <wps:cNvSpPr txBox="1"/>
                  <wps:spPr>
                    <a:xfrm><a:off x="0" y="0"/><a:ext cx="1390650" cy="278130"/></a:xfrm>
                    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                  </wps:spPr>
                  <wps:txbx>
                    <w:txbxContent>
                      <w:p><w:r><w:t>${text}</w:t></w:r></w:p>
                    </w:txbxContent>
                  </wps:txbx>
                  <wps:bodyPr/>
                </wps:wsp>
              </a:graphicData>
            </a:graphic>
          </wp:anchor>
        </w:drawing>
      </mc:Choice>
      <mc:Fallback><w:pict/></mc:Fallback>
    </mc:AlternateContent>
  </w:r>`;
}

const DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <w:document ${NS}>
    <w:body>
      <w:p>
        ${anchoredTextBoxXml(251847168, 'COMPANY NAME and memoryBlue')}
        ${anchoredTextBoxXml(251853312, 'Pipeline and Revenue Acceleration Proposal')}
      </w:p>
      <w:sectPr/>
    </w:body>
  </w:document>`;

describe('anchored object z-order (relativeHeight)', () => {
  test('anchored text box shapes carry their relativeHeight through parse', () => {
    const body = parseDocumentBody(DOCUMENT, null, null, null);
    const shapes: Array<{ relativeHeight?: number; text: string }> = [];
    for (const block of body.content) {
      if (block.type !== 'paragraph') continue;
      for (const content of block.content) {
        if (content.type !== 'run') continue;
        for (const rc of content.content) {
          if (rc.type === 'shape' && 'shape' in rc) {
            const shape = rc.shape as { relativeHeight?: number; textBody?: unknown };
            shapes.push({
              relativeHeight: shape.relativeHeight,
              text: JSON.stringify(shape.textBody ?? '').slice(0, 60),
            });
          }
        }
      }
    }
    expect(shapes.length).toBe(2);
    expect(shapes[0]?.relativeHeight).toBe(251847168);
    expect(shapes[1]?.relativeHeight).toBe(251853312);
  });
});

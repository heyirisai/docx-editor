/**
 * Two shapes Word draws that the model used to lose or mis-draw:
 *
 *  - a stroke-only CONNECTOR (`wps:cNvCnPr` / `a:prstGeom prst="line"`),
 *    which every one of these templates uses as the rule above a footer.
 *    Lifted as a filled shape it was outlined as a rectangle, so a `cy="0"`
 *    hairline painted a full-width box across the bottom of every page.
 *
 *  - shapes inside a GROUP (`wpg:wgp`). `isTextBoxDrawing` /
 *    `isFilledShapeDrawing` only look one level down, so a grouped callout
 *    panel — filled rect + accent bar + text box — reached no renderer and
 *    the body text that should wrap beside it ran full width.
 */

import { describe, expect, test } from 'bun:test';
import { parseDocumentBody } from '../documentParser';
import { serializeDocumentBody } from '../serializer/documentSerializer';
import type { Paragraph, Run, Shape, ShapeContent } from '../../types/document';

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"';

function documentWith(drawing: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document ${NS}>
      <w:body>
        <w:p>
          <w:r>
            <mc:AlternateContent>
              <mc:Choice Requires="wps">${drawing}</mc:Choice>
              <mc:Fallback><w:pict/></mc:Fallback>
            </mc:AlternateContent>
          </w:r>
        </w:p>
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

/** A footer rule: zero height, stroke only, no text body. */
const CONNECTOR = `<w:drawing>
  <wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="251658266"
    behindDoc="0" locked="1" layoutInCell="1" allowOverlap="1">
    <wp:simplePos x="0" y="0"/>
    <wp:positionH relativeFrom="page"><wp:align>center</wp:align></wp:positionH>
    <wp:positionV relativeFrom="page"><wp:posOffset>9464040</wp:posOffset></wp:positionV>
    <wp:extent cx="7772400" cy="0"/>
    <wp:wrapNone/>
    <wp:docPr id="1" name="Straight Connector 1"/>
    <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
      <wps:wsp>
        <wps:cNvCnPr/>
        <wps:spPr>
          <a:xfrm><a:off x="0" y="0"/><a:ext cx="7772400" cy="0"/></a:xfrm>
          <a:prstGeom prst="line"><a:avLst/></a:prstGeom>
          <a:ln w="12700"><a:solidFill><a:srgbClr val="5D564B"/></a:solidFill></a:ln>
        </wps:spPr>
        <wps:bodyPr/>
      </wps:wsp>
    </a:graphicData></a:graphic>
  </wp:anchor>
</w:drawing>`;

/** A grouped callout: white panel with text, plus an accent bar over it. */
const GROUP = `<w:drawing>
  <wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0"
    relativeHeight="251658248" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
    <wp:simplePos x="0" y="0"/>
    <wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH>
    <wp:positionV relativeFrom="margin"><wp:posOffset>1506855</wp:posOffset></wp:positionV>
    <wp:extent cx="2000000" cy="4000000"/>
    <wp:wrapSquare wrapText="bothSides"/>
    <wp:docPr id="211" name="Group 211"/>
    <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">
      <wpg:wgp>
        <wpg:grpSpPr>
          <a:xfrm>
            <a:off x="0" y="0"/><a:ext cx="2000000" cy="4000000"/>
            <a:chOff x="0" y="0"/><a:chExt cx="4000000" cy="8000000"/>
          </a:xfrm>
        </wpg:grpSpPr>
        <wps:wsp>
          <wps:cNvSpPr/>
          <wps:spPr>
            <a:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="8000000"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
            <a:ln w="15875"><a:solidFill><a:srgbClr val="333333"/></a:solidFill></a:ln>
          </wps:spPr>
          <wps:txbx><w:txbxContent>
            <w:p><w:r><w:t>Advantage</w:t></w:r></w:p>
          </w:txbxContent></wps:txbx>
          <wps:bodyPr lIns="182880" tIns="457200" rIns="182880" bIns="73152"/>
        </wps:wsp>
        <wps:wsp>
          <wps:cNvSpPr/>
          <wps:spPr>
            <a:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="800000"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            <a:solidFill><a:srgbClr val="3333CC"/></a:solidFill>
          </wps:spPr>
          <wps:bodyPr/>
        </wps:wsp>
      </wpg:wgp>
    </a:graphicData></a:graphic>
  </wp:anchor>
</w:drawing>`;

describe('stroke-only connectors', () => {
  test('reach the model flagged as a line, not as an outlined box', () => {
    const shapes = shapesIn(documentWith(CONNECTOR));
    expect(shapes.length).toBe(1);
    expect(shapes[0].lineShape).toBe('down');
    expect(shapes[0].renderOnly).toBe(true);
    // The declared extent is kept verbatim — a horizontal rule is 0 high.
    expect(shapes[0].size).toEqual({ width: 7772400, height: 0 });
    expect(shapes[0].outline?.color?.rgb).toBe('5D564B');
  });
});

describe('shapes inside a wpg:wgp group', () => {
  test('every child shape reaches the model, mapped into page space', () => {
    const shapes = shapesIn(documentWith(GROUP));
    expect(shapes.length).toBe(2);
    // chExt is 2x ext on both axes, so each child halves.
    expect(shapes[0].size).toEqual({ width: 2000000, height: 4000000 });
    expect(shapes[1].size).toEqual({ width: 2000000, height: 400000 });
    // Group anchor + wrap carry to each child so the body wraps beside them.
    for (const shape of shapes) {
      expect(shape.renderOnly).toBe(true);
      expect(shape.wrap?.type).toBe('square');
      expect(shape.position?.horizontal?.alignment).toBe('right');
      expect(shape.position?.vertical?.posOffset).toBe(1506855);
    }
    // Text inside the group's text box is parsed, not dropped.
    expect(shapes[0].textBody?.content.length).toBe(1);
    expect(shapes[1].fill?.color?.rgb).toBe('3333CC');
  });

  test('saving writes the preserved group once, with no modelled duplicate', () => {
    const xml = serializeDocumentBody(parseDocumentBody(documentWith(GROUP)));
    expect((xml.match(/<wpg:wgp>/g) ?? []).length).toBe(1);
    expect((xml.match(/<wps:wsp>/g) ?? []).length).toBe(2);
  });
});

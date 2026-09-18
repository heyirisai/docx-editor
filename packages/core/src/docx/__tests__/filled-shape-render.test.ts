/**
 * Decorative filled shapes — a `wps:wsp` with a solid fill and no `wps:txbx`
 * — are how modern templates paint their colour blocks: the navy band behind
 * a section, the accent bar beside a heading. Nothing modelled them:
 * `parseImage` returns null for a shape, `isTextBoxDrawing` requires a text
 * box, and the layout engine has no shape block, so they vanished from the
 * page while the white text laid over them stayed white-on-white.
 *
 * They now ride the text-box path with an empty body, marked `renderOnly` so
 * the serializer leaves the preserved source markup as the only thing written
 * back.
 */

import { describe, expect, test } from 'bun:test';
import { parseDocumentBody } from '../documentParser';
import { serializeDocumentBody } from '../serializer/documentSerializer';
import type { Paragraph, Run, ShapeContent } from '../../types/document';

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"';

/** `fill` empty renders a shape with neither fill nor outline. */
function shapeDrawing(fill: string): string {
  return `<w:drawing>
    <wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0"
      relativeHeight="251698176" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">
      <wp:simplePos x="0" y="0"/>
      <wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH>
      <wp:positionV relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionV>
      <wp:extent cx="7620000" cy="3810000"/>
      <wp:docPr id="7" name="Rectangle 7"/>
      <a:graphic>
        <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
          <wps:wsp>
            <wps:cNvSpPr/>
            <wps:spPr>
              <a:xfrm><a:off x="0" y="0"/><a:ext cx="7620000" cy="3810000"/></a:xfrm>
              <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
              ${fill}
            </wps:spPr>
            <wps:bodyPr/>
          </wps:wsp>
        </a:graphicData>
      </a:graphic>
    </wp:anchor>
  </w:drawing>`;
}

const SOLID_FILL = '<a:solidFill><a:srgbClr val="304050"/></a:solidFill>';

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
          <w:r><w:t>Problems we solve</w:t></w:r>
        </w:p>
        <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
      </w:body>
    </w:document>`;
}

function shapesIn(body: ReturnType<typeof parseDocumentBody>): ShapeContent[] {
  const out: ShapeContent[] = [];
  for (const block of body.content) {
    if (block.type !== 'paragraph') continue;
    for (const item of (block as Paragraph).content) {
      if (item.type !== 'run') continue;
      for (const rc of (item as Run).content) {
        if (rc.type === 'shape') out.push(rc as ShapeContent);
      }
    }
  }
  return out;
}

describe('decorative filled shapes', () => {
  test('a shape with neither fill nor outline draws nothing and is not lifted', () => {
    expect(shapesIn(parseDocumentBody(documentWith(shapeDrawing(''))))).toEqual([]);
  });

  test('it reaches the model as a render-only text box carrying the fill', () => {
    const body = parseDocumentBody(documentWith(shapeDrawing(SOLID_FILL)));
    const shapes = shapesIn(body);
    expect(shapes.length).toBe(1);
    const shape = shapes[0].shape;
    expect(shape.renderOnly).toBe(true);
    // shapeType 'textBox' routes it through the anchored text-box painter.
    expect(shape.shapeType).toBe('textBox');
    expect(shape.fill?.color?.rgb).toBe('304050');
    expect(shape.textBody?.content).toEqual([]);
    // 7620000 EMU x 3810000 EMU, straight off wp:extent.
    expect(shape.size).toEqual({ width: 7620000, height: 3810000 });
    expect(shape.position?.vertical?.relativeTo).toBe('page');
  });

  test('saving writes the preserved source once, not a second modelled copy', () => {
    const body = parseDocumentBody(documentWith(shapeDrawing(SOLID_FILL)));
    const xml = serializeDocumentBody(body);
    // The original markup round-trips verbatim...
    expect(xml).toContain('<mc:AlternateContent');
    expect(xml).toContain('<a:solidFill><a:srgbClr val="304050"/></a:solidFill>');
    // ...and the render-only shape adds no duplicate wps:wsp of its own.
    expect(xml.match(/<wps:wsp>/g)?.length ?? 0).toBe(1);
  });
});

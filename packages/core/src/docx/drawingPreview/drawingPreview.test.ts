/**
 * Render-only previews for drawings the model does not represent. Every
 * fixture is synthetic markup built here.
 */
import { describe, expect, test } from 'bun:test';
import { deriveDrawingPreview, deriveVmlPreview, UNSUPPORTED_DRAWING_LABELS } from './index';
import { parseXmlDocument, type XmlElement } from '../xmlParser';
import type { MediaFile, RelationshipMap, Theme } from '../../types/document';

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ' +
  'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:v="urn:schemas-microsoft-com:vml"';

const WPS_URI = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
const WPG_URI = 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup';

interface AnchorOpts {
  uri: string;
  body: string;
  cx?: number;
  cy?: number;
  wrap?: string;
  posH?: string;
  behindDoc?: boolean;
}

function anchor(o: AnchorOpts): XmlElement {
  const cx = o.cx ?? 1828800;
  const cy = o.cy ?? 914400;
  const xml =
    `<w:drawing ${NS}><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" ` +
    `relativeHeight="251660288" behindDoc="${o.behindDoc ? 1 : 0}" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/>` +
    (o.posH ??
      `<wp:positionH relativeFrom="column"><wp:posOffset>457200</wp:posOffset></wp:positionH>`) +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>228600</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    (o.wrap ?? '<wp:wrapNone/>') +
    `<wp:docPr id="7" name="Shape 7"/><wp:cNvGraphicFramePr/>` +
    `<a:graphic><a:graphicData uri="${o.uri}">${o.body}</a:graphicData></a:graphic>` +
    `</wp:anchor></w:drawing>`;
  return parseXmlDocument(xml) as XmlElement;
}

function wsp(spPrInner: string, cx = 1828800, cy = 914400, extra = '', off = '0 0'): string {
  const [x, y] = off.split(' ');
  return (
    `<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    `${spPrInner}</wps:spPr>${extra}<wps:bodyPr/></wps:wsp>`
  );
}

const RECT = '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>';
const SOLID = (hex: string) => `<a:solidFill><a:srgbClr val="${hex}"/></a:solidFill>`;
const NO_LINE = '<a:ln><a:noFill/></a:ln>';

const ctx = { rels: null, media: null, theme: null };

/** The SVG markup behind a preview's data URL. */
function svgOf(src: string | undefined): string {
  expect(src?.startsWith('data:image/svg+xml;charset=utf-8,')).toBe(true);
  return decodeURIComponent(src!.slice(src!.indexOf(',') + 1));
}

describe('deriveDrawingPreview — single shapes', () => {
  test('a filled rectangle is drawn at the anchor frame, not as a placeholder', () => {
    const out = deriveDrawingPreview(
      anchor({ uri: WPS_URI, body: wsp(`${RECT}${SOLID('1F4E79')}${NO_LINE}`) }),
      ctx
    );
    expect(out).toHaveLength(1);
    const img = out[0].image;
    expect(img.renderOnly).toBe(true);
    expect(img.size).toEqual({ width: 1828800, height: 914400 });
    expect(img.wrap.type).toBe('inFront');
    expect(img.position?.horizontal.posOffset).toBe(457200);
    expect(img.position?.vertical.posOffset).toBe(228600);
    expect(img.relativeHeight).toBe(251660288);
    const svg = svgOf(img.src);
    expect(svg).toContain('fill="#1f4e79"');
    expect(svg).toContain('stroke="none"');
    expect(svg).not.toContain('Unsupported');
    expect(img.alt).toBeUndefined();
  });

  test('a linear gradient becomes an SVG gradient in the shape direction', () => {
    const grad =
      '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="000000"/></a:gs>' +
      '<a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs></a:gsLst><a:lin ang="5400000" scaled="1"/></a:gradFill>';
    const svg = svgOf(
      deriveDrawingPreview(anchor({ uri: WPS_URI, body: wsp(`${RECT}${grad}`) }), ctx)[0].image.src
    );
    // 90° runs top to bottom.
    expect(svg).toContain('<linearGradient id="g0" x1="0.5" y1="0" x2="0.5" y2="1">');
    expect(svg).toContain('stop-color="#000000"');
    expect(svg).toContain('stop-color="#ffffff"');
    expect(svg).toContain('fill="url(#g0)"');
  });

  test('theme colours resolve through the document theme with lumMod/lumOff and alpha', () => {
    const theme = { colorScheme: { accent1: '336699' } } as Theme;
    const fill =
      '<a:solidFill><a:schemeClr val="accent1"><a:lumMod val="50000"/><a:alpha val="40000"/></a:schemeClr></a:solidFill>';
    const svg = svgOf(
      deriveDrawingPreview(anchor({ uri: WPS_URI, body: wsp(`${RECT}${fill}`) }), {
        ...ctx,
        theme,
      })[0].image.src
    );
    // #336699 at half luminance.
    expect(svg).toMatch(/fill="#1[9a]334d"/);
    expect(svg).toContain('fill-opacity="0.4"');
  });

  test('custom geometry with guide formulas is drawn as a path', () => {
    const cust =
      '<a:custGeom><a:avLst/><a:gdLst><a:gd name="T0" fmla="*/ w 1 2"/></a:gdLst><a:pathLst>' +
      '<a:path w="100" h="100"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="100" y="0"/></a:lnTo>' +
      '<a:lnTo><a:pt x="50" y="100"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>';
    const svg = svgOf(
      deriveDrawingPreview(
        anchor({
          uri: WPS_URI,
          body: wsp(`${cust}${SOLID('FF0000')}`, 1000000, 500000),
          cx: 1000000,
          cy: 500000,
        }),
        ctx
      )[0].image.src
    );
    expect(svg).toContain('d="M0 0L1000000 0L500000 500000Z"');
    expect(svg).not.toContain('Unsupported');
  });

  test('an unknown preset is a labelled placeholder occupying the exact extent', () => {
    const out = deriveDrawingPreview(
      anchor({
        uri: WPS_URI,
        body: wsp(
          `<a:prstGeom prst="cloud"><a:avLst/></a:prstGeom>${SOLID('00FF00')}`,
          2743200,
          1371600
        ),
        cx: 2743200,
        cy: 1371600,
        wrap: '<wp:wrapSquare wrapText="bothSides"/>',
      }),
      ctx
    );
    const img = out[0].image;
    expect(img.size).toEqual({ width: 2743200, height: 1371600 });
    // Wrapping is kept so text still flows around the reserved box.
    expect(img.wrap.type).toBe('square');
    expect(img.alt).toBe(UNSUPPORTED_DRAWING_LABELS.shape);
    const svg = svgOf(img.src);
    expect(svg).toContain('stroke-dasharray="4 3"');
    expect(svg).toContain(`>${UNSUPPORTED_DRAWING_LABELS.shape}</text>`);
    expect(svg).toContain('width="2743200" height="1371600"');
  });

  test('a chart is a placeholder labelled as a chart', () => {
    const out = deriveDrawingPreview(
      anchor({
        uri: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
        body: '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId9"/>',
      }),
      ctx
    );
    expect(out[0].image.alt).toBe(UNSUPPORTED_DRAWING_LABELS.chart);
    expect(svgOf(out[0].image.src)).toContain(UNSUPPORTED_DRAWING_LABELS.chart);
  });

  test('a rotated shape carries the rotation on its frame', () => {
    const body = wsp(`${RECT}${SOLID('000000')}`).replace(
      '<a:xfrm>',
      '<a:xfrm rot="5400000" flipH="1">'
    );
    const img = deriveDrawingPreview(anchor({ uri: WPS_URI, body }), ctx)[0].image;
    expect(img.transform).toEqual({ rotation: 90, flipH: true });
  });
});

// ============================================================================
// GROUPS
// ============================================================================

const PNG = 'data:image/png;base64,iVBORw0KGgo=';
const rels = new Map([
  ['rId5', { id: 'rId5', type: 'image', target: 'media/one.png' }],
]) as unknown as RelationshipMap;
const media = new Map<string, MediaFile>([
  [
    'word/media/one.png',
    { path: 'word/media/one.png', mimeType: 'image/png', dataUrl: PNG } as unknown as MediaFile,
  ],
]);

function pic(off: string, ext: string): string {
  const [x, y] = off.split(' ');
  const [cx, cy] = ext.split(' ');
  return (
    `<pic:pic><pic:nvPicPr><pic:cNvPr id="3" name="p"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="rId5"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>${RECT}</pic:spPr></pic:pic>`
  );
}

/** A group whose child space is 1000×500, drawn into a 1828800×914400 extent. */
function group(children: string): string {
  return (
    `<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/>` +
    `<a:chOff x="1000" y="2000"/><a:chExt cx="1000" cy="500"/></a:xfrm></wpg:grpSpPr>${children}</wpg:wgp>`
  );
}

describe('deriveDrawingPreview — groups', () => {
  test('a group is ONE frame at the anchor; children map through the child space', () => {
    const out = deriveDrawingPreview(
      anchor({
        uri: WPG_URI,
        body: group(
          wsp(`${RECT}${SOLID('DDDDDD')}`, 500, 250, '', '1000 2000') + pic('1500 2250', '500 250')
        ),
        wrap: '<wp:wrapTopAndBottom/>',
      }),
      { rels, media, theme: null }
    );
    expect(out).toHaveLength(1);
    const img = out[0].image;
    expect(img.wrap.type).toBe('topAndBottom');
    const svg = svgOf(img.src);
    // The card at the child origin fills the top-left quarter...
    expect(svg).toContain(
      '<g transform="translate(0 0)"><path d="M0 0H914400V457200H0Z" fill="#dddddd"'
    );
    // ...and the picture, embedded, the bottom-right quarter.
    expect(svg).toContain('<svg x="914400" y="457200" width="914400" height="457200"');
    expect(svg).toContain(`href="${PNG}"`);
  });

  test('an aligned group keeps its alignment; children stay relative to it', () => {
    const img = deriveDrawingPreview(
      anchor({
        uri: WPG_URI,
        body: group(pic('1000 2000', '500 500') + pic('1500 2000', '500 500')),
        posH: '<wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH>',
      }),
      { rels, media, theme: null }
    )[0].image;
    expect(img.position?.horizontal.alignment).toBe('center');
    expect(img.position?.horizontal.posOffset).toBeUndefined();
    const svg = svgOf(img.src);
    expect(svg).toContain('<svg x="0" y="0" width="914400"');
    expect(svg).toContain('<svg x="914400" y="0" width="914400"');
  });

  test('a nested group composes its own child space', () => {
    const nested =
      `<wpg:grpSp><wpg:grpSpPr><a:xfrm><a:off x="1500" y="2000"/><a:ext cx="500" cy="500"/>` +
      `<a:chOff x="0" y="0"/><a:chExt cx="100" cy="100"/></a:xfrm></wpg:grpSpPr>` +
      wsp(`${RECT}${SOLID('FF0000')}`, 50, 50, '', '50 0') +
      `</wpg:grpSp>`;
    const svg = svgOf(
      deriveDrawingPreview(anchor({ uri: WPG_URI, body: group(nested) }), ctx)[0].image.src
    );
    // Inner (50,0) → outer 1500 + 250 = 1750 → root (1750-1000) * 1828.8 = 1371600.
    expect(svg).toContain('<g transform="translate(1371600 0)"><path d="M0 0H457200V457200H0Z"');
  });

  test('a text box inside a group is drawn and marked, not dropped', () => {
    const txbx =
      '<wps:txbx><w:txbxContent><w:p><w:r><w:t>x</w:t></w:r></w:p></w:txbxContent></wps:txbx>';
    const img = deriveDrawingPreview(
      anchor({
        uri: WPG_URI,
        body: group(
          wsp(`${RECT}${SOLID('EEEEEE')}`, 500, 500, txbx, '1000 2000') +
            wsp(`${RECT}${SOLID('111111')}`, 500, 500, '', '1500 2000')
        ),
      }),
      ctx
    )[0].image;
    const svg = svgOf(img.src);
    expect(svg).toContain('fill="#eeeeee"');
    expect(svg).toContain(UNSUPPORTED_DRAWING_LABELS.textbox);
    expect(svg).toContain('fill="#111111"');
  });

  test('externally loaded pictures fall back to separate render-only images', () => {
    const external = new Map<string, MediaFile>([
      [
        'word/media/one.png',
        {
          path: 'word/media/one.png',
          mimeType: 'image/png',
          assetId: 'asset-1',
        } as unknown as MediaFile,
      ],
    ]);
    const out = deriveDrawingPreview(
      anchor({ uri: WPG_URI, body: group(pic('1000 2000', '500 500')) }),
      {
        rels,
        media: external,
        theme: null,
      }
    );
    expect(out).toHaveLength(2);
    expect(out[1].image.assetId).toBe('asset-1');
    expect(out[1].image.renderOnly).toBe(true);
  });

  test('hostile nesting is bounded', () => {
    let inner = '';
    for (let i = 0; i < 3000; i++) inner = `<wpg:grpSp>${inner}</wpg:grpSp>`;
    const el = anchor({ uri: WPG_URI, body: group(inner) });
    expect(() => deriveDrawingPreview(el, ctx)).not.toThrow();
  });
});

describe('deriveVmlPreview', () => {
  test('a drawn VML shape reserves its box at its absolute position', () => {
    const pict = parseXmlDocument(
      `<w:pict ${NS}><v:rect id="r1" style="position:absolute;margin-left:36pt;margin-top:18pt;width:144pt;height:72pt;z-index:-1;mso-position-horizontal-relative:page;mso-position-vertical-relative:page" fillcolor="#c00"/></w:pict>`
    ) as XmlElement;
    const out = deriveVmlPreview(pict);
    expect(out).toHaveLength(1);
    const img = out[0].image;
    expect(img.size).toEqual({ width: 144 * 12700, height: 72 * 12700 });
    expect(img.wrap.type).toBe('behind');
    expect(img.position?.horizontal).toEqual({ relativeTo: 'page', posOffset: 36 * 12700 });
    expect(img.position?.vertical).toEqual({ relativeTo: 'page', posOffset: 18 * 12700 });
    expect(img.alt).toBe(UNSUPPORTED_DRAWING_LABELS.shape);
  });

  test('a watermark is left to the watermark parser', () => {
    const pict = parseXmlDocument(
      `<w:pict ${NS}><v:shape id="PowerPlusWaterMarkObject1" type="#_x0000_t136" style="position:absolute;width:400pt;height:100pt"/></w:pict>`
    ) as XmlElement;
    expect(deriveVmlPreview(pict)).toHaveLength(0);
  });
});

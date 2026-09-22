/**
 * A grouped drawing is preserved as opaque source, so these derived previews
 * are the only thing the canvas can paint for it.
 */
import { describe, expect, test } from 'bun:test';
import { deriveGroupPreviewImages } from './groupPreview';
import { parseXmlDocument, type XmlElement } from './xmlParser';
import type { MediaFile, RelationshipMap } from '../types/document';

const NS = {
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  wpg: 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
};

const GROUP = `<mc:AlternateContent xmlns:mc="${NS.mc}" xmlns:wp="${NS.wp}" xmlns:wpg="${NS.wpg}" xmlns:a="${NS.a}" xmlns:pic="${NS.pic}" xmlns:r="${NS.r}">
  <mc:Choice Requires="wpg"><w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <wp:anchor behindDoc="1">
      <wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>
      <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
      <wp:extent cx="1828800" cy="914400"/>
      <a:graphic><a:graphicData><wpg:wgp>
        <wpg:grpSpPr><a:xfrm>
          <a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/>
          <a:chOff x="0" y="0"/><a:chExt cx="1828800" cy="914400"/>
        </a:xfrm></wpg:grpSpPr>
        <pic:pic>
          <pic:blipFill><a:blip r:embed="rId7"/></pic:blipFill>
          <pic:spPr><a:xfrm>
            <a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/>
          </a:xfrm></pic:spPr>
        </pic:pic>
      </wpg:wgp></a:graphicData></a:graphic>
    </wp:anchor>
  </w:drawing></mc:Choice>
</mc:AlternateContent>`;

const rels = new Map([
  ['rId7', { id: 'rId7', type: 'image', target: 'media/logo.png' }],
]) as unknown as RelationshipMap;

function mediaMap(file: MediaFile): Map<string, MediaFile> {
  return new Map([
    ['word/media/logo.png', file],
    ['media/logo.png', file],
  ]);
}

const groupEl = parseXmlDocument(GROUP) as XmlElement;

const GROUP_ALIGNED = GROUP.replace(
  '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>',
  '<wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH>'
);

const GROUP_ALT_PREFIX = GROUP.replace('r:embed="rId7"', 'rel:embed="rId7"').replace(
  `xmlns:r="${NS.r}"`,
  `xmlns:r="${NS.r}" xmlns:rel="${NS.r}"`
);

const EMBEDDED = {
  path: 'word/media/logo.png',
  filename: 'logo.png',
  mimeType: 'image/png',
  dataUrl: 'data:image/png;base64,iVBORw0KGgo=',
} as unknown as MediaFile;

describe('deriveGroupPreviewImages', () => {
  test('derives a paintable preview from embedded bytes', () => {
    const out = deriveGroupPreviewImages(
      groupEl,
      rels,
      mediaMap({
        path: 'word/media/logo.png',
        filename: 'logo.png',
        mimeType: 'image/png',
        dataUrl: 'data:image/png;base64,iVBORw0KGgo=',
      } as unknown as MediaFile)
    );
    expect(out).toHaveLength(1);
    expect(out[0].image.renderOnly).toBe(true);
  });

  test('derives a preview in externalMedia mode (assetId, no inline src)', () => {
    // buildMediaMap deliberately produces entries with an assetId and no bytes;
    // the painter resolves those lazily, so a preview must still be derived.
    const out = deriveGroupPreviewImages(
      groupEl,
      rels,
      mediaMap({
        path: 'word/media/logo.png',
        assetId: 'asset-logo',
        filename: 'logo.png',
        mimeType: 'image/png',
      } as unknown as MediaFile)
    );
    expect(out).toHaveLength(1);
    expect(out[0].image.assetId).toBe('asset-logo');
    expect(out[0].image.renderOnly).toBe(true);
  });

  test('skips a picture with neither bytes nor an assetId', () => {
    const out = deriveGroupPreviewImages(
      groupEl,
      rels,
      mediaMap({
        path: 'word/media/logo.png',
        filename: 'logo.png',
        mimeType: 'image/png',
      } as unknown as MediaFile)
    );
    expect(out).toHaveLength(0);
  });

  test('reads the relationship id whatever prefix is bound to it', () => {
    // Only the convention binds this namespace to `r`; assuming it silently
    // dropped every picture in the group.
    const el = parseXmlDocument(GROUP_ALT_PREFIX) as XmlElement;
    expect(deriveGroupPreviewImages(el, rels, mediaMap(EMBEDDED))).toHaveLength(1);
  });

  test('honours wp:align instead of pinning to the top-left', () => {
    const el = parseXmlDocument(GROUP_ALIGNED) as XmlElement;
    const out = deriveGroupPreviewImages(el, rels, mediaMap(EMBEDDED));
    expect(out).toHaveLength(1);
    expect(out[0].image.position?.horizontal.alignment).toBe('center');
    expect(out[0].image.position?.horizontal.relativeTo).toBe('margin');
    // The painters check `posOffset` BEFORE alignment, so emitting both would
    // pin the group back to the origin and make the alignment inert.
    expect(out[0].image.position?.horizontal.posOffset).toBeUndefined();
  });

  test('a deeply nested group does not blow the stack', () => {
    // A `.docx` is attacker-controlled XML; nesting must be bounded.
    let inner = '<wpg:wgp/>';
    for (let i = 0; i < 5000; i++) inner = `<wpg:grpSp>${inner}</wpg:grpSp>`;
    const hostile = GROUP.replace('<wpg:wgp>', `<wpg:wgp>${inner}`).replace(
      '</wpg:wgp>',
      '</wpg:wgp>'
    );
    const el = parseXmlDocument(hostile) as XmlElement;
    expect(() => deriveGroupPreviewImages(el, rels, mediaMap(EMBEDDED))).not.toThrow();
  });

  test("a picture in a NESTED group uses that group's coordinate space", () => {
    // Outer group: 1828800x914400 EMU mapped 1:1. Inner group sits at
    // (914400, 0) and halves its children (chExt 2x ext), so a picture at
    // inner-local (914400, 0) lands at outer 914400 + 914400/2 = 1371600 EMU.
    const nested = GROUP.replace(
      '<pic:pic>',
      `<wpg:grpSp><wpg:grpSpPr><a:xfrm>
         <a:off x="914400" y="0"/><a:ext cx="914400" cy="457200"/>
         <a:chOff x="0" y="0"/><a:chExt cx="1828800" cy="914400"/>
       </a:xfrm></wpg:grpSpPr>
       <pic:pic>
         <pic:blipFill><a:blip r:embed="rId7"/></pic:blipFill>
         <pic:spPr><a:xfrm><a:off x="914400" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm></pic:spPr>
       </pic:pic>
       </wpg:grpSp><pic:pic>`
    );
    const out = deriveGroupPreviewImages(
      parseXmlDocument(nested) as XmlElement,
      rels,
      mediaMap(EMBEDDED)
    );
    expect(out).toHaveLength(2);
    const offsets = out.map((d) => d.image.position?.horizontal.posOffset).sort((a, b) => a! - b!);
    // Outer picture at 0; nested one at 1371600 — NOT 914400, which is what
    // applying only the outer group's transform would give.
    expect(offsets[0]).toBe(0);
    expect(offsets[1]).toBe(1371600);
    // And it is scaled by the inner group too: 914400 * 0.5 EMU wide.
    const nestedImg = out.find((d) => d.image.position?.horizontal.posOffset === 1371600)!;
    expect(nestedImg.image.size?.width).toBe(457200);
  });
});

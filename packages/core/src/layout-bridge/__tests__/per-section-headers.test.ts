/**
 * Headers belong to sections, not to documents.
 *
 * The layout pass used to resolve ONE header/footer pair and apply it
 * everywhere — React took the last section's, Vue the first. On a proposal
 * template whose cover declares no header and whose body declares one, that
 * meant the cover's body was pushed down by a band it does not have (so the
 * full-bleed artwork no longer reached the page top) while the body pages lost
 * the logo and page numbers they do declare.
 */
import { describe, expect, test } from 'bun:test';
import { resolveSectionHeaderFooters } from '../sectionGeometry';
import { extendMarginsForHeaderFooter } from '../headerFooterMargins';
import type { Document, HeaderFooter, SectionProperties } from '../../types/document';
import type { FlowBlock, PageMargins } from '../../layout-engine/types';
import type { HeaderFooterContent } from '../../layout-painter';

const hf = (id: string): HeaderFooter =>
  ({ hdrFtrType: 'default', content: [], id }) as unknown as HeaderFooter;

function docWithSections(props: SectionProperties[]): Document {
  return {
    package: {
      headers: new Map([
        ['rIdBody', hf('body-header')],
        ['rIdBack', hf('back-header')],
      ]),
      footers: new Map([['rIdFoot', hf('body-footer')]]),
      document: {
        content: [],
        sections: props.map((p) => ({ properties: p, content: [] })),
        finalSectionProperties: props[props.length - 1],
      },
    },
  } as unknown as Document;
}

const COVER: SectionProperties = { marginTop: 0, marginBottom: 0, headerDistance: 708 };
const BODY: SectionProperties = {
  marginTop: 1200,
  marginBottom: 1200,
  headerReferences: [{ type: 'default', rId: 'rIdBody' }],
  footerReferences: [{ type: 'default', rId: 'rIdFoot' }],
} as SectionProperties;
const BACK: SectionProperties = {
  marginTop: 0,
  headerReferences: [{ type: 'default', rId: 'rIdBack' }],
} as SectionProperties;

describe('resolveSectionHeaderFooters', () => {
  test('gives each section its own, in Page.sectionIndex order', () => {
    const out = resolveSectionHeaderFooters(docWithSections([COVER, BODY, BACK]));
    expect(out).toHaveLength(3);
    expect(out[0].header).toBeNull(); // the cover declares none
    expect((out[1].header as { id: string } | null)?.id).toBe('body-header');
    expect((out[2].header as { id: string } | null)?.id).toBe('back-header');
    expect((out[1].footer as { id: string } | null)?.id).toBe('body-footer');
  });

  test("a section's own header distance travels with it", () => {
    const out = resolveSectionHeaderFooters(docWithSections([COVER, BODY, BACK]));
    expect(out[0].headerDistance).toBeCloseTo((708 / 1440) * 96, 0);
  });

  test('a document with no sections still resolves one entry', () => {
    const doc = {
      package: { headers: new Map(), footers: new Map(), document: { content: [] } },
    } as unknown as Document;
    expect(resolveSectionHeaderFooters(doc)).toHaveLength(1);
  });
});

describe("a cover's first-page-only artwork stays on the cover", () => {
  test('an inherited `first` ref is not promoted to a later section\u2019s default', () => {
    // The cover declares only a first-page footer, under `w:titlePg`. A later
    // section that declares nothing inherits that ref (§17.6) but has no
    // `titlePg` of its own — the no-titlePg fallback must not turn the cover
    // artwork into that section's every-page footer.
    const doc = docWithSections([
      {
        titlePg: true,
        footerReferences: [{ type: 'first', rId: 'rIdFoot' }],
      } as SectionProperties,
      {} as SectionProperties,
    ]);
    const out = resolveSectionHeaderFooters(doc);

    expect((out[0].firstFooter as { id: string } | null)?.id).toBe('body-footer');
    expect(out[0].footer).toBeNull();
    // Section 1 inherited the ref but must render no footer at all.
    expect(out[1].footer).toBeNull();
  });

  test("a section's OWN first-page-only ref still serves as its default", () => {
    // The Word quirk the fallback exists for: only a `first` ref, no titlePg.
    const doc = docWithSections([
      { footerReferences: [{ type: 'first', rId: 'rIdFoot' }] } as SectionProperties,
    ]);
    const out = resolveSectionHeaderFooters(doc);
    expect((out[0].footer as { id: string } | null)?.id).toBe('body-footer');
  });
});

describe('the margin extension is per section', () => {
  const band = (height: number): HeaderFooterContent =>
    ({ blocks: [], measures: [], height, flowHeight: height }) as HeaderFooterContent;

  const margins = (top: number, header: number): PageMargins => ({
    top,
    right: 0,
    bottom: 0,
    left: 0,
    header,
    footer: 0,
  });

  test('a section with no header keeps its authored top margin', () => {
    // The cover: `w:pgMar top="0"`, no header. Nothing to clear, so the body
    // starts at the page top and the full-bleed artwork reaches it.
    const out = extendMarginsForHeaderFooter({
      pageSize: { w: 816, h: 1056 },
      margins: margins(0, 47),
      finalMargins: margins(0, 47),
      sections: [{}, { header: band(34) }],
      bodyBlocks: [],
    });
    expect(out.margins.top).toBe(0);
  });

  test('the `w:header` distance alone does not push a body that has no band', () => {
    const out = extendMarginsForHeaderFooter({
      pageSize: { w: 816, h: 1056 },
      margins: margins(0, 47),
      finalMargins: margins(0, 47),
      sections: [{ header: undefined, footer: undefined }],
      bodyBlocks: [],
    });
    expect(out.margins.top).toBe(0);
    expect(out.margins.bottom).toBe(0);
  });

  test('each section break carries the band of the section at its own index', () => {
    // `collectSectionConfigs` reads break i as section i's geometry, so the
    // extension has to line up with that or a section grows by a stranger's
    // header.
    const cover = {
      kind: 'sectionBreak',
      id: 'sb-0',
      margins: margins(0, 47),
    } as unknown as FlowBlock;
    const body = {
      kind: 'sectionBreak',
      id: 'sb-1',
      margins: margins(10, 47),
    } as unknown as FlowBlock;
    extendMarginsForHeaderFooter({
      pageSize: { w: 816, h: 1056 },
      margins: margins(0, 47),
      finalMargins: margins(0, 47),
      // Section 0 has nothing; section 1 has a band.
      sections: [{}, { header: band(34) }, {}],
      bodyBlocks: [cover, body],
    });
    expect((cover as { margins: PageMargins }).margins.top).toBe(0);
    expect((body as { margins: PageMargins }).margins.top).toBe(81);
  });

  test('a tall header in one section does not reach into the one before it', () => {
    const cover = {
      kind: 'sectionBreak',
      id: 'sb-0',
      margins: margins(0, 47),
    } as unknown as FlowBlock;
    extendMarginsForHeaderFooter({
      pageSize: { w: 816, h: 1056 },
      margins: margins(0, 47),
      finalMargins: margins(0, 47),
      sections: [{}, { header: band(300) }],
      bodyBlocks: [cover],
    });
    expect((cover as { margins: PageMargins }).margins.top).toBe(0);
  });

  test('without `sections` the single-pair behaviour is unchanged', () => {
    const out = extendMarginsForHeaderFooter({
      pageSize: { w: 816, h: 1056 },
      margins: margins(10, 47),
      finalMargins: margins(10, 47),
      headers: [band(34)],
      footers: [],
      bodyBlocks: [],
    });
    expect(out.margins.top).toBe(81);
  });
});

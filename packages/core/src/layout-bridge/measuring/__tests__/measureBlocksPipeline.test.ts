import { describe, expect, test } from 'bun:test';
import { measureBlocksWithFloats } from '../measureBlocksPipeline';
import type { FloatingImageZone } from '../floatingZones';
import type { FlowBlock, Measure, ParagraphBlock, TextBoxBlock } from '../../../layout-engine';

// A page-relative topAndBottom banner pinned to the top of the page. Its anchor
// (the textBox block) sits AFTER the first paragraph, but Word reserves the band
// from the top of the page — so the band must reach the preceding block.
function makeBlocks(): FlowBlock[] {
  const para = (id: string): ParagraphBlock =>
    ({
      kind: 'paragraph',
      id,
      pmStart: 0,
      pmEnd: 0,
      runs: [],
      paragraphProperties: {},
    }) as unknown as ParagraphBlock;
  const banner: TextBoxBlock = {
    kind: 'textBox',
    id: 'banner',
    pmStart: 0,
    pmEnd: 0,
    width: 600,
    height: 100,
    displayMode: 'block',
    wrapType: 'topAndBottom',
    position: {
      vertical: { relativeTo: 'page', posOffset: 0 },
      horizontal: { relativeTo: 'column', posOffset: 0 },
    },
    content: [],
  } as unknown as TextBoxBlock;
  return [para('p0'), banner, para('p1')];
}

describe('measureBlocksWithFloats — topAndBottom page-pinned band', () => {
  test('reserves a full-width band that reaches the block before the anchor', () => {
    const seen: Array<FloatingImageZone[] | undefined> = [];
    const measureBlock = (block: FlowBlock, _w: number, zones?: FloatingImageZone[]): Measure => {
      seen.push(zones);
      if (block.kind === 'textBox') {
        return { kind: 'textBox', width: 600, height: 100, innerMeasures: [] } as Measure;
      }
      return { kind: 'paragraph', lines: [], totalHeight: 20 } as Measure;
    };

    // marginTop=50: a page-relative posOffset=0 banner sits at content Y -50,
    // so its band intrudes into content as [0, height - marginTop] = [0, 50].
    measureBlocksWithFloats(makeBlocks(), 600, measureBlock, {
      pageWidth: 700,
      pageHeight: 900,
      marginLeft: 50,
      marginTop: 50,
      contentWidth: 600,
      contentHeight: 800,
    });

    // Block 0 (the paragraph BEFORE the banner's anchor) must see the band.
    const block0Zones = seen[0];
    expect(block0Zones).toBeDefined();
    expect(block0Zones).toHaveLength(1);
    expect(block0Zones?.[0].fullWidthBlock).toBe(true);
    expect(block0Zones?.[0].topY).toBe(0);
    expect(block0Zones?.[0].bottomY).toBe(50);
  });
});

// A tall float anchored on one page must not keep reserving space in
// paragraphs that land on a later page.
describe('measureBlocksWithFloats — a float does not cross a page break', () => {
  function blocksWithBreak(): FlowBlock[] {
    const para = (id: string, pageBreakBefore = false): ParagraphBlock =>
      ({
        kind: 'paragraph',
        id,
        pmStart: 0,
        pmEnd: 0,
        runs: [],
        paragraphProperties: {},
        attrs: pageBreakBefore ? { pageBreakBefore: true } : {},
      }) as unknown as ParagraphBlock;

    // A tall square-wrap text box on the cover, then a page break, then body.
    const cover: TextBoxBlock = {
      kind: 'textBox',
      id: 'coverBox',
      pmStart: 0,
      pmEnd: 0,
      width: 300,
      height: 5000, // far taller than the flow that follows
      displayMode: 'float',
      wrapType: 'square',
      position: {
        vertical: { relativeTo: 'paragraph', posOffset: 0 },
        horizontal: { relativeTo: 'column', posOffset: 0 },
      },
      content: [],
    } as unknown as TextBoxBlock;

    return [para('cover'), cover, para('afterBreak', true), para('body')];
  }

  test('zones stop at the page break instead of squeezing later pages', () => {
    const seen: Array<{ id: string; zones?: FloatingImageZone[] }> = [];
    const measureBlock = (block: FlowBlock, _w: number, zones?: FloatingImageZone[]): Measure => {
      seen.push({ id: (block as { id: string }).id, zones });
      if (block.kind === 'textBox') {
        return { kind: 'textBox', width: 300, height: 5000, innerMeasures: [] } as Measure;
      }
      return { kind: 'paragraph', lines: [], totalHeight: 20 } as Measure;
    };

    measureBlocksWithFloats(blocksWithBreak(), 600, measureBlock, {
      pageWidth: 700,
      pageHeight: 900,
      marginLeft: 50,
      marginTop: 50,
      contentWidth: 600,
      contentHeight: 800,
    });

    const afterBreak = seen.find((s) => s.id === 'afterBreak');
    const body = seen.find((s) => s.id === 'body');
    expect(afterBreak?.zones).toBeUndefined();
    expect(body?.zones).toBeUndefined();
  });
});

// Two floats on opposite sides of a page break must keep their own zones: the
// grouping pass used to re-anchor both to the first float, so clearing at the
// break dropped the second page's zone with no anchor left to restore it.
describe('measureBlocksWithFloats — floats either side of a page break', () => {
  test('the float after the break still reserves space on its own page', () => {
    const para = (id: string, pageBreakBefore = false): ParagraphBlock =>
      ({
        kind: 'paragraph',
        id,
        pmStart: 0,
        pmEnd: 0,
        runs: [],
        paragraphProperties: {},
        attrs: pageBreakBefore ? { pageBreakBefore: true } : {},
      }) as unknown as ParagraphBlock;

    const box = (id: string): TextBoxBlock =>
      ({
        kind: 'textBox',
        id,
        pmStart: 0,
        pmEnd: 0,
        width: 200,
        height: 100,
        displayMode: 'float',
        wrapType: 'square',
        position: {
          vertical: { relativeTo: 'paragraph', posOffset: 0 },
          horizontal: { relativeTo: 'column', posOffset: 0 },
        },
        content: [],
      }) as unknown as TextBoxBlock;

    // Anchors are within ANCHOR_PROXIMITY (4) and the Y ranges overlap, so the
    // two boxes would previously be merged and re-anchored to the first.
    const blocks: FlowBlock[] = [
      para('p1'),
      box('boxPage1'),
      para('p2'),
      para('breaker', true),
      box('boxPage2'),
      para('afterSecondBox'),
    ];

    const seen: Array<{ id: string; zones?: FloatingImageZone[] }> = [];
    const measureBlock = (block: FlowBlock, _w: number, zones?: FloatingImageZone[]): Measure => {
      seen.push({ id: (block as { id: string }).id, zones });
      if (block.kind === 'textBox') {
        return { kind: 'textBox', width: 200, height: 100, innerMeasures: [] } as Measure;
      }
      return { kind: 'paragraph', lines: [], totalHeight: 20 } as Measure;
    };

    measureBlocksWithFloats(blocks, 600, measureBlock, {
      pageWidth: 700,
      pageHeight: 900,
      marginLeft: 50,
      marginTop: 50,
      contentWidth: 600,
      contentHeight: 800,
    });

    // Page one keeps its wrap...
    expect(seen.find((s) => s.id === 'p2')?.zones?.length).toBeGreaterThan(0);
    // ...and page two gets its own, rather than inheriting nothing.
    expect(seen.find((s) => s.id === 'afterSecondBox')?.zones?.length).toBeGreaterThan(0);
  });
});

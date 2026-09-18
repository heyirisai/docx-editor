/**
 * Header/footer `behindDoc` floats paint under the flow text, and must keep
 * document order among themselves — a template's background band is written
 * before the logo that sits on it, so the logo has to stay on top.
 */
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import {
  renderHeaderFooterContent,
  type HeaderFooterContent,
  type HeaderFooterLayoutInfo,
} from '../renderPage/headerFooter';
import type { ImageRun, ParagraphBlock, ParagraphMeasure } from '../../layout-engine/types';
import type { RenderContext } from '../renderPage';

beforeAll(() => GlobalRegistrator.register());
afterAll(() => GlobalRegistrator.unregister());

const ctx: RenderContext = { pageNumber: 1, totalPages: 1, section: 'body', contentWidth: 600 };
const layout: HeaderFooterLayoutInfo = {
  flowTop: 48,
  flowLeft: 72,
  contentWidth: 600,
  pageWidth: 744,
  pageHeight: 1123,
  margins: { top: 102, right: 72, bottom: 72, left: 72 },
};

function floatImage(alt: string, wrapType: string): ImageRun {
  return {
    kind: 'image',
    src: 'data:image/png;base64,iVBORw0KGgo=',
    alt,
    width: 100,
    height: 50,
    wrapType,
    position: {
      horizontal: { relativeTo: 'column', posOffset: 0 },
      vertical: { relativeTo: 'paragraph', posOffset: 0 },
    },
  } as ImageRun;
}

function paintedAltOrder(...runs: ImageRun[]): string[] {
  const block: ParagraphBlock = {
    kind: 'paragraph',
    id: 'p1',
    runs: [...runs, { kind: 'text', text: 'FLOW' }],
  } as ParagraphBlock;
  const measure: ParagraphMeasure = {
    kind: 'paragraph',
    lines: [
      {
        fromRun: 0,
        fromChar: 0,
        toRun: 0,
        toChar: 4,
        width: 40,
        ascent: 12,
        descent: 4,
        lineHeight: 18,
      },
    ],
    totalHeight: 18,
  };
  const content: HeaderFooterContent = {
    blocks: [block],
    measures: [measure],
    height: 18,
    flowHeight: 18,
    visualTop: 0,
    visualBottom: 18,
  };
  const el = renderHeaderFooterContent(content, ctx, { document }, layout);
  return Array.from(el.querySelectorAll<HTMLImageElement>('img')).map((img) => img.alt);
}

describe('behindDoc floats in a header/footer', () => {
  test('keep document order among themselves', () => {
    // Earlier in the document paints first, i.e. lower in the stack.
    expect(paintedAltOrder(floatImage('band', 'behind'), floatImage('logo', 'behind'))).toEqual([
      'band',
      'logo',
    ]);
  });

  test('still paint before the flow content', () => {
    const el = paintedAltOrder(floatImage('band', 'behind'), floatImage('front', 'square'));
    // The square-wrapped float is appended after the flow; the behind one leads.
    expect(el[0]).toBe('band');
    expect(el).toContain('front');
  });
});

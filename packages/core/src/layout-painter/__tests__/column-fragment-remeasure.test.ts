/**
 * When a page carries floats the painter re-measures each paragraph, because
 * only here is the fragment's real Y on the page known. That re-measure has to
 * use the FRAGMENT's width: in a multi-column section the fragment is one
 * column wide, and re-breaking its lines against the whole content width
 * painted them straight across the page, one column's text over the other's
 * (ideagen's two-column "Core Capabilities" block).
 */
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import { renderPage } from '../renderPage';
import type {
  FlowBlock,
  Measure,
  Page,
  ParagraphBlock,
  TextBoxBlock,
} from '../../layout-engine/types';
import { measureParagraph } from '../../layout-bridge/measuring';

let originalGetContext: typeof HTMLCanvasElement.prototype.getContext | undefined;

beforeAll(() => {
  GlobalRegistrator.register();
  // happy-dom has no canvas; a fixed 9px-per-character metric is enough to
  // compare line counts between two widths.
  originalGetContext = HTMLCanvasElement.prototype.getContext;
  HTMLCanvasElement.prototype.getContext = function getContext(type: string) {
    if (type !== '2d') return null;
    return {
      font: '',
      measureText: (text: string) => ({
        width: text.length * 9,
        actualBoundingBoxAscent: 12,
        actualBoundingBoxDescent: 4,
      }),
    } as unknown as CanvasRenderingContext2D;
  } as typeof HTMLCanvasElement.prototype.getContext;
});

afterAll(() => {
  if (originalGetContext) HTMLCanvasElement.prototype.getContext = originalGetContext;
  GlobalRegistrator.unregister();
});

const PAGE_W = 816;
const PAGE_H = 1056;
const MARGINS = { top: 96, right: 96, bottom: 96, left: 96 };
const CONTENT_W = PAGE_W - MARGINS.left - MARGINS.right; // 624
const COLUMN_W = 288; // two 288px columns with a 48px gutter

const TEXT =
  'Ideagen Quality Management delivers a comprehensive suite of capabilities ' +
  'designed to support organisations in achieving quality excellence, regulatory ' +
  'compliance, and operational efficiency across every regulated industry.';

function paragraph(id: string): ParagraphBlock {
  return { kind: 'paragraph', id, runs: [{ kind: 'text', text: TEXT }] } as ParagraphBlock;
}

/** A full-width band so the page has a float and the re-measure path runs. */
function band(): TextBoxBlock {
  return {
    kind: 'textBox',
    id: 'band',
    width: CONTENT_W,
    height: 400,
    content: [],
    displayMode: 'float',
    cssFloat: 'left',
    wrapType: 'square',
    position: {
      horizontal: { relativeTo: 'column', posOffset: 0 },
      vertical: { relativeTo: 'paragraph', posOffset: 0 },
    },
  } as unknown as TextBoxBlock;
}

function renderColumnPage(fragmentWidth: number): HTMLElement {
  const para = paragraph('col-para');
  const measure = measureParagraph(para, fragmentWidth);
  const blockLookup = new Map<string, { block: FlowBlock; measure: Measure }>([
    ['col-para', { block: para, measure }],
    [
      'band',
      {
        block: band(),
        measure: { kind: 'textBox', width: CONTENT_W, height: 400, innerMeasures: [] },
      },
    ],
  ]);
  const page: Page = {
    number: 1,
    size: { w: PAGE_W, h: PAGE_H },
    margins: MARGINS,
    fragments: [
      {
        kind: 'textBox',
        blockId: 'band',
        x: MARGINS.left,
        y: MARGINS.top,
        width: CONTENT_W,
        height: 400,
        isFloating: true,
      },
      {
        kind: 'paragraph',
        blockId: 'col-para',
        x: MARGINS.left,
        y: MARGINS.top + 420,
        width: fragmentWidth,
        height: measure.totalHeight,
        fromLine: 0,
        toLine: measure.lines.length,
      },
    ],
  } as unknown as Page;

  return renderPage(
    page,
    { pageNumber: 1, totalPages: 1, section: 'body', contentWidth: CONTENT_W },
    { document, blockLookup }
  );
}

function lineCount(el: HTMLElement): number {
  return el.querySelectorAll('[data-block-id="col-para"] .layout-line').length;
}

describe('the painter re-measure inside a column', () => {
  test('breaks lines at the fragment width, not the page content width', () => {
    const narrow = lineCount(renderColumnPage(COLUMN_W));
    const full = lineCount(renderColumnPage(CONTENT_W));
    // A 288px column needs strictly more lines than the 624px content area.
    expect(narrow).toBeGreaterThan(full);
  });

  test('a full-width fragment is unaffected', () => {
    const para = paragraph('col-para');
    const standalone = measureParagraph(para, CONTENT_W).lines.length;
    expect(lineCount(renderColumnPage(CONTENT_W))).toBe(standalone);
  });
});

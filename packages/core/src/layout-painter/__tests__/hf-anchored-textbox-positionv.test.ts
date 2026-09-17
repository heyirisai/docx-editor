import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import { PAGE_CLASS_NAMES, renderPage, type HeaderFooterContent } from '../renderPage';
import type { Page, TextBoxBlock, TextBoxMeasure } from '../../layout-engine/types';

beforeAll(() => GlobalRegistrator.register());
afterAll(() => GlobalRegistrator.unregister());

// Cashout cover: the footer date sits in a text box anchored with a NEGATIVE
// positionV so it paints above the footer band, over the body artwork.
const FOOTER_DISTANCE = 37.8;
const BOX_HEIGHT = 35.5;
const OFFSET_PX = -49.8;
const PAGE_H = 1123;

function makePage(): Page {
  return {
    number: 1,
    fragments: [],
    margins: { top: 151, right: 96, bottom: 113, left: 96, header: 37.8, footer: FOOTER_DISTANCE },
    size: { w: 794, h: PAGE_H },
  };
}

function makeFooter(): HeaderFooterContent {
  const block: TextBoxBlock = {
    kind: 'textBox',
    id: 'date-box',
    width: 112,
    height: BOX_HEIGHT,
    content: [],
    displayMode: 'float',
    position: {
      horizontal: { relativeTo: 'margin', posOffset: 4845050 },
      vertical: { relativeTo: 'paragraph', posOffset: -474535 },
    },
    pmStart: 0,
    pmEnd: 1,
  } as TextBoxBlock;
  const measure: TextBoxMeasure = {
    kind: 'textBox',
    width: 112,
    height: BOX_HEIGHT,
    content: [],
  } as unknown as TextBoxMeasure;
  return {
    blocks: [block],
    measures: [measure],
    height: BOX_HEIGHT,
    flowHeight: 0,
    visualTop: OFFSET_PX,
    visualBottom: OFFSET_PX + BOX_HEIGHT,
  } as HeaderFooterContent;
}

function footerBoxTop(): { footerTop: number; contentTop: number; boxTop: number } {
  const pageEl = renderPage(
    makePage(),
    { pageNumber: 1, totalPages: 1, section: 'body' },
    { document, footerContent: makeFooter() }
  );
  const footerEl = pageEl.querySelector(`.${PAGE_CLASS_NAMES.footer}`) as HTMLElement;
  const contentEl = footerEl.firstElementChild as HTMLElement;
  const boxEl = contentEl.querySelector('div') as HTMLElement;
  return {
    footerTop: parseFloat(footerEl.style.top),
    contentTop: parseFloat(contentEl.style.top),
    boxTop: parseFloat(boxEl.style.top),
  };
}

describe('HF anchored text box positionV', () => {
  test('paints at bandTop + posOffset, matching Word', () => {
    const { footerTop, contentTop, boxTop } = footerBoxTop();
    const absolute = footerTop + contentTop + boxTop;
    const bandTop = PAGE_H - FOOTER_DISTANCE;
    expect(absolute).toBeCloseTo(bandTop + OFFSET_PX, 0);
  });

  test('the footer does not clip content anchored above the band', () => {
    const pageEl = renderPage(
      makePage(),
      { pageNumber: 1, totalPages: 1, section: 'body' },
      { document, footerContent: makeFooter() }
    );
    const footerEl = pageEl.querySelector(`.${PAGE_CLASS_NAMES.footer}`) as HTMLElement;
    expect(footerEl.style.overflow).not.toBe('hidden');
  });
});

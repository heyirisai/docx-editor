/**
 * Header/footer content paints behind the body, as in Word: a header picture
 * that reaches into the body area (e.g. a page-wide logo with a white
 * background) must not cover body text. The header and footer bands are
 * inserted before the body content area, so DOM order — not a z-index the
 * float could escape with — puts them underneath it.
 */
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import { PAGE_CLASS_NAMES, renderPage, type HeaderFooterContent } from '../renderPage';
import type { ImageRun, Page, ParagraphBlock, ParagraphMeasure } from '../../layout-engine/types';

beforeAll(() => GlobalRegistrator.register());
afterAll(() => GlobalRegistrator.unregister());

const EMU_PER_PX = 9525;

function makePage(): Page {
  return {
    number: 1,
    fragments: [],
    margins: { top: 96, right: 120, bottom: 96, left: 120, header: 48, footer: 48 },
    size: { w: 794, h: 1123 },
  };
}

function storyWithPicture(): HeaderFooterContent {
  const picture: ImageRun = {
    kind: 'image',
    src: 'data:image/png;base64,iVBORw0KGgo=',
    width: 797,
    height: 144,
    wrapType: 'square',
    displayMode: 'block',
    cssFloat: 'none',
    position: {
      horizontal: { relativeTo: 'page', posOffset: 0 },
      vertical: { relativeTo: 'paragraph', posOffset: -58 * EMU_PER_PX },
    },
  };
  const block: ParagraphBlock = { kind: 'paragraph', id: 'hp', runs: [picture] };
  const measure: ParagraphMeasure = { kind: 'paragraph', lines: [], totalHeight: 18 };
  return { blocks: [block], measures: [measure], height: 18, flowHeight: 18 };
}

describe('header/footer paint behind the body', () => {
  test('header and footer bands come before the body content area', () => {
    const pageEl = renderPage(
      makePage(),
      { pageNumber: 1, totalPages: 1, section: 'body' },
      { document, headerContent: storyWithPicture(), footerContent: storyWithPicture() }
    );
    const order = Array.from(pageEl.children).map((el) => el.className);
    const content = order.indexOf(PAGE_CLASS_NAMES.content);
    expect(content).toBeGreaterThanOrEqual(0);
    expect(order.indexOf(PAGE_CLASS_NAMES.header)).toBeLessThan(content);
    expect(order.indexOf(PAGE_CLASS_NAMES.footer)).toBeLessThan(content);
  });

  test('the header picture itself carries no z-index that could lift it over the body', () => {
    const pageEl = renderPage(
      makePage(),
      { pageNumber: 1, totalPages: 1, section: 'body' },
      { document, headerContent: storyWithPicture() }
    );
    const img = pageEl.querySelector(`.${PAGE_CLASS_NAMES.header} img`) as HTMLElement;
    expect(img).not.toBeNull();
    expect(img.style.zIndex).toBe('');
  });

  test('bands render even without header/footer content (double-click targets)', () => {
    const pageEl = renderPage(
      makePage(),
      { pageNumber: 1, totalPages: 1, section: 'body' },
      {
        document,
      }
    );
    const order = Array.from(pageEl.children).map((el) => el.className);
    expect(order.indexOf(PAGE_CLASS_NAMES.header)).toBeLessThan(
      order.indexOf(PAGE_CLASS_NAMES.content)
    );
  });
});

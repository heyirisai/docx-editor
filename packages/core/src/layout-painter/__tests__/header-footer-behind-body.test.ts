/**
 * Header/footer pictures paint behind the body, as in Word: a header picture
 * that reaches into the body area (e.g. a page-wide logo with a white
 * background) must not cover body text. The header and footer bands are
 * inserted before the body content area, so DOM order — not a z-index the
 * float could escape with — puts them underneath it. (Floating text boxes keep
 * their own front z-index.)
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

const box = (el: HTMLElement) => ({
  top: parseFloat(el.style.top),
  bottom: parseFloat(el.style.top) + parseFloat(el.style.height),
});

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

  test('an empty header/footer click strip stays in the margin, clear of the body', () => {
    // Word's "Narrow" preset: 0.5in margins and 0.5in header/footer distance.
    // The strip at the distance would lie under the body (painted above the
    // bands), making "Double-click to add header" unreachable.
    const narrow: Page = {
      ...makePage(),
      margins: { top: 48, right: 48, bottom: 48, left: 48, header: 48, footer: 48 },
    };
    const pageEl = renderPage(
      narrow,
      { pageNumber: 1, totalPages: 1, section: 'body' },
      { document }
    );
    const header = pageEl.querySelector(`.${PAGE_CLASS_NAMES.header}`) as HTMLElement;
    const footer = pageEl.querySelector(`.${PAGE_CLASS_NAMES.footer}`) as HTMLElement;
    expect(box(header)).toEqual({ top: 24, bottom: 48 });
    expect(box(footer)).toEqual({ top: 1123 - 48, bottom: 1123 - 24 });
    // Clear of the body, so they stay under it (no lift).
    expect(header.style.zIndex).toBe('');
    expect(footer.style.zIndex).toBe('');

    // Default margins leave room at the distance: the strip stays there.
    const roomy = renderPage(
      makePage(),
      { pageNumber: 1, totalPages: 1, section: 'body' },
      { document }
    );
    expect(box(roomy.querySelector(`.${PAGE_CLASS_NAMES.header}`) as HTMLElement)).toEqual({
      top: 48,
      bottom: 72,
    });
    expect(box(roomy.querySelector(`.${PAGE_CLASS_NAMES.footer}`) as HTMLElement)).toEqual({
      top: 1123 - 72,
      bottom: 1123 - 48,
    });
  });

  test('at a zero margin an empty header/footer strip sits on the page edge, lifted over the body', () => {
    // No margin leaves no room clear of the body: without a strip on the page
    // edge there is no "Double-click to add header" target at all.
    const flush: Page = {
      ...makePage(),
      margins: { top: 0, right: 48, bottom: 0, left: 48, header: 48, footer: 48 },
    };
    const pageEl = renderPage(
      flush,
      { pageNumber: 1, totalPages: 1, section: 'body' },
      { document }
    );
    const header = pageEl.querySelector(`.${PAGE_CLASS_NAMES.header}`) as HTMLElement;
    const footer = pageEl.querySelector(`.${PAGE_CLASS_NAMES.footer}`) as HTMLElement;
    const content = pageEl.querySelector(`.${PAGE_CLASS_NAMES.content}`) as HTMLElement;

    expect(box(header).top).toBe(0);
    expect(box(header).bottom).toBeGreaterThan(0);
    expect(box(footer).bottom).toBe(1123);
    expect(box(footer).top).toBeLessThan(1123);

    // The strips overlap the body area, which still follows them in DOM order,
    // so only a z-index above the body's (auto) keeps them hit-testable.
    const order = Array.from(pageEl.children);
    expect(order.indexOf(header)).toBeLessThan(order.indexOf(content));
    expect(order.indexOf(footer)).toBeLessThan(order.indexOf(content));
    expect(content.style.zIndex).toBe('');
    expect(Number(header.style.zIndex)).toBeGreaterThan(0);
    expect(Number(footer.style.zIndex)).toBeGreaterThan(0);

    // Header/footer CONTENT is never lifted: at the same margin it stays behind.
    const withContent = renderPage(
      flush,
      { pageNumber: 1, totalPages: 1, section: 'body' },
      { document, headerContent: storyWithPicture(), footerContent: storyWithPicture() }
    );
    expect(
      (withContent.querySelector(`.${PAGE_CLASS_NAMES.header}`) as HTMLElement).style.zIndex
    ).toBe('');
    expect(
      (withContent.querySelector(`.${PAGE_CLASS_NAMES.footer}`) as HTMLElement).style.zIndex
    ).toBe('');
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

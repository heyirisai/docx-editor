/**
 * `renderPages` with `sectionHeaderFooters` paints each page's own section
 * header/footer — the first-page pair only on a `titlePg` section's first
 * page — and the section's own header distance.
 */

import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import type { Page, ParagraphBlock, ParagraphMeasure } from '../../layout-engine/types';
import { PAGE_CLASS_NAMES, renderPages, type HeaderFooterContent } from '../renderPage';
import type { SectionHeaderFooterContent } from '../renderPage/headerFooter';

beforeAll(() => GlobalRegistrator.register());
afterAll(() => GlobalRegistrator.unregister());

function story(text: string): HeaderFooterContent {
  const block: ParagraphBlock = {
    kind: 'paragraph',
    id: `hf-${text}`,
    runs: [{ kind: 'text', text }],
    attrs: {},
  };
  const measure: ParagraphMeasure = {
    kind: 'paragraph',
    lines: [
      {
        fromRun: 0,
        fromChar: 0,
        toRun: 0,
        toChar: text.length,
        width: text.length * 7,
        ascent: 11,
        descent: 4,
        lineHeight: 15,
      },
    ],
    totalHeight: 15,
  };
  return { blocks: [block], measures: [measure], height: 15 };
}

function page(number: number, sectionIndex: number, firstPageOfSection?: boolean): Page {
  return {
    number,
    fragments: [],
    margins: { top: 96, right: 96, bottom: 96, left: 96, header: 48, footer: 48 },
    size: { w: 816, h: 1056 },
    sectionIndex,
    ...(firstPageOfSection ? { firstPageOfSection } : {}),
  };
}

function headerText(pageEl: Element): string {
  return pageEl.querySelector(`.${PAGE_CLASS_NAMES.header}`)?.textContent ?? '';
}

describe('renderPages — per-section header/footer', () => {
  test("each page paints its own section's header", () => {
    const sections: SectionHeaderFooterContent[] = [
      { titlePg: false },
      { header: story('Section A'), footer: story('Footer A'), titlePg: false, headerDistance: 0 },
      {
        header: story('Section B'),
        firstPageHeader: story('Section B first'),
        titlePg: true,
      },
    ];
    const pages = [page(1, 0, true), page(2, 1, true), page(3, 1), page(4, 2, true), page(5, 2)];
    const container = document.createElement('div');
    renderPages(pages, container, {
      // The document-wide stories must not leak onto any page.
      headerContent: story('Document-wide'),
      footerContent: story('Document-wide footer'),
      sectionHeaderFooters: sections,
    });

    const pageEls = Array.from(container.querySelectorAll(`.${PAGE_CLASS_NAMES.page}`));
    expect(pageEls).toHaveLength(5);
    expect(pageEls.map(headerText)).toEqual([
      '',
      'Section A',
      'Section A',
      'Section B first',
      'Section B',
    ]);
    expect(
      pageEls.map((el) => el.querySelector(`.${PAGE_CLASS_NAMES.footer}`)?.textContent ?? '')
    ).toEqual(['', 'Footer A', 'Footer A', '', '']);

    // Section A states `w:header="0"`: its header sits at the page top.
    const headerA = pageEls[1].querySelector(`.${PAGE_CLASS_NAMES.header}`) as HTMLElement;
    const headerB = pageEls[4].querySelector(`.${PAGE_CLASS_NAMES.header}`) as HTMLElement;
    expect(headerA.style.top).toBe('0px');
    expect(headerB.style.top).toBe('48px');
  });
});

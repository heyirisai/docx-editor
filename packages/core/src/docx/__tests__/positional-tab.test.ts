/**
 * `<w:ptab>` — the absolute-position tab (ECMA-376 §17.3.3.19). Unlike
 * `<w:tab>` it does not walk the paragraph's tab stops: it names a boundary
 * and how the following text sits against it. Word footers use a centre
 * `w:ptab` then a right `w:ptab` to push the PAGE field to the right margin;
 * with the element unparsed, "Hilb Group" and the page number ran together.
 */

import { describe, expect, test } from 'bun:test';
import { parseDocumentBody } from '../documentParser';
import { serializeDocumentBody } from '../serializer/documentSerializer';
import { positionalTabStop } from '../../prosemirror/utils/tabCalculator';
import type { Paragraph, Run, TabContent } from '../../types/document';

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

function bodyWith(runs: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document ${NS}><w:body><w:p>${runs}</w:p>
      <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
    </w:body></w:document>`;
}

function tabsIn(xml: string): TabContent[] {
  const out: TabContent[] = [];
  for (const block of parseDocumentBody(xml).content) {
    if (block.type !== 'paragraph') continue;
    for (const item of (block as Paragraph).content) {
      if (item.type !== 'run') continue;
      for (const rc of (item as Run).content) {
        if (rc.type === 'tab') out.push(rc as TabContent);
      }
    }
  }
  return out;
}

describe('w:ptab parsing', () => {
  test('carries relativeTo / alignment / leader', () => {
    const tabs = tabsIn(
      bodyWith(
        '<w:r><w:ptab w:relativeTo="margin" w:alignment="right" w:leader="dot"/></w:r>' +
          '<w:r><w:tab/></w:r>'
      )
    );
    expect(tabs.length).toBe(2);
    expect(tabs[0].ptab).toEqual({ relativeTo: 'margin', alignment: 'right', leader: 'dot' });
    // A plain `w:tab` stays a plain tab.
    expect(tabs[1].ptab).toBeUndefined();
  });

  test('an unknown attribute value falls back instead of reaching layout', () => {
    const tabs = tabsIn(
      bodyWith('<w:r><w:ptab w:relativeTo="bogus" w:alignment="sideways" w:leader="zzz"/></w:r>')
    );
    expect(tabs[0].ptab).toEqual({ relativeTo: 'margin', alignment: 'left' });
  });

  test('round-trips as w:ptab, not as a plain w:tab', () => {
    const xml = serializeDocumentBody(
      parseDocumentBody(bodyWith('<w:r><w:ptab w:relativeTo="margin" w:alignment="center"/></w:r>'))
    );
    expect(xml).toContain('<w:ptab w:relativeTo="margin" w:alignment="center" w:leader="none"/>');
  });
});

describe('positionalTabStop', () => {
  const geometry = { contentWidthPx: 600, indentLeftPx: 40, indentRightPx: 20 };

  test('margin: left/centre/right of the whole text area', () => {
    expect(positionalTabStop({ relativeTo: 'margin', alignment: 'left' }, geometry).pos).toBe(0);
    expect(positionalTabStop({ relativeTo: 'margin', alignment: 'center' }, geometry).pos).toBe(
      4500 // 300px
    );
    const right = positionalTabStop({ relativeTo: 'margin', alignment: 'right' }, geometry);
    expect(right).toEqual({ val: 'end', pos: 9000, leader: undefined }); // 600px
  });

  test('indent: bounded by the paragraph indents instead', () => {
    expect(positionalTabStop({ relativeTo: 'indent', alignment: 'left' }, geometry).pos).toBe(600); // 40px
    expect(positionalTabStop({ relativeTo: 'indent', alignment: 'right' }, geometry).pos).toBe(
      8700 // 580px
    );
  });

  test('alignment maps onto the tab-stop model Word uses for w:tabs', () => {
    expect(positionalTabStop({ relativeTo: 'margin', alignment: 'left' }, geometry).val).toBe(
      'start'
    );
    expect(positionalTabStop({ relativeTo: 'margin', alignment: 'center' }, geometry).val).toBe(
      'center'
    );
    expect(positionalTabStop({ relativeTo: 'margin', alignment: 'right' }, geometry).val).toBe(
      'end'
    );
  });
});

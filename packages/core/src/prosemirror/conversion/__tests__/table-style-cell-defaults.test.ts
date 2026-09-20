/**
 * Where a table style's properties sit in the cascade, measured against Word:
 *
 *  - the style's OWN `w:tcPr` (§17.7.6.8) is the cell default for the whole
 *    table. A style that parks `<w:vAlign w:val="center"/>` there is how a
 *    tall merged label column gets its text centred; we read only the cell's
 *    own `w:tcPr`, so every such cell sat at the top.
 *
 *  - a table style's RUN properties sit between the document defaults and the
 *    paragraph style (§17.7.2). Above the paragraph style, a `firstRow`
 *    conditional repainted body text white that `Normal` had coloured; below
 *    the document defaults, it could never win at all.
 */

import { afterAll, beforeAll, describe, test, expect } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { toProseDoc } from '../toProseDoc';
import type { Document, Style, Table, TableCell } from '../../../types/document';

beforeAll(() => GlobalRegistrator.register());
afterAll(() => GlobalRegistrator.unregister());

const TABLE_STYLE: Style = {
  styleId: 'Branded',
  type: 'table',
  name: 'Branded',
  tcPr: { verticalAlign: 'center' },
  tblStylePr: [
    {
      type: 'firstRow',
      rPr: { color: { rgb: 'FFFFFF' } },
      tcPr: { shading: { fill: { rgb: '001B49' }, pattern: 'clear' } },
    },
  ],
};

/** `Normal` colours body text, as most branded templates do. */
const NORMAL: Style = {
  styleId: 'Normal',
  type: 'paragraph',
  name: 'Normal',
  default: true,
  rPr: { color: { rgb: '262626' } },
};

function cell(text: string): TableCell {
  return {
    type: 'tableCell',
    content: [{ type: 'paragraph', content: [{ type: 'run', content: [{ type: 'text', text }] }] }],
  };
}

const TABLE: Table = {
  type: 'table',
  formatting: {
    styleId: 'Branded',
    look: { firstRow: true, lastRow: false, firstColumn: false, lastColumn: false },
  },
  rows: [{ type: 'tableRow', cells: [cell('Header')] }],
};

function docWith(table: Table): Document {
  return {
    package: {
      document: { content: [table] },
      styles: { styles: [NORMAL, TABLE_STYLE] },
    },
  };
}

function firstCellNode(doc: Document) {
  const pm = toProseDoc(doc, { styles: doc.package.styles });
  return pm.child(0).child(0).child(0);
}

function firstRunColor(doc: Document): string | undefined {
  const textNode = firstCellNode(doc).child(0).child(0);
  return textNode.marks.find((m) => m.type.name === 'textColor')?.attrs.rgb as string | undefined;
}

describe("a table style's own w:tcPr", () => {
  test('supplies the cell default for vertical alignment', () => {
    expect(firstCellNode(docWith(TABLE)).attrs.verticalAlign).toBe('center');
  });

  test("the cell's own w:tcPr still wins", () => {
    const table: Table = {
      ...TABLE,
      rows: [
        {
          type: 'tableRow',
          cells: [{ ...cell('Header'), formatting: { verticalAlign: 'bottom' } }],
        },
      ],
    };
    expect(firstCellNode(docWith(table)).attrs.verticalAlign).toBe('bottom');
  });
});

describe('w:tblBorders cascade', () => {
  /**
   * Word's Table Grid: a full box plus inside rules, all hairline black.
   */
  const GRID: Style = {
    styleId: 'TableGrid',
    type: 'table',
    name: 'Table Grid',
    tblPr: {
      borders: {
        top: { style: 'single', size: 4, color: { auto: true } },
        left: { style: 'single', size: 4, color: { auto: true } },
        bottom: { style: 'single', size: 4, color: { auto: true } },
        right: { style: 'single', size: 4, color: { auto: true } },
        insideH: { style: 'single', size: 4, color: { auto: true } },
        insideV: { style: 'single', size: 4, color: { auto: true } },
      },
    },
  };

  /** Two columns so the first cell has an INSIDE edge on its right. */
  function twoColumnTable(direct?: Record<string, unknown>): Table {
    return {
      type: 'table',
      formatting: { styleId: 'TableGrid', ...(direct ?? {}) },
      rows: [{ type: 'tableRow', cells: [cell('A'), cell('B')] }],
    };
  }

  function docWithGrid(table: Table): Document {
    return {
      package: {
        document: { content: [table] },
        styles: { styles: [NORMAL, GRID] },
      },
    };
  }

  function firstCellBorders(doc: Document) {
    const pm = toProseDoc(doc, { styles: doc.package.styles });
    return pm.child(0).child(0).child(0).attrs.borders as
      | Record<string, { style?: string } | undefined>
      | undefined;
  }

  test("a table with no direct borders takes the style's", () => {
    const borders = firstCellBorders(docWithGrid(twoColumnTable()));
    expect(borders?.left?.style).toBe('single');
    expect(borders?.right?.style).toBe('single');
  });

  test("switching the OUTER box off keeps the style's inside rules", () => {
    // COMET's bio tables: Table Grid with `top/left/bottom/right = none` and
    // no `insideH`/`insideV` of their own. Word still draws the rule between
    // the photo column and the text; taking the whole `w:tblBorders` element
    // from the first level that had one dropped it.
    const none = { style: 'none' as const, size: 0, color: { auto: true } };
    const borders = firstCellBorders(
      docWithGrid(
        twoColumnTable({
          borders: { top: none, left: none, bottom: none, right: none },
        })
      )
    );
    expect(borders?.left?.style).toBe('none');
    expect(borders?.right?.style).toBe('single');
  });

  test("a direct insideV still overrides the style's", () => {
    const borders = firstCellBorders(
      docWithGrid(
        twoColumnTable({
          borders: { insideV: { style: 'double', size: 8, color: { rgb: 'FF0000' } } },
        })
      )
    );
    expect(borders?.right?.style).toBe('double');
    expect(borders?.left?.style).toBe('single');
  });
});

describe('table-style run properties in the cascade', () => {
  test("the paragraph style's colour wins over a firstRow conditional", () => {
    // #262626 from `Normal`, NOT the conditional's white — which is what
    // Word paints, and what stops header text going white-on-white when the
    // cell cancels the conditional's shading.
    expect(firstRunColor(docWith(TABLE))).toBe('262626');
  });

  test('but a document default does NOT — the table style outranks it', () => {
    const docDefaultsOnly: Document = {
      package: {
        document: { content: [TABLE] },
        styles: {
          docDefaults: { rPr: { color: { rgb: '000000' } } },
          styles: [
            { styleId: 'Normal', type: 'paragraph', name: 'Normal', default: true },
            TABLE_STYLE,
          ],
        },
      },
    };
    expect(firstRunColor(docDefaultsOnly)).toBe('FFFFFF');
  });

  test('the conditional still supplies what no run-property layer sets', () => {
    // Shading is a CELL property, so the conditional is the only source.
    expect(firstCellNode(docWith(TABLE)).attrs.backgroundColor).toBe('001B49');
  });
});

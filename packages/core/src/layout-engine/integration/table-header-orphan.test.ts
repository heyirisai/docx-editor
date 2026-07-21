/**
 * IF-1288 — Word never strands a table's header row(s) alone at the bottom
 * of a page: when only the header row fits in the remaining space, the
 * whole table starts on the next page. View-time pagination only — the
 * document (and export) is unchanged.
 */

import { describe, test, expect } from 'bun:test';

import { layoutDocument } from '../index';
import type {
  ParagraphBlock,
  ParagraphMeasure,
  TableBlock,
  TableFragment,
  TableMeasure,
} from '../types';
import { makeLayoutOptions } from './helpers';

const LINE = 20;

function para(id: number): ParagraphBlock {
  return {
    kind: 'paragraph',
    id,
    runs: [{ kind: 'text', text: 'x' }],
  } as unknown as ParagraphBlock;
}

function paraMeasure(lines: number): ParagraphMeasure {
  return {
    kind: 'paragraph',
    lines: Array.from({ length: lines }, () => ({
      fromRun: 0,
      fromChar: 0,
      toRun: 0,
      toChar: 0,
      width: 10,
      ascent: LINE * 0.8,
      descent: LINE * 0.2,
      lineHeight: LINE,
    })),
    totalHeight: lines * LINE,
  };
}

function buildHeaderTable(rowHeights: number[]): { block: TableBlock; measure: TableMeasure } {
  const block = {
    kind: 'table',
    id: 100,
    columnWidths: [200],
    rows: rowHeights.map((_, i) => ({
      id: i * 10,
      isHeader: i === 0,
      cells: [{ id: i * 10 + 1, blocks: [para(i * 10 + 1)] }],
    })),
  } as unknown as TableBlock;

  const measure: TableMeasure = {
    kind: 'table',
    columnWidths: [200],
    totalWidth: 200,
    totalHeight: rowHeights.reduce((a, b) => a + b, 0),
    rows: rowHeights.map((h) => ({
      height: h,
      cells: [{ blocks: [paraMeasure(1)], width: 200, height: LINE }],
    })),
  };

  return { block, measure };
}

describe('Layout engine — table header orphan rule', () => {
  test('a header row that would sit alone at the page bottom moves with the table to the next page', () => {
    const options = makeLayoutOptions();
    const contentHeight = options.pageSize.h - options.margins.top - options.margins.bottom;

    // Filler paragraph leaves room for exactly one table row.
    const fillerLines = Math.floor((contentHeight - LINE) / LINE);
    const filler = para(1);
    const fillerMeasure = paraMeasure(fillerLines);

    const { block, measure } = buildHeaderTable([LINE, LINE, LINE]);
    const layout = layoutDocument([filler, block], [fillerMeasure, measure], options);

    const tableFrags: Array<{ page: number; frag: TableFragment }> = [];
    layout.pages.forEach((page, pageIndex) => {
      for (const frag of page.fragments) {
        if (frag.kind === 'table')
          tableFrags.push({ page: pageIndex, frag: frag as TableFragment });
      }
    });

    expect(tableFrags.length).toBeGreaterThan(0);
    // The FIRST table fragment starts on page 2 (index 1) and carries the
    // header row together with data rows — never a lone header on page 1.
    expect(tableFrags[0].page).toBe(1);
    expect(tableFrags[0].frag.fromRow).toBe(0);
    expect(tableFrags[0].frag.toRow).toBeGreaterThan(1);
  });

  test('a mid-row split suppresses the repeated header on the continuation (Word quirk)', () => {
    const options = makeLayoutOptions();
    const contentHeight = options.pageSize.h - options.margins.top - options.margins.bottom;

    // Header + one data row far taller than a page: the data row must split
    // mid-content, and Word does NOT repeat the header above the remainder.
    const tallRow = contentHeight * 1.5;
    const { block, measure } = buildHeaderTable([LINE, tallRow]);
    measure.rows[1].cells[0].blocks = [paraMeasure(Math.ceil(tallRow / LINE))];
    measure.rows[1].cells[0].height = tallRow;
    measure.totalHeight = LINE + tallRow;

    const layout = layoutDocument([block], [measure], options);
    const continuations: TableFragment[] = [];
    for (const page of layout.pages) {
      for (const frag of page.fragments) {
        if (frag.kind === 'table' && (frag as TableFragment).continuesFromPrev) {
          continuations.push(frag as TableFragment);
        }
      }
    }
    expect(continuations.length).toBeGreaterThan(0);
    for (const frag of continuations) {
      if (frag.topClip && frag.topClip > 0) {
        expect(frag.headerRowCount).toBeUndefined();
      }
    }
  });

  test('a headerless table still fills the remaining space greedily', () => {
    const options = makeLayoutOptions();
    const contentHeight = options.pageSize.h - options.margins.top - options.margins.bottom;
    const fillerLines = Math.floor((contentHeight - LINE) / LINE);

    const { block, measure } = buildHeaderTable([LINE, LINE, LINE]);
    (block.rows[0] as { isHeader?: boolean }).isHeader = false;

    const layout = layoutDocument([para(1), block], [paraMeasure(fillerLines), measure], options);

    const firstTablePage = layout.pages.findIndex((p) =>
      p.fragments.some((f) => f.kind === 'table')
    );
    // Without a header row the first row may sit at the bottom of page 1.
    expect(firstTablePage).toBe(0);
  });
});

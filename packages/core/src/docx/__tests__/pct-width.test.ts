import { describe, expect, test } from 'bun:test';
import type { Node as PMNode } from 'prosemirror-model';
import { tableCellSpec } from '../../prosemirror/extensions/nodes/TableExtension/specs';
import type { Document, Table, TableCell } from '../../types/document';
import { toProseDoc } from '../../prosemirror/conversion/toProseDoc';
import { parseTableMeasurement } from '../tableParser/properties';
import { parseXml, type XmlElement } from '../xmlParser';

function measurement(attrs: string) {
  const doc = parseXml(
    `<w:tblW xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ${attrs}/>`
  );
  return parseTableMeasurement((doc.elements as XmlElement[])[0]);
}

describe('pct table width', () => {
  test('a literal percent is stored as fiftieths of a percent', () => {
    expect(measurement('w:w="100%" w:type="pct"')).toEqual({ value: 5000, type: 'pct' });
    expect(measurement('w:type="pct" w:w="50%"')).toEqual({ value: 2500, type: 'pct' });
  });

  test('a bare fiftieths value is left alone', () => {
    expect(measurement('w:w="5000" w:type="pct"')).toEqual({ value: 5000, type: 'pct' });
    expect(measurement('w:w="2500" w:type="pct"')).toEqual({ value: 2500, type: 'pct' });
  });

  test('dxa widths are twips, even when the attribute looks like a percent', () => {
    expect(measurement('w:w="2800" w:type="dxa"')).toEqual({ value: 2800, type: 'dxa' });
  });

  test('a literal percent on a cell margin stays the parsed number', () => {
    const doc = parseXml(
      `<w:top xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:w="10%" w:type="pct"/>`
    );
    expect(parseTableMeasurement((doc.elements as XmlElement[])[0])).toEqual({
      value: 10,
      type: 'pct',
    });
  });

  test('a cell width in fiftieths is a CSS percent', () => {
    const dom = tableCellSpec.toDOM?.({
      attrs: { width: 2500, widthType: 'pct', colspan: 1, rowspan: 1, noWrap: false },
    } as never) as [string, { style: string }, number];
    expect(dom[1].style).toContain('width: 50%');
  });

  test('a grid-derived cell width uses fiftieths, so an 80/20 grid stays 80/20', () => {
    const cell = (text: string): TableCell => ({
      type: 'tableCell',
      content: [
        { type: 'paragraph', content: [{ type: 'run', content: [{ type: 'text', text }] }] },
      ],
    });
    const table: Table = {
      type: 'table',
      columnWidths: [8000, 2000],
      rows: [{ type: 'tableRow', cells: [cell('Wide'), cell('Narrow')] }],
    };
    const doc: Document = { package: { document: { content: [table] } } };
    let pmTable: PMNode | undefined;
    toProseDoc(doc).descendants((node) => {
      if (node.type.name === 'table') {
        pmTable = node;
        return false;
      }
      return true;
    });
    if (!pmTable) throw new Error('expected a table');
    const wide = pmTable.child(0).child(0);
    const narrow = pmTable.child(0).child(1);
    expect(wide.attrs.width).toBe(4000);
    expect(narrow.attrs.width).toBe(1000);
    const wideStyle = (tableCellSpec.toDOM?.(wide) as [string, { style: string }])[1].style;
    const narrowStyle = (tableCellSpec.toDOM?.(narrow) as [string, { style: string }])[1].style;
    expect(wideStyle).toContain('width: 80%');
    expect(narrowStyle).toContain('width: 20%');
  });
});

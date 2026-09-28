/**
 * Upstream-parity guard for the `fixTables` port in `gridAwareTableEditing`.
 *
 * `fixTablesRespectingGrid` replaces prosemirror-tables' own repair pass, so an
 * upstream fix or behaviour change would otherwise never reach us and no test
 * would notice. For tables WITHOUT a covering `w:tblGrid` the port is meant to
 * be identical to upstream, so these tests run both over the same malformed
 * shapes and require the same repaired doc. The version pin makes any
 * prosemirror-tables upgrade a deliberate change that re-reads the port.
 */

import { describe, test, expect } from 'bun:test';
import { createRequire } from 'node:module';
import { EditorState } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import { fixTables } from 'prosemirror-tables';
import { singletonManager } from '../../../schema';
import { fixTablesRespectingGrid } from './gridAwareTableEditing';

const schema = singletonManager.getSchema();

/**
 * The prosemirror-tables version `fixTableRespectingGrid` / `changedDescendants`
 * were ported from. When this fails after an upgrade: diff upstream's
 * `fixTable` and `changedDescendants` against the port, carry over any change,
 * then update this constant.
 */
const PORTED_FROM_VERSION = '1.8.5';

type CellSpec = { text?: string; colspan?: number; rowspan?: number; colwidth?: number[] | null };

function cell(spec: CellSpec = {}): PMNode {
  const { text = 'x', colspan = 1, rowspan = 1, colwidth = null } = spec;
  return schema.nodes.tableCell.create(
    { colspan, rowspan, colwidth },
    schema.nodes.paragraph.create(null, text ? schema.text(text) : null)
  );
}

function table(rows: CellSpec[][], columnWidths: number[] | null = null): PMNode {
  return schema.nodes.table.create(
    { columnWidths },
    rows.map((r) => schema.nodes.tableRow.create(null, r.map(cell)))
  );
}

function docWith(t: PMNode): PMNode {
  return schema.nodes.doc.create(null, [schema.nodes.paragraph.create(), t]);
}

function repairBoth(doc: PMNode): { ours: PMNode; upstream: PMNode } {
  const state = EditorState.create({ schema, doc });
  const ours = fixTablesRespectingGrid(state);
  const upstream = fixTables(state);
  return { ours: (ours ?? state.tr).doc, upstream: (upstream ?? state.tr).doc };
}

describe('gridAwareTableEditing: parity with upstream fixTables', () => {
  test('installed prosemirror-tables is the version the port was taken from', () => {
    const require = createRequire(import.meta.url);
    const { version } = require('prosemirror-tables/package.json') as { version: string };
    expect(version).toBe(PORTED_FROM_VERSION);
  });

  const shapes: Array<[string, PMNode]> = [
    [
      'collision (colspan overlapping a rowspan)',
      table([
        [{ rowspan: 2 }, {}, {}],
        [{ colspan: 2 }, {}],
      ]),
    ],
    ['overlong rowspan', table([[{ rowspan: 3 }, {}], [{}]])],
    [
      'colwidth mismatch',
      table([
        [{ colwidth: [100] }, { colwidth: [200] }],
        [{ colwidth: [150] }, { colwidth: [200] }],
      ]),
    ],
    ['short first row (padded at the start)', table([[{}], [{}, {}, {}]])],
    ['short middle row (padded at the end)', table([[{}, {}, {}], [{}], [{}, {}, {}]])],
    ['short row with a grid too narrow to cover it', table([[{}, {}, {}], [{}]], [1000, 1000])],
  ];

  for (const [name, t] of shapes) {
    test(name, () => {
      const doc = docWith(t);
      const { ours, upstream } = repairBoth(doc);
      // Guard against a vacuous pass: every shape must actually need repair.
      expect(upstream.eq(doc)).toBe(false);
      expect(ours.toJSON()).toEqual(upstream.toJSON());
    });
  }

  test('a well-formed table is left alone by both', () => {
    const doc = docWith(
      table([
        [{}, {}],
        [{}, {}],
      ])
    );
    const state = EditorState.create({ schema, doc });
    expect(fixTablesRespectingGrid(state)).toBeUndefined();
    expect(fixTables(state)).toBeUndefined();
  });

  test('diverges from upstream only for a short row inside a covering w:tblGrid', () => {
    const doc = docWith(
      table(
        [
          [{}, {}, {}],
          [{}, {}],
        ],
        [1000, 1000, 1000]
      )
    );
    const state = EditorState.create({ schema, doc });
    expect(fixTablesRespectingGrid(state)).toBeUndefined();
    expect(fixTables(state)?.doc.eq(doc)).toBe(false);
  });
});

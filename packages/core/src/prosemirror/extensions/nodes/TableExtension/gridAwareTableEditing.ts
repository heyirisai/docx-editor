/**
 * `tableEditing()` from prosemirror-tables, minus one normalisation that
 * breaks Word tables: padding short rows with empty cells.
 *
 * prosemirror-tables assumes every row covers the same number of columns and,
 * on every doc-changing transaction, `fixTables` "repairs" a short row by
 * inserting empty cells — at the START of the row when it is the only short
 * row and the first one (`fixTable`'s `side` rule). In OOXML a short row is
 * legal: its cells simply don't cover the whole `w:tblGrid` (`w:gridBefore` /
 * `w:gridAfter`, §17.4.14-15, or an implicit trailing gap) and Word draws
 * nothing in the uncovered grid columns. Padding such a row shifts every cell
 * one grid column over, so a header row laid out on a grid with hairline
 * spacer columns lands its cells on the wrong (hairline) columns.
 *
 * This plugin runs the same repair pass but leaves a `missing` problem alone
 * when the table carries a `w:tblGrid` (`columnWidths`) wide enough to hold the
 * widest row — the grid, not the widest row, defines the columns, so the row is
 * short by design. Every other problem (collisions, overlong rowspans, colwidth
 * mismatches, zero-sized tables, and short rows in grid-less tables) is fixed
 * exactly as upstream does.
 *
 * `fixTableRespectingGrid` is a port of prosemirror-tables' `fixTable`
 * (MIT, © Marijn Haverbeke and others); keep it in sync on upgrade.
 */

import { Plugin, type EditorState, type Transaction } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import {
  TableMap,
  fixTablesKey,
  removeColSpan,
  tableEditing,
  tableNodeTypes,
  type TableRole,
} from 'prosemirror-tables';

type CellAttrs = Parameters<typeof removeColSpan>[0];

/** True when the table's `w:tblGrid` defines at least as many columns as the widest row. */
function hasCoveringGrid(table: PMNode, map: TableMap): boolean {
  const grid = table.attrs.columnWidths as unknown;
  return Array.isArray(grid) && grid.length >= map.width;
}

function fixTableRespectingGrid(
  state: EditorState,
  table: PMNode,
  tablePos: number,
  tr: Transaction | undefined
): Transaction | undefined {
  const map = TableMap.get(table);
  if (!map.problems) return tr;
  const keepShortRows = hasCoveringGrid(table, map);
  const problems = keepShortRows ? map.problems.filter((p) => p.type !== 'missing') : map.problems;
  if (problems.length === 0) return tr;
  if (!tr) tr = state.tr;

  const mustAdd: number[] = [];
  for (let i = 0; i < map.height; i++) mustAdd.push(0);
  for (const prob of problems) {
    if (prob.type === 'collision') {
      const cell = table.nodeAt(prob.pos);
      if (!cell) continue;
      const attrs = cell.attrs as CellAttrs;
      for (let j = 0; j < attrs.rowspan; j++) mustAdd[prob.row + j] += prob.n;
      tr.setNodeMarkup(
        tr.mapping.map(tablePos + 1 + prob.pos),
        null,
        removeColSpan(attrs, attrs.colspan - prob.n, prob.n)
      );
    } else if (prob.type === 'missing') {
      mustAdd[prob.row] += prob.n;
    } else if (prob.type === 'overlong_rowspan') {
      const cell = table.nodeAt(prob.pos);
      if (!cell) continue;
      tr.setNodeMarkup(tr.mapping.map(tablePos + 1 + prob.pos), null, {
        ...cell.attrs,
        rowspan: cell.attrs.rowspan - prob.n,
      });
    } else if (prob.type === 'colwidth mismatch') {
      const cell = table.nodeAt(prob.pos);
      if (!cell) continue;
      tr.setNodeMarkup(tr.mapping.map(tablePos + 1 + prob.pos), null, {
        ...cell.attrs,
        colwidth: prob.colwidth,
      });
    } else if (prob.type === 'zero_sized') {
      const pos = tr.mapping.map(tablePos);
      tr.delete(pos, pos + table.nodeSize);
    }
  }

  let first: number | undefined;
  let last: number | undefined;
  for (let i = 0; i < mustAdd.length; i++) {
    if (mustAdd[i]) {
      if (first == null) first = i;
      last = i;
    }
  }
  for (let i = 0, pos = tablePos + 1; i < map.height; i++) {
    const row = table.child(i);
    const end = pos + row.nodeSize;
    const add = mustAdd[i];
    if (add > 0) {
      let role: TableRole = 'cell';
      if (row.firstChild) role = row.firstChild.type.spec.tableRole as TableRole;
      const nodes: PMNode[] = [];
      for (let j = 0; j < add; j++) {
        const node = tableNodeTypes(state.schema)[role].createAndFill();
        if (node) nodes.push(node);
      }
      const side = (i === 0 || first === i - 1) && last === i ? pos + 1 : end - 1;
      tr.insert(tr.mapping.map(side), nodes);
    }
    pos = end;
  }
  return tr.setMeta(fixTablesKey, { fixTables: true });
}

/** Port of prosemirror-tables' `changedDescendants` (not exported). */
function changedDescendants(
  old: PMNode,
  cur: PMNode,
  offset: number,
  f: (node: PMNode, pos: number) => void
): void {
  const oldSize = old.childCount;
  const curSize = cur.childCount;
  outer: for (let i = 0, j = 0; i < curSize; i++) {
    const child = cur.child(i);
    for (let scan = j, e = Math.min(oldSize, i + 3); scan < e; scan++) {
      if (old.child(scan) === child) {
        j = scan + 1;
        offset += child.nodeSize;
        continue outer;
      }
    }
    f(child, offset);
    if (j < oldSize && old.child(j).sameMarkup(child)) {
      changedDescendants(old.child(j), child, offset + 1, f);
    } else {
      child.nodesBetween(0, child.content.size, f, offset + 1);
    }
    offset += child.nodeSize;
  }
}

/**
 * `fixTables` that keeps rows legitimately short of their `w:tblGrid`.
 *
 * @internal
 */
export function fixTablesRespectingGrid(
  state: EditorState,
  oldState?: EditorState
): Transaction | undefined {
  let tr: Transaction | undefined;
  const check = (node: PMNode, pos: number) => {
    if (node.type.spec.tableRole === 'table') {
      tr = fixTableRespectingGrid(state, node, pos, tr);
    }
  };
  if (!oldState) state.doc.descendants(check);
  else if (oldState.doc !== state.doc) changedDescendants(oldState.doc, state.doc, 0, check);
  return tr;
}

/**
 * prosemirror-tables' `tableEditing()` with the grid-respecting table repair.
 *
 * The upstream plugin's `appendTransaction` is `normalizeSelection(state,
 * fixTables(state, oldState))`. We run our repair first, then hand upstream a
 * state in which the doc is unchanged (so its own `fixTables` is a no-op) and
 * let it normalise the selection exactly as before.
 *
 * @internal
 */
export function gridAwareTableEditing(): Plugin {
  const base = tableEditing();
  const baseAppend = base.spec.appendTransaction!;
  return new Plugin({
    ...base.spec,
    appendTransaction(trs, oldState, state) {
      const fix = fixTablesRespectingGrid(state, oldState);
      if (!fix) {
        // Doc unchanged from upstream's point of view → only selection
        // normalisation runs.
        return baseAppend.call(base, trs, state, state);
      }
      // Upstream normalises the selection of the post-repair doc; it only ever
      // sets a selection (no steps), so graft that onto the repair.
      const fixedState = state.apply(fix);
      const normalize = baseAppend.call(base, [fix], fixedState, fixedState);
      if (normalize?.selectionSet) fix.setSelection(normalize.selection);
      return fix;
    },
  });
}

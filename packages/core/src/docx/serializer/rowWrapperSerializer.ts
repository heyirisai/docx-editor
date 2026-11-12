/**
 * Row-level cell wrappers (`w:tr > w:sdt > w:sdtContent > w:tc`,
 * `w:tr > w:customXml > w:tc`) — ECMA-376 `CT_Row` / `CT_SdtContentCell` /
 * `CT_CustomXmlCell`. See `TableCell.rowWrappers`.
 */

import type { TableCell, TableRow } from '../../types/document';
import { getChildElements, isWellFormedXmlElement, parseXml, type XmlElement } from '../xmlParser';

/** `w:customXmlPr` children — ECMA-376 `CT_CustomXmlPr` (both empty elements). */
const CUSTOM_XML_PR_CHILDREN = new Set(['w:placeholder', 'w:attr']);

/** `el`'s element children, or null if it holds non-whitespace character data. */
function elementChildrenOnly(el: XmlElement): XmlElement[] | null {
  for (const node of el.elements ?? []) {
    if (node.type === 'element') continue;
    if (String(node.text ?? node.cdata ?? '').trim() !== '') return null;
  }
  return getChildElements(el);
}

/**
 * Whether a row-level `w:customXml` wrapper's captured tags are safe to write
 * back verbatim: together they must form one well-formed `w:customXml` element
 * holding at most a `w:customXmlPr` of empty `w:placeholder` / `w:attr`
 * children (`CT_CustomXmlRow` minus its cells). The parser only ever captures
 * that shape, but `TableCell.rowWrappers` is public model state, so anything
 * else (an early close, injected rows or runs) is rejected and the wrapper
 * dropped rather than interpolated into `document.xml`.
 */
export function isTrustedRowCustomXml(startXml: string, endXml: string): boolean {
  if (typeof startXml !== 'string' || !startXml.startsWith('<w:customXml')) return false;
  if (endXml !== '</w:customXml>') return false;
  const xml = `${startXml}${endXml}`;
  if (!isWellFormedXmlElement(xml)) return false;
  let root: XmlElement | undefined;
  try {
    root = parseXml(xml).elements?.find((e) => e.type === 'element');
  } catch {
    return false;
  }
  if (root?.name !== 'w:customXml') return false;
  const children = elementChildrenOnly(root);
  if (!children || children.length > 1) return false;
  const pr = children[0];
  if (!pr) return true;
  if (pr.name !== 'w:customXmlPr') return false;
  const prChildren = elementChildrenOnly(pr);
  return (
    !!prChildren &&
    prChildren.every(
      (el) => CUSTOM_XML_PR_CHILDREN.has(el.name ?? '') && elementChildrenOnly(el)?.length === 0
    )
  );
}

/** A row-level wrapper to emit around a cell (see `TableCell.rowWrappers`). */
interface PlannedRowWrapper {
  key: string;
  open: string;
  close: string;
  /** A control's first cell: always opens a fresh `w:sdt`. */
  fresh: boolean;
  /** A later cell of a control: can only join it while it is still open. */
  continuation: boolean;
}

/**
 * Serialize a row's cells, re-emitting the row-level `w:sdt` / `w:customXml`
 * wrappers they were parsed from (`w:tr > w:sdt > w:sdtContent > w:tc`).
 *
 * A control's leading cell holds the control as its single {@link BlockSdt};
 * that block is peeled back off and its (possibly answered) properties are
 * written as the wrapper around the `w:tc`. Consecutive cells sharing a wrapper
 * go back inside the same element. If editing removed the control from the
 * cell, its wrapper is dropped rather than emitted without properties.
 */
export function serializeRowCells(
  cells: TableCell[],
  serializeTableCell: (cell: TableCell) => string
): string {
  const out: string[] = [];
  const open: PlannedRowWrapper[] = [];

  for (const cell of cells) {
    const chain = cell.rowWrappers;
    if (!chain || chain.length === 0) {
      while (open.length > 0) out.push(open.pop()!.close);
      out.push(serializeTableCell(cell));
      continue;
    }

    let content = cell.content;
    const planned: PlannedRowWrapper[] = [];
    for (const w of chain) {
      if (w.kind === 'customXml') {
        // Untrusted tags: drop the wrapper, keep the cell (see isTrustedRowCustomXml).
        if (!isTrustedRowCustomXml(w.startXml, w.endXml)) continue;
        planned.push({
          key: `x${w.id}`,
          open: w.startXml,
          close: w.endXml,
          fresh: false,
          continuation: false,
        });
      } else if (w.leading) {
        const sdt = content.length === 1 && content[0]!.type === 'blockSdt' ? content[0] : null;
        if (!sdt) continue;
        const props = sdt.properties;
        planned.push({
          key: `s${w.id}`,
          open: `<w:sdt>${props.rawPropertiesXml ?? ''}${props.rawEndPropertiesXml ?? ''}<w:sdtContent>`,
          close: '</w:sdtContent></w:sdt>',
          fresh: true,
          continuation: false,
        });
        content = sdt.content as TableCell['content'];
      } else {
        planned.push({
          key: `s${w.id}`,
          open: '',
          close: '</w:sdtContent></w:sdt>',
          fresh: false,
          continuation: true,
        });
      }
    }

    let common = 0;
    while (
      common < open.length &&
      common < planned.length &&
      !planned[common]!.fresh &&
      open[common]!.key === planned[common]!.key
    ) {
      common++;
    }
    while (open.length > common) out.push(open.pop()!.close);
    for (const w of planned.slice(common)) {
      // A continuation whose control is no longer open has nothing to join.
      if (w.continuation) continue;
      out.push(w.open);
      open.push(w);
    }

    out.push(serializeTableCell(content === cell.content ? cell : { ...cell, content }));
  }

  while (open.length > 0) out.push(open.pop()!.close);
  return out.join('');
}

/** `xml` is absent, or one well-formed `<w:${name}>` element. */
function isTrustedSdtPart(xml: string | undefined, name: string): boolean {
  if (xml === undefined || xml === '') return true;
  if (typeof xml !== 'string' || !new RegExp(`^<w:${name}[\\s/>]`).test(xml)) return false;
  if (!isWellFormedXmlElement(xml)) return false;
  try {
    const roots = (parseXml(xml).elements ?? []).filter((e) => e.type === 'element');
    return roots.length === 1 && roots[0]!.name === `w:${name}`;
  } catch {
    return false;
  }
}

/**
 * Serialize a table's rows, re-emitting the table-level `w:sdt` /
 * `w:customXml` wrappers they were parsed from
 * (`w:tbl > w:sdt > w:sdtContent > w:tr`, ECMA-376 `CT_SdtRow` /
 * `CT_CustomXmlRow`). See `TableRow.tableWrappers`.
 *
 * Consecutive rows sharing a wrapper go back inside the same element. A
 * wrapper that was already closed (an unwrapped row was inserted inside a
 * multi-row control) is not reopened, so a control is never duplicated.
 * Wrappers whose captured tags are not the shape the parser produces are
 * dropped and their rows kept.
 */
export function serializeTableRows(
  rows: TableRow[],
  serializeTableRow: (row: TableRow) => string
): string {
  const out: string[] = [];
  const open: { key: string; close: string }[] = [];
  const closed = new Set<string>();
  const closeTo = (depth: number) => {
    while (open.length > depth) {
      const w = open.pop()!;
      closed.add(w.key);
      out.push(w.close);
    }
  };

  for (const row of rows) {
    const planned: { key: string; open: string; close: string }[] = [];
    for (const w of row.tableWrappers ?? []) {
      if (w.kind === 'customXml') {
        if (!isTrustedRowCustomXml(w.startXml, w.endXml)) continue;
        planned.push({ key: `x${w.id}`, open: w.startXml, close: w.endXml });
      } else {
        const pr = w.properties?.rawPropertiesXml;
        const endPr = w.properties?.rawEndPropertiesXml;
        if (!isTrustedSdtPart(pr, 'sdtPr') || !isTrustedSdtPart(endPr, 'sdtEndPr')) continue;
        planned.push({
          key: `s${w.id}`,
          open: `<w:sdt>${pr ?? ''}${endPr ?? ''}<w:sdtContent>`,
          close: '</w:sdtContent></w:sdt>',
        });
      }
    }

    let common = 0;
    while (
      common < open.length &&
      common < planned.length &&
      open[common]!.key === planned[common]!.key
    ) {
      common++;
    }
    closeTo(common);
    for (const w of planned.slice(common)) {
      if (closed.has(w.key)) break;
      out.push(w.open);
      open.push(w);
    }
    out.push(serializeTableRow(row));
  }

  closeTo(0);
  return out.join('');
}

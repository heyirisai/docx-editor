/**
 * Tables (`w:tbl`), rows (`w:tr`), and cells (`w:tc`).
 */

import type { TableFormatting, TableRowFormatting, TableCellFormatting } from '../formatting';
import type { Paragraph } from './paragraph';
import type { BlockSdt } from './sdt';
import type {
  TablePropertyChange,
  TableRowPropertyChange,
  TableCellPropertyChange,
  TableStructuralChangeInfo,
} from './trackedChange';

/**
 * Table cell (`w:tc`). Holds nested block content (paragraphs and nested
 * tables), cell-level formatting (borders, shading, vertical merge),
 * tracked property changes, and tracked structural changes for cell
 * insert/delete/merge operations.
 */
export interface TableCell {
  type: 'tableCell';
  /** Cell formatting */
  formatting?: TableCellFormatting;
  /** Cell-level tracked property changes (w:tcPrChange) */
  propertyChanges?: TableCellPropertyChange[];
  /** Tracked structural changes (cell insert/delete/merge) */
  structuralChange?: TableStructuralChangeInfo;
  /**
   * Cell content (`EG_BlockLevelElts`): paragraphs, nested tables, and
   * block-level content controls (`w:sdt` directly inside `w:tc`, e.g. a
   * Yes/No drop-down wrapping the answer paragraph).
   */
  content: (Paragraph | Table | BlockSdt)[];
  /**
   * Row-level wrappers this cell sat inside, outermost first. `CT_Row` lets a
   * cell appear inside a `w:sdt` (`w:tr > w:sdt > w:sdtContent > w:tc`, a
   * content control wrapping a whole cell) or a `w:customXml`
   * (`w:tr > w:customXml > w:tc`) instead of directly under the row.
   *
   * For each `sdt` wrapper whose `leading` cell this is, `content` is one
   * {@link BlockSdt} carrying that control's properties and holding the cell's
   * real blocks (so the control is found and answered like any cell-level
   * control); the serializer peels it back off and re-emits the wrapper around
   * the `w:tc`. Absent for ordinary cells.
   */
  rowWrappers?: TableCellRowWrapper[];
}

/**
 * One row-level wrapper around a table cell (see {@link TableCell.rowWrappers}).
 * `id` is unique within the row; consecutive cells sharing an `id` sat in the
 * same wrapper (`CT_SdtContentCell` / `CT_CustomXmlCell` can hold several
 * `w:tc`).
 */
export type TableCellRowWrapper =
  | {
      kind: 'sdt';
      id: number;
      /**
       * True for the first cell inside the control: its `content` holds the
       * control's {@link BlockSdt}. Later cells of the same control carry
       * `leading: false` and their plain content.
       */
      leading: boolean;
    }
  | {
      kind: 'customXml';
      id: number;
      /** Verbatim start tag plus `w:customXmlPr`, e.g. `<w:customXml w:element="x"><w:customXmlPr/>`. */
      startXml: string;
      /** Verbatim end tag, e.g. `</w:customXml>`. */
      endXml: string;
    };

/**
 * Table row (`w:tr`) — an ordered list of `TableCell` plus row-level
 * formatting (height, repeated header, cantSplit) and tracked changes
 * for inserts/deletes.
 */
export interface TableRow {
  type: 'tableRow';
  /** Row formatting */
  formatting?: TableRowFormatting;
  /** Row-level tracked property changes (w:trPrChange) */
  propertyChanges?: TableRowPropertyChange[];
  /** Tracked structural changes (row insert/delete) */
  structuralChange?: TableStructuralChangeInfo;
  /** Cells in this row */
  cells: TableCell[];
}

/**
 * Table (`w:tbl`) — a block-level grid of rows × cells. Tables carry
 * their own formatting layer (borders, shading, alignment, indent,
 * floating placement) and an explicit column-width grid in twips. Tables
 * can nest arbitrarily through `TableCell.content`.
 *
 * See ECMA-376 §17.4.
 */
export interface Table {
  type: 'table';
  /** Table formatting */
  formatting?: TableFormatting;
  /** Table-level tracked property changes (w:tblPrChange) */
  propertyChanges?: TablePropertyChange[];
  /** Column widths in twips */
  columnWidths?: number[];
  /** Table rows */
  rows: TableRow[];
}

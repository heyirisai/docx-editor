/**
 * Measurement results for flow blocks.
 *
 * Split out of `types.ts` so that file stays under its line budget; every name
 * here is re-exported from `types.ts`, which remains the entry point the
 * engine, the adapters' `measureBlock`, and the painter import from.
 */

/**
 * A measured line within a paragraph.
 */
export type MeasuredLine = {
  /** Starting run index (inclusive). */
  fromRun: number;
  /** Starting character index within fromRun. */
  fromChar: number;
  /** Ending run index (inclusive). */
  toRun: number;
  /** Ending character index within toRun (exclusive). */
  toChar: number;
  /** Total width of the line in pixels. */
  width: number;
  /** Ascent (height above baseline) in pixels. */
  ascent: number;
  /** Descent (height below baseline) in pixels. */
  descent: number;
  /** Total line height in pixels. */
  lineHeight: number;
  /** Left offset from floating images (pixels from content left edge). */
  leftOffset?: number;
  /** Right offset from floating images (pixels from content right edge). */
  rightOffset?: number;
  /** Optional split segments for centered floating exclusions. */
  segments?: MeasuredLineSegment[];
  /**
   * Vertical space inserted before this line to skip past floats that leave
   * no usable horizontal width at the natural line Y. Painters render this
   * as marginTop on the line element; measurement adds it to totalHeight.
   */
  floatSkipBefore?: number;
};

export type MeasuredLineSegment = {
  fromRun: number;
  fromChar: number;
  toRun: number;
  toChar: number;
  width: number;
  leftOffset: number;
  availableWidth: number;
};

/**
 * Measurement result for a paragraph block.
 */
export type ParagraphMeasure = {
  kind: 'paragraph';
  lines: MeasuredLine[];
  totalHeight: number;
};

/**
 * Measurement result for an image block.
 */
export type ImageMeasure = {
  kind: 'image';
  width: number;
  height: number;
};

/**
 * Measurement result for a table cell.
 */
export type TableCellMeasure = {
  blocks: Measure[];
  width: number;
  height: number;
  colSpan?: number;
  rowSpan?: number;
};

/**
 * Measurement result for a table row.
 */
export type TableRowMeasure = {
  cells: TableCellMeasure[];
  height: number;
};

/**
 * Measurement result for a table block.
 */
export type TableMeasure = {
  kind: 'table';
  rows: TableRowMeasure[];
  columnWidths: number[];
  totalWidth: number;
  totalHeight: number;
};

/**
 * Measurement result for section break (no visual size).
 */
export type SectionBreakMeasure = {
  kind: 'sectionBreak';
};

/**
 * Measurement result for page break (no visual size).
 */
export type PageBreakMeasure = {
  kind: 'pageBreak';
};

/**
 * Measurement result for column break (no visual size).
 */
export type ColumnBreakMeasure = {
  kind: 'columnBreak';
};

/**
 * Measurement result for a text box block.
 */
export type TextBoxMeasure = {
  kind: 'textBox';
  width: number;
  height: number;
  /**
   * Pre-measured inner block measures, 1:1 with `TextBoxBlock.content`
   * (avoids re-measuring during render). A table inside the box measures as
   * a `TableMeasure`.
   */
  innerMeasures: Array<ParagraphMeasure | TableMeasure>;
};

/**
 * Union of all measurement types.
 */
export type Measure =
  | ParagraphMeasure
  | ImageMeasure
  | TableMeasure
  | TextBoxMeasure
  | SectionBreakMeasure
  | PageBreakMeasure
  | ColumnBreakMeasure;

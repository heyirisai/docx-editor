/**
 * Paragraph property parser (w:pPr → ParagraphFormatting).
 *
 * Owns parseParagraphProperties plus its two leaf helpers (parseBorderSpec
 * and parseTabStops). Shading, color, and run-default parsing live in
 * runProperties.ts and are imported here.
 */

import type {
  Theme,
  ParagraphFormatting,
  BorderSpec,
  TabStop,
  TabStopAlignment,
  TabLeader,
} from '../../types/document';
import {
  findChild,
  getAttribute,
  parseBooleanElement,
  parseNumericAttribute,
  findChildren,
  type XmlElement,
} from '../xmlParser';
import { parseColorValue, parseShadingProperties, parseRunProperties } from './runProperties';

/**
 * Parse border specification
 */
export function parseBorderSpec(border: XmlElement | null): BorderSpec | undefined {
  if (!border) return undefined;

  const style = getAttribute(border, 'w', 'val');
  if (!style) return undefined;

  const spec: BorderSpec = {
    style: style as BorderSpec['style'],
  };

  const colorVal = getAttribute(border, 'w', 'color');
  const themeColor = getAttribute(border, 'w', 'themeColor');
  if (colorVal || themeColor) {
    spec.color = parseColorValue(
      colorVal,
      themeColor,
      getAttribute(border, 'w', 'themeTint'),
      getAttribute(border, 'w', 'themeShade')
    );
  }

  const sz = parseNumericAttribute(border, 'w', 'sz');
  if (sz !== undefined) spec.size = sz;

  const space = parseNumericAttribute(border, 'w', 'space');
  if (space !== undefined) spec.space = space;

  const shadowAttr = getAttribute(border, 'w', 'shadow');
  if (shadowAttr) spec.shadow = shadowAttr === '1' || shadowAttr === 'true';

  const frame = getAttribute(border, 'w', 'frame');
  if (frame) spec.frame = frame === '1' || frame === 'true';

  return spec;
}

/**
 * Parse tab stops (w:tabs)
 */
export function parseTabStops(tabs: XmlElement | null): TabStop[] | undefined {
  if (!tabs) return undefined;

  const tabElements = findChildren(tabs, 'w', 'tab');
  if (tabElements.length === 0) return undefined;

  const result: TabStop[] = [];

  for (const tab of tabElements) {
    const pos = parseNumericAttribute(tab, 'w', 'pos');
    const val = getAttribute(tab, 'w', 'val');

    if (pos !== undefined && val) {
      const tabStop: TabStop = {
        position: pos,
        alignment: val as TabStopAlignment,
      };

      const leader = getAttribute(tab, 'w', 'leader');
      if (leader) {
        tabStop.leader = leader as TabLeader;
      }

      result.push(tabStop);
    }
  }

  return result.length > 0 ? result : undefined;
}

/**
 * Parse paragraph formatting properties (w:pPr)
 */
export function parseParagraphProperties(
  pPr: XmlElement | null,
  theme: Theme | null
): ParagraphFormatting | undefined {
  if (!pPr) return undefined;

  const formatting: ParagraphFormatting = {};

  // Alignment
  const jc = findChild(pPr, 'w', 'jc');
  if (jc) {
    const val = getAttribute(jc, 'w', 'val');
    if (val) formatting.alignment = val as ParagraphFormatting['alignment'];
  }

  // Bidi
  const bidi = findChild(pPr, 'w', 'bidi');
  if (bidi) formatting.bidi = parseBooleanElement(bidi);

  // Spacing
  const spacing = findChild(pPr, 'w', 'spacing');
  if (spacing) {
    const before = parseNumericAttribute(spacing, 'w', 'before');
    if (before !== undefined) formatting.spaceBefore = before;

    const after = parseNumericAttribute(spacing, 'w', 'after');
    if (after !== undefined) formatting.spaceAfter = after;

    const line = parseNumericAttribute(spacing, 'w', 'line');
    if (line !== undefined) formatting.lineSpacing = line;

    const lineRule = getAttribute(spacing, 'w', 'lineRule');
    if (lineRule) formatting.lineSpacingRule = lineRule as ParagraphFormatting['lineSpacingRule'];

    const beforeAuto = getAttribute(spacing, 'w', 'beforeAutospacing');
    if (beforeAuto) formatting.beforeAutospacing = beforeAuto === '1' || beforeAuto === 'true';

    const afterAuto = getAttribute(spacing, 'w', 'afterAutospacing');
    if (afterAuto) formatting.afterAutospacing = afterAuto === '1' || afterAuto === 'true';
  }

  // Indentation
  const ind = findChild(pPr, 'w', 'ind');
  if (ind) {
    const left = parseNumericAttribute(ind, 'w', 'left');
    if (left !== undefined) formatting.indentLeft = left;

    const right = parseNumericAttribute(ind, 'w', 'right');
    if (right !== undefined) formatting.indentRight = right;

    const firstLine = parseNumericAttribute(ind, 'w', 'firstLine');
    if (firstLine !== undefined) formatting.indentFirstLine = firstLine;

    const hanging = parseNumericAttribute(ind, 'w', 'hanging');
    if (hanging !== undefined) {
      formatting.indentFirstLine = -hanging;
      formatting.hangingIndent = true;
    }
  }

  // Borders
  const pBdr = findChild(pPr, 'w', 'pBdr');
  if (pBdr) {
    const borders: ParagraphFormatting['borders'] = {};
    const top = parseBorderSpec(findChild(pBdr, 'w', 'top'));
    if (top) borders.top = top;
    const bottom = parseBorderSpec(findChild(pBdr, 'w', 'bottom'));
    if (bottom) borders.bottom = bottom;
    const left = parseBorderSpec(findChild(pBdr, 'w', 'left'));
    if (left) borders.left = left;
    const right = parseBorderSpec(findChild(pBdr, 'w', 'right'));
    if (right) borders.right = right;
    const between = parseBorderSpec(findChild(pBdr, 'w', 'between'));
    if (between) borders.between = between;
    const bar = parseBorderSpec(findChild(pBdr, 'w', 'bar'));
    if (bar) borders.bar = bar;

    if (Object.keys(borders).length > 0) {
      formatting.borders = borders;
    }
  }

  // Shading
  const shd = findChild(pPr, 'w', 'shd');
  if (shd) {
    formatting.shading = parseShadingProperties(shd);
  }

  // Tab stops
  const tabs = findChild(pPr, 'w', 'tabs');
  if (tabs) {
    formatting.tabs = parseTabStops(tabs);
  }

  // Page break control
  const keepNext = findChild(pPr, 'w', 'keepNext');
  if (keepNext) formatting.keepNext = parseBooleanElement(keepNext);

  const keepLines = findChild(pPr, 'w', 'keepLines');
  if (keepLines) formatting.keepLines = parseBooleanElement(keepLines);

  const widowControl = findChild(pPr, 'w', 'widowControl');
  if (widowControl) formatting.widowControl = parseBooleanElement(widowControl);

  const pageBreakBefore = findChild(pPr, 'w', 'pageBreakBefore');
  if (pageBreakBefore) formatting.pageBreakBefore = parseBooleanElement(pageBreakBefore);

  const contextualSpacing = findChild(pPr, 'w', 'contextualSpacing');
  if (contextualSpacing) formatting.contextualSpacing = parseBooleanElement(contextualSpacing);

  // Numbering properties
  const numPr = findChild(pPr, 'w', 'numPr');
  if (numPr) {
    const numId = findChild(numPr, 'w', 'numId');
    const ilvl = findChild(numPr, 'w', 'ilvl');

    if (numId || ilvl) {
      formatting.numPr = {};
      if (numId) {
        const val = parseNumericAttribute(numId, 'w', 'val');
        if (val !== undefined) formatting.numPr.numId = val;
      }
      if (ilvl) {
        const val = parseNumericAttribute(ilvl, 'w', 'val');
        if (val !== undefined) formatting.numPr.ilvl = val;
      }
    }
  }

  // Outline level
  const outlineLvl = findChild(pPr, 'w', 'outlineLvl');
  if (outlineLvl) {
    const val = parseNumericAttribute(outlineLvl, 'w', 'val');
    if (val !== undefined) formatting.outlineLevel = val;
  }

  // Style reference
  const pStyle = findChild(pPr, 'w', 'pStyle');
  if (pStyle) {
    const val = getAttribute(pStyle, 'w', 'val');
    if (val) formatting.styleId = val;
  }

  // Suppress line numbers
  const suppressLineNumbers = findChild(pPr, 'w', 'suppressLineNumbers');
  if (suppressLineNumbers)
    formatting.suppressLineNumbers = parseBooleanElement(suppressLineNumbers);

  // Suppress auto hyphens
  const suppressAutoHyphens = findChild(pPr, 'w', 'suppressAutoHyphens');
  if (suppressAutoHyphens)
    formatting.suppressAutoHyphens = parseBooleanElement(suppressAutoHyphens);

  // Run properties for this paragraph (default run formatting)
  const rPr = findChild(pPr, 'w', 'rPr');
  if (rPr) {
    formatting.runProperties = parseRunProperties(rPr, theme);
  }

  return Object.keys(formatting).length > 0 ? formatting : undefined;
}

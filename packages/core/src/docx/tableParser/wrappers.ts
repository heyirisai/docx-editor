/**
 * Row- and table-level wrapper collection for the table parser: cells inside
 * `w:tr > w:sdt|w:customXml` (`CT_Row`) and rows inside
 * `w:tbl > w:sdt|w:customXml` (`CT_Tbl`).
 */

import type { TableRowTableWrapper } from '../../types/document';
import { parseSdtProperties } from '../sdtProperties';
import { elementToXml, findChild, getLocalName, type XmlElement } from '../xmlParser';

export type CollectedRowWrapper =
  | {
      kind: 'sdt';
      id: number;
      leading: boolean;
      sdtPr: XmlElement | null;
      sdtEndPr: XmlElement | null;
    }
  | { kind: 'customXml'; id: number; startXml: string; endXml: string };

/** Deepest row-level wrapper nesting honoured (guards hostile input). */
const MAX_ROW_WRAPPER_DEPTH = 16;

/**
 * The `w:tc` elements of a row in document order, with the row-level wrappers
 * (`w:sdt` → `w:sdtContent`, `w:customXml`) each one sits inside.
 *
 * Word writes `w:tr > w:sdt > w:sdtContent > w:tc` for a content control that
 * wraps a whole cell (questionnaire answer cells). Collecting only direct
 * `w:tc` children dropped those cells and every paragraph in them, leaving a
 * short row.
 */
export function collectRowCells(
  trElement: XmlElement
): { element: XmlElement; wrappers: CollectedRowWrapper[] }[] {
  const out: { element: XmlElement; wrappers: CollectedRowWrapper[] }[] = [];
  let nextId = 0;

  const walk = (container: XmlElement, wrappers: CollectedRowWrapper[]): void => {
    for (const child of container.elements ?? []) {
      if (child.type !== 'element' || !child.name) continue;
      const localName = getLocalName(child.name);
      if (localName === 'tc') {
        // Snapshot the chain: only the first cell of a control carries it, so
        // the shared wrapper's `leading` flag flips for the cells after it.
        out.push({ element: child, wrappers: wrappers.map((w) => ({ ...w })) });
        for (const w of wrappers) if (w.kind === 'sdt') w.leading = false;
        continue;
      }
      if (wrappers.length >= MAX_ROW_WRAPPER_DEPTH) continue;
      if (localName === 'sdt') {
        const sdtContent = findChild(child, 'w', 'sdtContent');
        if (!sdtContent) continue;
        const wrapper: CollectedRowWrapper = {
          kind: 'sdt',
          id: nextId++,
          leading: true,
          sdtPr: findChild(child, 'w', 'sdtPr'),
          sdtEndPr: findChild(child, 'w', 'sdtEndPr'),
        };
        walk(sdtContent, [...wrappers, wrapper]);
      } else if (localName === 'customXml') {
        const { startXml, endXml } = captureCustomXmlTags(child);
        walk(child, [...wrappers, { kind: 'customXml', id: nextId++, startXml, endXml }]);
      }
    }
  };

  walk(trElement, []);
  return out;
}

/**
 * Table-level wrapper ids are unique across every table parsed, not per
 * table: the row serializer groups rows by id, so a row pasted into another
 * table must not merge with that table's own control of the same index.
 */
let nextTableWrapperId = 0;

/**
 * The `w:tr` elements of a table in document order, with the table-level
 * wrappers (`w:sdt` → `w:sdtContent`, `w:customXml`) each one sits inside,
 * outermost first.
 *
 * `CT_Tbl` row content is `tr | customXml | sdt` (`CT_SdtRow`,
 * `CT_CustomXmlRow`): Word writes `w:tbl > w:sdt > w:sdtContent > w:tr` for a
 * content control around whole rows (a repeating section, or a row-level
 * control). Collecting only direct `w:tr` children dropped those rows and
 * everything in them.
 */
export function collectTableRows(
  tblElement: XmlElement
): { element: XmlElement; wrappers: TableRowTableWrapper[] }[] {
  const out: { element: XmlElement; wrappers: TableRowTableWrapper[] }[] = [];

  const walk = (container: XmlElement, wrappers: TableRowTableWrapper[]): void => {
    for (const child of container.elements ?? []) {
      if (child.type !== 'element' || !child.name) continue;
      const localName = getLocalName(child.name);
      if (localName === 'tr') {
        out.push({ element: child, wrappers: wrappers.map((w) => ({ ...w })) });
        continue;
      }
      if (wrappers.length >= MAX_ROW_WRAPPER_DEPTH) continue;
      if (localName === 'sdt') {
        const sdtContent = findChild(child, 'w', 'sdtContent');
        if (!sdtContent) continue;
        walk(sdtContent, [
          ...wrappers,
          {
            kind: 'sdt',
            id: nextTableWrapperId++,
            properties: parseSdtProperties(
              findChild(child, 'w', 'sdtPr'),
              findChild(child, 'w', 'sdtEndPr')
            ),
          },
        ]);
      } else if (localName === 'customXml') {
        const { startXml, endXml } = captureCustomXmlTags(child);
        walk(child, [
          ...wrappers,
          { kind: 'customXml', id: nextTableWrapperId++, startXml, endXml },
        ]);
      }
    }
  };

  walk(tblElement, []);
  return out;
}

/**
 * Verbatim start tag (with its `w:customXmlPr`) and end tag of a row-level
 * `w:customXml`, so the wrapper is re-emitted unchanged around its cells.
 */
function captureCustomXmlTags(element: XmlElement): { startXml: string; endXml: string } {
  const marker = '\u0000CELLS\u0000';
  const shell: XmlElement = {
    ...element,
    elements: [
      ...(element.elements ?? []).filter(
        (el) => el.type === 'element' && getLocalName(el.name ?? '') === 'customXmlPr'
      ),
      { type: 'text', text: marker },
    ],
  };
  const xml = elementToXml(shell);
  const at = xml.indexOf(marker);
  return { startXml: xml.slice(0, at), endXml: xml.slice(at + marker.length) };
}

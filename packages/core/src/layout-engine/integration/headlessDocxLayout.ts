/**
 * Headless DOCX → layout harness for integration suites.
 *
 * Runs the same compute pass the adapters run (`computeLayout` with the
 * float-aware measure pipeline and the shared `measureTableBlock`), minus the
 * DOM. Callers install `installCanvasDocumentStub()` first so text measures at
 * 0.5em per character and line heights come from the OS/2 ratio table.
 *
 * Kept in a non-`.test.ts` file so bun test doesn't run it as a suite.
 */

import type { EditorState } from 'prosemirror-state';
import type { Document } from '../../types/document';

/** Parse DOCX bytes without the network-bound font preload. */
export async function parseDocxHeadless(bytes: Uint8Array): Promise<Document> {
  const { parseDocx } = await import('../../docx/parser');
  return parseDocx(bytes, { preloadFonts: false });
}

/** The document's PM doc, as the editor builds it on load. */
export async function toProseDocHeadless(doc: Document) {
  const { toProseDoc } = await import('../../prosemirror/conversion/toProseDoc');
  return toProseDoc(doc, { styles: doc.package.styles });
}

/**
 * Lay out `state` (defaults to a plugin-free state built from `doc`) with the
 * first section's page geometry.
 */
export async function layoutDocxHeadless(doc: Document, state?: EditorState) {
  const { EditorState: PMEditorState } = await import('prosemirror-state');
  const { schema } = await import('../../prosemirror/schema');
  const { computeLayout } = await import('../../editor/computeLayout');
  const bridge = await import('../../layout-bridge');
  const { measureParagraph } = await import('../../layout-bridge/measuring/measureParagraph');

  const body = doc.package.document!;
  const sectionProperties = body.sections?.[0]?.properties ?? body.finalSectionProperties ?? null;
  const finalSectionProperties = body.finalSectionProperties ?? sectionProperties;
  const styles = doc.package.styles;
  const theme = doc.package.theme ?? null;

  const editorState = state ?? PMEditorState.create({ doc: await toProseDocHeadless(doc), schema });
  const pageSize = bridge.getPageSize(sectionProperties);
  const margins = bridge.getMargins(sectionProperties);

  // Mirrors the React adapter's measureBlock (minus caching).
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const measureBlock = (block: any, width: number, zones?: any, cumulativeY?: number): any => {
    switch (block.kind) {
      case 'paragraph':
        return measureParagraph(block, width, {
          floatingZones: zones,
          paragraphYOffset: cumulativeY ?? 0,
        });
      case 'table':
        return bridge.measureTableBlock(block, width, measureBlock);
      case 'image':
        return { kind: 'image', width: block.width ?? 100, height: block.height ?? 100 };
      case 'textBox': {
        const m = block.margins ?? { top: 3.6, bottom: 3.6, left: 7.2, right: 7.2 };
        const innerWidth = (block.width ?? 200) - m.left - m.right;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const inner = block.content.map((p: any) => measureParagraph(p, innerWidth));
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const h = inner.reduce((s: number, x: any) => s + x.totalHeight, 0);
        return {
          kind: 'textBox',
          width: block.width ?? 200,
          height: block.height ?? h + m.top + m.bottom,
          innerMeasures: inner,
        };
      }
      default:
        return { kind: block.kind };
    }
  };

  const hf = bridge.resolveHeaderFooter(doc, sectionProperties);
  return computeLayout({
    state: editorState,
    document: doc,
    pageSize,
    margins,
    columns: bridge.getColumns(sectionProperties),
    finalPageSize: bridge.getPageSize(finalSectionProperties),
    finalMargins: bridge.getMargins(finalSectionProperties),
    finalColumns: bridge.getColumns(finalSectionProperties),
    pageGap: 20,
    contentWidth: pageSize.w - margins.left - margins.right,
    theme,
    styles,
    sectionProperties,
    finalSectionProperties,
    headerContent: hf.header,
    footerContent: hf.footer,
    firstPageHeaderContent: hf.firstHeader,
    firstPageFooterContent: hf.firstFooter,
    measureBlocks: (blocks, w, geom) =>
      bridge.measureBlocksWithFloats(blocks, w, measureBlock, geom),
    getHfPmDoc: () => null,
  });
}

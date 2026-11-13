import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, beforeAll } from 'bun:test';

beforeAll(() => GlobalRegistrator.register());
afterAll(() => GlobalRegistrator.unregister());
import { afterEach, describe, expect, test } from 'bun:test';
import { act, cleanup, renderHook } from '@testing-library/react';
import { EditorState } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { schema } from '@eigenpal/docx-editor-core/prosemirror';
import type { Document, SectionProperties } from '@eigenpal/docx-editor-core/types/document';
import { useHeaderFooterEditing } from './useHeaderFooterEditing';

afterEach(() => {
  cleanup();
});

/**
 * Two sections, no header parts anywhere: section 0 ends at the first
 * paragraph (its sectPr is a PM paragraph attr), section 1 is the body sectPr.
 */
function setup() {
  const firstSectPr: SectionProperties = { sectionStart: 'nextPage' };
  const finalSectPr: SectionProperties = { pageWidth: 12240, pageHeight: 15840 };
  let state = EditorState.create({
    doc: schema.node('doc', null, [
      schema.node('paragraph', { _sectionProperties: firstSectPr }),
      schema.node('paragraph'),
    ]),
  });
  const view = {
    get state() {
      return state;
    },
    dispatch(tr: Parameters<EditorState['apply']>[0]) {
      state = state.apply(tr);
    },
  } as unknown as EditorView;
  const document = {
    package: {
      document: {
        content: [
          { type: 'paragraph', content: [], sectionProperties: firstSectPr },
          { type: 'paragraph', content: [] },
        ],
        finalSectionProperties: finalSectPr,
      },
      headers: new Map(),
      footers: new Map(),
      relationships: new Map(),
    },
  } as unknown as Document;
  const pushed: Document[] = [];
  const { result } = renderHook(() =>
    useHeaderFooterEditing({
      document,
      pushDocument: (d) => void pushed.push(d),
      hfEditorRef: { current: null },
      containerRef: { current: null },
      initialSectionProperties: firstSectPr,
      finalSectionProperties: finalSectPr,
      hfEditPosition: null,
      setHfEditPosition: () => {},
      hfEditIsFirstPage: false,
      setHfEditIsFirstPage: () => {},
      mainEditorRef: { current: { getView: () => view } },
    })
  );
  return { result, pushed, view: () => view };
}

describe('useHeaderFooterEditing: double-click on a page showing no header', () => {
  test("creates the part in the clicked page's own section", () => {
    const { result, pushed, view } = setup();
    act(() =>
      result.current.handleHeaderFooterDoubleClick('header', 1, null, {
        index: 0,
        firstPageOfSection: true,
      })
    );
    const sectPr = view().state.doc.child(0).attrs._sectionProperties as SectionProperties;
    const rId = sectPr.headerReferences?.[0]?.rId;
    expect(sectPr.headerReferences).toEqual([{ type: 'default', rId: rId! }]);
    const doc = pushed.at(-1)!;
    expect(doc.package.headers?.has(rId!)).toBe(true);
    expect(doc.package.document.finalSectionProperties?.headerReferences).toBeUndefined();
  });

  test('creates the part in the body sectPr for the last section', () => {
    const { result, pushed, view } = setup();
    act(() =>
      result.current.handleHeaderFooterDoubleClick('footer', 2, null, {
        index: 1,
        firstPageOfSection: true,
      })
    );
    const refs = pushed.at(-1)!.package.document.finalSectionProperties?.footerReferences;
    expect(refs).toHaveLength(1);
    expect(pushed.at(-1)!.package.footers?.has(refs![0]!.rId)).toBe(true);
    const sectPr = view().state.doc.child(0).attrs._sectionProperties as SectionProperties;
    expect(sectPr.footerReferences).toBeUndefined();
  });
});

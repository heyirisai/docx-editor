import { describe, test, expect } from 'bun:test';
import { ref, shallowRef } from 'vue';
import { EditorState } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { schema } from '@eigenpal/docx-editor-core/prosemirror';
import type { Document, SectionProperties } from '@eigenpal/docx-editor-core/types/document';
import { usePagesPointer, type UsePagesPointerOptions } from './usePagesPointer';

/**
 * Double-clicking the header of a page that shows none creates the part in
 * that page's OWN section (as Word does), not in the body-level sectPr.
 * Two sections: section 0's sectPr is a PM paragraph attr, section 1 is final.
 */
function usePointerFixture() {
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
  const doc = {
    package: {
      document: { content: [], finalSectionProperties: finalSectPr },
      headers: new Map(),
      footers: new Map(),
      relationships: new Map(),
    },
  } as unknown as Document;
  const set: Document[] = [];
  const api = usePagesPointer({
    editorView: shallowRef(view),
    pagesRef: ref(null),
    pagesViewportRef: ref(null),
    selectedImage: shallowRef(null),
    imageInteracting: ref(false),
    hyperlinkPopupData: ref(null),
    readOnly: ref(false),
    zoom: ref(1),
    layout: ref({
      pages: [
        { number: 1, sectionIndex: 0, firstPageOfSection: true },
        { number: 2, sectionIndex: 1, firstPageOfSection: true },
      ],
    }),
    getCommands: () => ({}),
    getDocument: () => doc,
    reLayout: () => {},
    emit: () => {},
    clearOverlay: () => {},
    setDocument: (d: Document) => void set.push(d),
  } as unknown as UsePagesPointerOptions);

  const dblclickHeader = (pageNumber: number) => {
    const pageEl = { dataset: { pageNumber: String(pageNumber) } };
    const headerEl = {
      closest: (sel: string) =>
        sel === '.layout-page-header' ? headerEl : sel === '.layout-page' ? pageEl : null,
    };
    api.handlePagesDoubleClick({ target: headerEl } as unknown as MouseEvent);
  };
  const firstSectionRefs = () =>
    (state.doc.child(0).attrs._sectionProperties as SectionProperties).headerReferences;
  return { dblclickHeader, set, firstSectionRefs };
}

describe('Vue header double-click on a page showing no header', () => {
  test("creates the part in the clicked page's own section", () => {
    const { dblclickHeader, set, firstSectionRefs } = usePointerFixture();
    dblclickHeader(1);
    const refs = firstSectionRefs();
    expect(refs).toHaveLength(1);
    expect(set.at(-1)!.package.headers?.has(refs![0]!.rId)).toBe(true);
    expect(set.at(-1)!.package.document.finalSectionProperties?.headerReferences).toBeUndefined();
  });

  test('creates the part in the body sectPr for the last section', () => {
    const { dblclickHeader, set, firstSectionRefs } = usePointerFixture();
    dblclickHeader(2);
    expect(set.at(-1)!.package.document.finalSectionProperties?.headerReferences).toHaveLength(1);
    expect(firstSectionRefs()).toBeUndefined();
  });
});

import { useCallback, useMemo, useState } from 'react';
import type {
  Document,
  HeaderFooter,
  BlockContent,
  SectionProperties,
} from '@eigenpal/docx-editor-core/types/document';
import {
  resolveHeaderFooter,
  type HeaderFooterDoubleClickHandler,
} from '@eigenpal/docx-editor-core/layout-bridge';
import { fromProseDoc, proseDocToBlocks } from '@eigenpal/docx-editor-core/prosemirror/conversion';
import type { EditorView } from 'prosemirror-view';
import type { InlineHeaderFooterEditorRef } from '../../InlineHeaderFooterEditor';

/**
 * Owns the inline header/footer editing mode: which slot is being
 * edited (`hfEditPosition`), whether the first-page variant applies
 * (`hfEditIsFirstPage`), the resolved header/footer content for the
 * current section, plus the double-click → edit, save, remove, and
 * "click out" workflows.
 *
 * Empty headers/footers are materialised on first double-click so the
 * user can start typing — the helper writes the new HeaderFooter into
 * `package.headers` / `package.footers` and registers the relationship
 * so the serializer picks it up (#274).
 */
export function useHeaderFooterEditing({
  document,
  pushDocument,
  hfEditorRef,
  containerRef,
  initialSectionProperties,
  finalSectionProperties,
  hfEditPosition,
  setHfEditPosition,
  hfEditIsFirstPage,
  setHfEditIsFirstPage,
  mainEditorRef,
}: {
  document: Document | null;
  pushDocument: (doc: Document) => void;
  hfEditorRef: React.RefObject<InlineHeaderFooterEditorRef | null>;
  containerRef: React.RefObject<HTMLDivElement | null>;
  initialSectionProperties: SectionProperties | undefined;
  finalSectionProperties: SectionProperties | undefined;
  // State + setters live in the parent so `getActiveEditorView` (declared
  // before this hook is called) can read `hfEditPosition` for routing.
  hfEditPosition: 'header' | 'footer' | null;
  setHfEditPosition: React.Dispatch<React.SetStateAction<'header' | 'footer' | null>>;
  hfEditIsFirstPage: boolean;
  setHfEditIsFirstPage: React.Dispatch<React.SetStateAction<boolean>>;
  /** Main body view: a non-final section's `w:sectPr` is a paragraph attr there. */
  mainEditorRef?: React.RefObject<{ getView(): EditorView | null } | null>;
}) {
  // The part a double-click opened when the clicked page shows its own
  // section's header/footer (sections paint their own parts, ECMA-376 §17.10).
  const [editTarget, setEditTarget] = useState<{
    position: 'header' | 'footer';
    rId: string;
    isFirstPage: boolean;
  } | null>(null);

  const { headerContent, footerContent, firstPageHeaderContent, firstPageFooterContent } =
    useMemo(() => {
      const { header, footer, firstHeader, firstFooter } = resolveHeaderFooter(
        document ?? null,
        finalSectionProperties ?? initialSectionProperties
      );
      const resolved = {
        headerContent: header,
        footerContent: footer,
        firstPageHeaderContent: firstHeader,
        firstPageFooterContent: firstFooter,
      };
      if (editTarget && hfEditPosition === editTarget.position) {
        const map =
          editTarget.position === 'header'
            ? document?.package?.headers
            : document?.package?.footers;
        const hf = map?.get(editTarget.rId);
        if (hf) {
          const key =
            editTarget.position === 'header'
              ? editTarget.isFirstPage
                ? 'firstPageHeaderContent'
                : 'headerContent'
              : editTarget.isFirstPage
                ? 'firstPageFooterContent'
                : 'footerContent';
          resolved[key] = hf;
        }
      }
      return resolved;
    }, [document, initialSectionProperties, finalSectionProperties, editTarget, hfEditPosition]);

  const handleHeaderFooterDoubleClick = useCallback<HeaderFooterDoubleClickHandler>(
    (position, pageNumber, shown, section) => {
      // No scroll-to-page-1 — the HF content is shared across all pages by
      // `r:id`, so the painter renders the same edits on every page in real
      // time. Whichever page the user double-clicked, the chrome bar floats
      // over THAT page's header and edits propagate visually to all others.
      // Edit the part the clicked page shows, when it shows one.
      const shownMap =
        position === 'header' ? document?.package?.headers : document?.package?.footers;
      if (shown && shownMap?.has(shown.rId)) {
        setEditTarget({ position, rId: shown.rId, isFirstPage: shown.isFirstPage });
        setHfEditIsFirstPage(shown.isFirstPage);
        setHfEditPosition(position);
        return;
      }
      setEditTarget(null);
      if (!document?.package) return;
      const pkg = document.package;

      // The clicked page's own section (ECMA-376 §17.10): its `w:sectPr` lives
      // on a section-break paragraph of the main PM doc, or is the body-level
      // one for the last section. Without page info, the last section.
      const view = mainEditorRef?.current?.getView() ?? null;
      const sections: { sp: SectionProperties | undefined; pos: number | null }[] = [];
      view?.state.doc.descendants((node, pos) => {
        if (node.type.name === 'table') return false;
        const sp = node.attrs._sectionProperties as SectionProperties | null | undefined;
        if (node.type.name === 'paragraph' && sp) sections.push({ sp, pos });
        return true;
      });
      sections.push({ sp: pkg.document?.finalSectionProperties, pos: null });
      const index = Math.min(section?.index ?? sections.length - 1, sections.length - 1);

      // References and titlePg inherit from earlier sections.
      const refKey = position === 'header' ? 'headerReferences' : 'footerReferences';
      let titlePg = false;
      for (let i = 0; i <= index; i++) titlePg = sections[i]!.sp?.titlePg ?? titlePg;
      const isFirstPage =
        titlePg && (section ? section.firstPageOfSection : (pageNumber ?? 1) === 1);
      const hdrFtrType: 'first' | 'default' = isFirstPage ? 'first' : 'default';
      let inherited: string | undefined;
      for (let i = 0; i <= index; i++) {
        const ref = sections[i]!.sp?.[refKey]?.find((r) => r.type === hdrFtrType);
        if (ref?.rId) inherited = ref.rId;
      }
      setHfEditIsFirstPage(isFirstPage);
      const mapKey = position === 'header' ? 'headers' : 'footers';
      if (inherited && pkg[mapKey]?.has(inherited)) {
        setEditTarget({ position, rId: inherited, isFirstPage });
        setHfEditPosition(position);
        return;
      }

      // Materialise an empty header/footer in THAT section's sectPr, as Word
      // does, so the user can start typing. Later sections that do not name
      // their own part inherit it.
      const target = sections[index]!;
      if (!target.sp || (target.pos !== null && !view)) return;

      const existingRels = pkg.relationships;
      let rId = `rId_new_${position}_${hdrFtrType}`;
      for (let n = 2; existingRels?.has(rId) || pkg[mapKey]?.has(rId); n++) {
        rId = `rId_new_${position}_${hdrFtrType}_${n}`;
      }
      const emptyHf: HeaderFooter = {
        type: position === 'header' ? 'header' : 'footer',
        hdrFtrType,
        content: [{ type: 'paragraph', content: [] }],
      };
      const newMap = new Map(pkg[mapKey] ?? []);
      newMap.set(rId, emptyHf);

      // Register the rel so the serializer wires up content types + doc rels (#274).
      const usedTargets = new Set<string>();
      for (const rel of existingRels?.values() ?? []) {
        if (rel.target) usedTargets.add(rel.target);
      }
      let targetNum = 1;
      while (usedTargets.has(`${position}${targetNum}.xml`)) targetNum++;
      const relType =
        position === 'header'
          ? 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header'
          : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer';
      const newRelationships = new Map(existingRels);
      newRelationships.set(rId, {
        id: rId,
        type: relType,
        target: `${position}${targetNum}.xml`,
      });

      const sectionProps: SectionProperties = {
        ...target.sp,
        [refKey]: [...(target.sp[refKey] ?? []), { type: hdrFtrType, rId }],
      };
      let base: Document = document;
      if (target.pos !== null && view) {
        // A non-final section: its sectPr is a PM paragraph attribute.
        const node = view.state.doc.nodeAt(target.pos)!;
        view.dispatch(
          view.state.tr.setNodeMarkup(target.pos, undefined, {
            ...node.attrs,
            _sectionProperties: sectionProps,
          })
        );
        base = fromProseDoc(view.state.doc, document);
      }
      const basePkg = base.package;
      pushDocument({
        ...base,
        package: {
          ...basePkg,
          [mapKey]: newMap,
          relationships: newRelationships,
          document:
            target.pos === null && basePkg.document
              ? { ...basePkg.document, finalSectionProperties: sectionProps }
              : basePkg.document,
        },
      });
      setEditTarget({ position, rId, isFirstPage });
      setHfEditPosition(position);
    },
    [document, mainEditorRef, pushDocument, setHfEditPosition, setHfEditIsFirstPage]
  );

  const handleHeaderFooterSave = useCallback(
    (content: BlockContent[]) => {
      if (!hfEditPosition || !document?.package) {
        setHfEditPosition(null);
        return;
      }

      const pkg = document.package;
      const sectionProps = pkg.document?.finalSectionProperties;
      const refs =
        hfEditPosition === 'header'
          ? sectionProps?.headerReferences
          : sectionProps?.footerReferences;
      const targetType = hfEditIsFirstPage ? 'first' : 'default';
      const target =
        editTarget && editTarget.position === hfEditPosition
          ? { type: editTarget.isFirstPage ? 'first' : 'default', rId: editTarget.rId }
          : null;
      const activeRef =
        target ??
        refs?.find((r) => r.type === targetType) ??
        refs?.find((r) => r.type === 'default') ??
        refs?.find((r) => r.type === 'first') ??
        refs?.[0];
      const mapKey = hfEditPosition === 'header' ? 'headers' : 'footers';
      const map = pkg[mapKey];

      if (activeRef?.rId && map) {
        const existing = map.get(activeRef.rId);
        const updated: HeaderFooter = {
          type: hfEditPosition,
          hdrFtrType: activeRef.type as 'default' | 'first' | 'even',
          ...existing,
          content,
        };
        const newMap = new Map(map);
        newMap.set(activeRef.rId, updated);

        const newDoc: Document = {
          ...document,
          package: {
            ...pkg,
            [mapKey]: newMap,
          },
        };
        pushDocument(newDoc);
      }

      setHfEditPosition(null);
    },
    [hfEditPosition, hfEditIsFirstPage, editTarget, document, pushDocument, setHfEditPosition]
  );

  const handleBodyClick = useCallback(() => {
    if (!hfEditPosition) return;
    // Save current HF contents (if dirty) then close.
    const view = hfEditorRef.current?.getView();
    if (view) {
      const blocks = proseDocToBlocks(view.state.doc);
      handleHeaderFooterSave(blocks);
    } else {
      setHfEditPosition(null);
    }
  }, [hfEditPosition, handleHeaderFooterSave, hfEditorRef, setHfEditPosition]);

  const handleRemoveHeaderFooter = useCallback(() => {
    if (!hfEditPosition || !document?.package) {
      setHfEditPosition(null);
      return;
    }

    const pkg = document.package;
    const sectionProps = pkg.document?.finalSectionProperties;
    const refKey = hfEditPosition === 'header' ? 'headerReferences' : 'footerReferences';
    const mapKey = hfEditPosition === 'header' ? 'headers' : 'footers';
    const refs = sectionProps?.[refKey];
    const delTargetType = hfEditIsFirstPage ? 'first' : 'default';
    if (
      editTarget &&
      editTarget.position === hfEditPosition &&
      !refs?.some((r) => r.rId === editTarget.rId)
    ) {
      // The part belongs to an earlier section only; removing it would need
      // that section's sectPr rewritten — leave it in place.
      setHfEditPosition(null);
      return;
    }
    const activeRef =
      (editTarget && editTarget.position === hfEditPosition
        ? refs?.find((r) => r.rId === editTarget.rId)
        : undefined) ??
      refs?.find((r) => r.type === delTargetType) ??
      refs?.find((r) => r.type === 'default') ??
      refs?.find((r) => r.type === 'first') ??
      refs?.[0];

    if (activeRef?.rId) {
      const newMap = new Map(pkg[mapKey] ?? []);
      newMap.delete(activeRef.rId);

      const newRefs = (refs ?? []).filter((r) => r.rId !== activeRef.rId);

      const newDoc: Document = {
        ...document,
        package: {
          ...pkg,
          [mapKey]: newMap,
          document: pkg.document
            ? {
                ...pkg.document,
                finalSectionProperties: {
                  ...sectionProps,
                  [refKey]: newRefs,
                },
              }
            : pkg.document,
        },
      };
      pushDocument(newDoc);
    }

    setHfEditPosition(null);
  }, [hfEditPosition, hfEditIsFirstPage, editTarget, document, pushDocument, setHfEditPosition]);

  const getHfTargetElement = useCallback(
    (pos: 'header' | 'footer'): HTMLElement | null => {
      const pagesContainer = containerRef.current?.querySelector('.paged-editor__pages');
      if (!pagesContainer) return null;
      const className = pos === 'header' ? '.layout-page-header' : '.layout-page-footer';
      return pagesContainer.querySelector(className);
    },
    [containerRef]
  );

  return {
    headerContent,
    footerContent,
    firstPageHeaderContent,
    firstPageFooterContent,
    handleHeaderFooterDoubleClick,
    handleHeaderFooterSave,
    handleBodyClick,
    handleRemoveHeaderFooter,
    getHfTargetElement,
  };
}

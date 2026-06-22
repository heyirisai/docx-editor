/**
 * Selection-overlay composable — owns the text-caret blink + selection-
 * rect painter (`updateSelectionOverlay`), the cleanup
 * (`clearOverlay`), and the lifecycle for the caret blink interval.
 *
 * The parent still owns the `onSelectionUpdate` callback that the
 * editor view dispatches into, because `useDocxEditor` consumes it at
 * construction time — but the parent's body delegates the overlay
 * repaint to `updateSelectionOverlay` from this composable.
 *
 * Writes back into `selectedImage` (from `useImageActions`) when the
 * PM doc holds a NodeSelection on an image — the overlay rerolls the
 * image's bounding box after layout repaints so resize / move / rotate
 * gestures keep their handles anchored.
 */

import { onBeforeUnmount, type Ref, type ShallowRef } from 'vue';
import type { EditorView } from 'prosemirror-view';
import { NodeSelection } from 'prosemirror-state';
import {
  getSelectionRectsFromDom,
  getCaretPositionFromDom,
} from '@heyirisai/docx-editor-core/layout-bridge/clickToPositionDom';
import {
  findBodyPmAnchor,
  applyCellSelectionHighlight,
} from '@heyirisai/docx-editor-core/layout-bridge';
import { findImageElement } from '@heyirisai/docx-editor-core/layout-painter';
import type { ImageSelectionInfo } from '../components/imageSelectionTypes';
import { Z_INDEX } from '../styles/zIndex';

export interface UseSelectionSyncOptions {
  editorView: Ref<EditorView | null>;
  pagesRef: Ref<HTMLElement | null>;
  /**
   * Current zoom factor. The caret + selection rects are painted into the
   * `.docx-editor-vue__pages` container, which carries `transform: scale(zoom)`.
   * `getBoundingClientRect` returns post-transform (viewport) coordinates, so
   * the rects must be divided by zoom to land in the container's own
   * coordinate space — otherwise the parent's scale multiplies them a second
   * time and the highlight drifts off the text at any zoom ≠ 100%.
   */
  zoom: Ref<number>;
  selectedImage: ShallowRef<ImageSelectionInfo | null>;
  /**
   * True while the user is editing a header/footer. When set, the body PM's
   * caret + selection rects MUST stay hidden so the user doesn't see two
   * carets blinking simultaneously (one in the painted header, one in the
   * body). The HF view's own caret rect is drawn by DocxEditor.vue.
   */
  isHfEditing?: Ref<boolean>;
  /**
   * True while a resize / move / rotate gesture is in flight. Suppresses the
   * post-transaction "clear the overlay" path so the handles don't vanish
   * mid-drag when an intermediate state momentarily isn't an image
   * NodeSelection. Mirrors React's `isImageInteractingRef`.
   */
  imageInteracting?: Ref<boolean>;
}

export interface UseSelectionSyncReturn {
  clearOverlay: () => void;
  updateSelectionOverlay: () => void;
}

export function useSelectionSync(opts: UseSelectionSyncOptions): UseSelectionSyncReturn {
  let caretBlinkInterval: ReturnType<typeof setInterval> | null = null;
  let caretEl: HTMLElement | null = null;
  let imageSyncRaf: number | null = null;

  function clearOverlay() {
    const container = opts.pagesRef.value;
    if (!container) return;
    container.querySelectorAll('.vue-sel-rect, .vue-caret').forEach((el) => el.remove());
    if (caretBlinkInterval !== null) {
      clearInterval(caretBlinkInterval);
      caretBlinkInterval = null;
    }
    caretEl = null;
  }

  /**
   * Re-derive `selectedImage` from the LIVE body PM selection, deferred a frame
   * so the painter has repainted and PM positions resolve against the fresh
   * DOM. Mirrors React's `handleSelectionChange`.
   *
   * An inline image pushed onto another page (e.g. by pressing Enter above it)
   * keeps its `NodeSelection`, which PM maps forward — so resolving the element
   * by the *current* `sel.from` makes the overlay follow the image instead of
   * latching onto whatever now sits at the stale position. Resolution goes
   * through the body-scoped `findBodyPmAnchor` so a header/footer run (a
   * separate PM doc whose positions overlap the body's) can never match.
   *
   * When the doc no longer holds an image `NodeSelection`, the overlay is
   * dropped — unless a resize / move / rotate gesture is mid-flight, whose
   * intermediate transactions would otherwise flicker it away.
   */
  function syncSelectedImageToSelection() {
    if (imageSyncRaf !== null) cancelAnimationFrame(imageSyncRaf);
    imageSyncRaf = requestAnimationFrame(() => {
      imageSyncRaf = null;
      const container = opts.pagesRef.value;
      const view = opts.editorView.value;
      if (!container || !view) return;
      // HF editing drives its own selection model; leave the body image alone.
      if (opts.isHfEditing?.value) return;

      const sel = view.state.selection;
      if (sel instanceof NodeSelection && sel.node.type.name === 'image') {
        const anchor = findBodyPmAnchor(container, sel.from);
        const imgEl = anchor ? findImageElement(anchor) : null;
        if (imgEl) {
          const prev = opts.selectedImage.value;
          if (
            !prev ||
            prev.element !== imgEl ||
            prev.pmPos !== sel.from ||
            prev.width !== imgEl.offsetWidth ||
            prev.height !== imgEl.offsetHeight
          ) {
            opts.selectedImage.value = {
              element: imgEl,
              pmPos: sel.from,
              width: imgEl.offsetWidth,
              height: imgEl.offsetHeight,
            };
          }
          return;
        }
      }
      // Not an image NodeSelection (or it resolved off-screen): drop the
      // overlay so it can't strand on a stale spot — but never mid-gesture.
      if (!opts.imageInteracting?.value) {
        opts.selectedImage.value = null;
      }
    });
  }

  function updateSelectionOverlay() {
    const container = opts.pagesRef.value;
    const view = opts.editorView.value;
    if (!container || !view) return;

    clearOverlay();

    // In HF edit mode the body PM has no business showing a caret or
    // selection — the user is editing the header/footer above.
    if (opts.isHfEditing?.value) return;

    // Keep the image overlay glued to the live selection after every change.
    syncSelectedImageToSelection();

    // Paint the multi-cell selection highlight on the body table cells. Runs
    // before the image/text branches so it both lights up an active
    // CellSelection and clears a stale highlight when the selection moves to
    // text or an image. Mirrors React's `applyCellSelectionHighlight` call.
    applyCellSelectionHighlight(container, view.state);

    // An image NodeSelection is painted by ImageSelectionOverlay, not here —
    // suppress the text caret / selection rects so they don't double up. Gate
    // on the live selection (not `selectedImage`, which lags a frame behind the
    // deferred sync above) so the caret reappears the instant focus leaves the
    // image.
    const sel = view.state.selection;
    if (sel instanceof NodeSelection && sel.node.type.name === 'image') return;

    const { from, to, empty } = sel;

    // Account for scroll offset: overlays are position:absolute inside the
    // scrollable container, so we need to add scrollTop/scrollLeft to convert
    // viewport-relative coordinates from getBoundingClientRect to container-relative.
    const scrollTop = container.scrollTop;
    const scrollLeft = container.scrollLeft;

    // The container is scaled via `transform: scale(zoom)`, but the geometry
    // from `getBoundingClientRect`/`getClientRects` is post-transform px.
    // Divide those by zoom so the overlay divs — children of the scaled
    // container — render at the right spot once the parent's scale is applied.
    // The caret *height* is returned in layout px by getCaretPositionFromDom
    // (it normalizes its one scaled measurement by `zoom`; #928), so it is
    // used as-is and the parent's scale grows it to match the line.
    const zoom = opts.zoom.value || 1;

    if (empty) {
      // Draw blinking caret
      const overlayRect = container.getBoundingClientRect();
      const caret = getCaretPositionFromDom(container, from, overlayRect, zoom);
      if (caret) {
        const el = document.createElement('div');
        el.className = 'vue-caret';
        el.style.cssText = `
          position: absolute;
          left: ${caret.x / zoom + scrollLeft}px;
          top: ${caret.y / zoom + scrollTop}px;
          width: 2px;
          height: ${caret.height}px;
          background: var(--doc-caret, #000);
          pointer-events: none;
          z-index: ${Z_INDEX.selectionOverlay};
        `;
        container.appendChild(el);
        caretEl = el;

        // Blink
        let visible = true;
        caretBlinkInterval = setInterval(() => {
          visible = !visible;
          if (caretEl) caretEl.style.opacity = visible ? '1' : '0';
        }, 530);
      }
      return;
    }

    // Draw selection highlight rects (character-level)
    const overlayRect = container.getBoundingClientRect();
    const rects = getSelectionRectsFromDom(container, from, to, overlayRect);

    for (const rect of rects) {
      const el = document.createElement('div');
      el.className = 'vue-sel-rect';
      el.style.cssText = `
        position: absolute;
        left: ${rect.x / zoom + scrollLeft}px;
        top: ${rect.y / zoom + scrollTop}px;
        width: ${rect.width / zoom}px;
        height: ${rect.height / zoom}px;
        background: rgba(66, 133, 244, 0.3);
        pointer-events: none;
        z-index: ${Z_INDEX.selectionOverlay};
      `;
      container.appendChild(el);
    }
  }

  onBeforeUnmount(() => {
    if (imageSyncRaf !== null) cancelAnimationFrame(imageSyncRaf);
    clearOverlay();
  });

  return {
    clearOverlay,
    updateSelectionOverlay,
  };
}

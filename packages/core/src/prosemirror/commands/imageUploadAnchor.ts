import { Plugin, PluginKey } from 'prosemirror-state';
import { Decoration, DecorationSet, type EditorView } from 'prosemirror-view';

interface ImageUploadAnchorAction {
  add?: { id: object; pos: number };
  remove?: { id: object };
}

/** @internal */
export interface ImageUploadAnchor {
  id: object;
  fallbackPos: number;
  tracked: boolean;
}

const imageUploadAnchorKey = new PluginKey<DecorationSet>('imageUploadAnchor');

/**
 * Tracks async image insertion positions through intervening transactions.
 *
 * @internal
 */
export const imageUploadAnchorPlugin = new Plugin<DecorationSet>({
  key: imageUploadAnchorKey,
  state: {
    init: () => DecorationSet.empty,
    apply(transaction, anchors) {
      let next = anchors.map(transaction.mapping, transaction.doc);
      const action = transaction.getMeta(imageUploadAnchorKey) as
        | ImageUploadAnchorAction
        | undefined;

      if (action?.add) {
        next = next.add(transaction.doc, [
          Decoration.widget(action.add.pos, () => document.createTextNode(''), {
            id: action.add.id,
            // Keep a batch anchor after each image inserted at its position.
            side: 1,
          }),
        ]);
      }
      if (action?.remove) {
        next = next.remove(
          next.find(undefined, undefined, (spec) => spec.id === action.remove?.id)
        );
      }
      return next;
    },
  },
});

/** @internal */
export function createImageUploadAnchor(view: EditorView): ImageUploadAnchor {
  const anchor: ImageUploadAnchor = {
    id: {},
    fallbackPos: view.state.selection.from,
    tracked: imageUploadAnchorKey.getState(view.state) !== undefined,
  };
  if (anchor.tracked) {
    view.dispatch(
      view.state.tr
        .setMeta(imageUploadAnchorKey, {
          add: { id: anchor.id, pos: anchor.fallbackPos },
        } satisfies ImageUploadAnchorAction)
        .setMeta('addToHistory', false)
    );
  }
  return anchor;
}

/** @internal */
export function resolveImageUploadAnchor(
  view: EditorView,
  anchor: ImageUploadAnchor
): number | null {
  if (!anchor.tracked) {
    return Math.min(anchor.fallbackPos, view.state.doc.content.size);
  }
  return (
    imageUploadAnchorKey
      .getState(view.state)
      ?.find(undefined, undefined, (spec) => spec.id === anchor.id)[0]?.from ?? null
  );
}

/** @internal */
export function removeImageUploadAnchor(view: EditorView, anchor: ImageUploadAnchor): void {
  if (!anchor.tracked || view.isDestroyed) return;
  view.dispatch(
    view.state.tr
      .setMeta(imageUploadAnchorKey, {
        remove: { id: anchor.id },
      } satisfies ImageUploadAnchorAction)
      .setMeta('addToHistory', false)
  );
}

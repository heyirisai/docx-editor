/**
 * Image commands — thin re-exports from the extension system.
 *
 * Wrap-type transitions for floating images. Inline↔anchor conversions are
 * structural and live in a follow-up; this surface only covers anchor↔anchor.
 */

import type { Command, EditorState, Transaction } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import type { EditorView } from 'prosemirror-view';
import { singletonManager } from '../schema';
import { makeRevisionInfo } from '../plugins/revisionIds';
import {
  createImageUploadAnchor,
  removeImageUploadAnchor,
  resolveImageUploadAnchor,
  type ImageUploadAnchor,
} from './imageUploadAnchor';
import type {
  ImageLayoutTarget,
  SetImageWrapTypeOptions,
} from '../extensions/nodes/ImageExtension';

/**
 * Insert an image node at `pos`, wrapping with the `insertion` mark when
 * suggesting mode is active. Centralizes the tracked-image-insert flow
 * so React `useFileIO`, Vue `useImageActions`, and clipboard-paste
 * paths all share one source of truth — adding a fresh image in
 * suggesting mode always round-trips as `<w:ins>{run with drawing}</w:ins>`.
 *
 * Caller responsibility: produce the `image` node via `schema.nodes.image.create`.
 * This helper handles the dispatch + optional mark application.
 *
 * @public
 */
export function insertImageNode(
  state: EditorState,
  dispatch: ((tr: Transaction) => void) | undefined,
  imageNode: PMNode,
  pos: number
): boolean {
  if (!dispatch) return true;
  dispatch(buildImageInsertionTransaction(state, imageNode, pos).scrollIntoView());
  return true;
}

/**
 * Default max width (px) for an image inserted from a file picker — the content
 * area of a US Letter page at 96dpi (~6.375in). Images wider than this are
 * scaled down to fit the column, matching Word and keeping the painter's
 * `max-width: 100%` from shrinking the rendered height out from under the
 * reserved line height (which would leave a gap below the image).
 */
export const INSERT_IMAGE_MAX_WIDTH_PX = 612;

/**
 * Result returned by a host that persists an image outside the collaborative
 * document. The identifier must be opaque: it is safe to share in document
 * state, but it is not itself a URL or bearer credential.
 *
 * @public
 */
export interface ImageUploadResult {
  assetId: string;
}

/**
 * Metadata supplied to a host image uploader after the browser has decoded
 * the local file. Dimensions are the original pixel dimensions, before the
 * editor scales the rendered node to the page width.
 *
 * @public
 */
export interface ImageUploadContext {
  width: number;
  height: number;
}

/**
 * Optional host hook for persisting image bytes before an image node is
 * inserted. When present, the editor inserts only the returned opaque asset
 * identity and never places a data URL in ProseMirror.
 *
 * @public
 */
export type ImageUploadHandler = (
  file: File,
  context: ImageUploadContext
) => Promise<ImageUploadResult>;

/**
 * Options for {@link insertImageFromFile}.
 *
 * @public
 */
export interface InsertImageFromFileOptions {
  maxWidth?: number;
  imageUploadHandler?: ImageUploadHandler;
  onError?: (error: unknown) => void;
  onInserted?: () => void;
}

function readFileAsDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as string);
    reader.onerror = () => reject(reader.error ?? new Error('Failed to read image'));
    reader.readAsDataURL(file);
  });
}

function loadImageSize(src: string): Promise<{ width: number; height: number }> {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () =>
      resolve({
        width: image.naturalWidth || 1,
        height: image.naturalHeight || 1,
      });
    image.onerror = () => reject(new Error('Failed to decode image'));
    image.src = src;
  });
}

function buildImageInsertionTransaction(
  state: EditorState,
  imageNode: PMNode,
  pos: number
): Transaction {
  const transaction = state.tr.insert(pos, imageNode);
  const info = makeRevisionInfo(state);
  const insertionType = state.schema.marks.insertion;
  if (info && insertionType) {
    transaction.addMark(
      pos,
      pos + imageNode.nodeSize,
      insertionType.create({
        revisionId: info.revisionId,
        author: info.author,
        date: info.date,
      })
    );
  }
  return transaction;
}

async function insertImageFromFileAtAnchor(
  view: EditorView,
  file: File,
  opts: InsertImageFromFileOptions | undefined,
  anchor: ImageUploadAnchor
): Promise<boolean> {
  const maxWidth = opts?.maxWidth ?? INSERT_IMAGE_MAX_WIDTH_PX;
  let objectUrl: string | null = null;

  try {
    const imageSource = opts?.imageUploadHandler
      ? (objectUrl = URL.createObjectURL(file))
      : await readFileAsDataUrl(file);
    const naturalSize = await loadImageSize(imageSource);

    let width = naturalSize.width;
    let height = naturalSize.height;
    if (width > maxWidth) {
      height = Math.max(1, Math.round(height * (maxWidth / width)));
      width = maxWidth;
    }

    const upload = opts?.imageUploadHandler
      ? await opts.imageUploadHandler(file, naturalSize)
      : null;
    if (upload && !upload.assetId.trim()) {
      throw new Error('Image upload returned an empty asset ID');
    }
    if (view.isDestroyed) return false;

    const insertionPos = resolveImageUploadAnchor(view, anchor);
    if (insertionPos === null) return false;

    const imageNode = view.state.schema.nodes.image.create({
      src: upload ? '' : imageSource,
      assetId: upload?.assetId,
      alt: file.name,
      width,
      height,
      // Newly uploaded external assets have no package relationship yet; the
      // collaboration exporter creates it from the manifest. Legacy inline
      // images retain the established temporary-rId behavior.
      rId: upload ? undefined : `rId_img_${Date.now()}_${Math.round(Math.random() * 1e9)}`,
      wrapType: 'inline',
      displayMode: 'inline',
    });
    view.dispatch(
      buildImageInsertionTransaction(view.state, imageNode, insertionPos).scrollIntoView()
    );
    if (!anchor.tracked) {
      anchor.fallbackPos = insertionPos + imageNode.nodeSize;
    }
    view.focus();
    opts?.onInserted?.();
    return true;
  } finally {
    if (objectUrl) URL.revokeObjectURL(objectUrl);
  }
}

/**
 * Read an image `File` (from a file picker or drop), fit it to the page width,
 * and insert it inline at the current selection. This is the single source of
 * truth for "insert an image from a file" — the React and Vue adapters both
 * call it, so insertion behaves identically: no intermediate dialog, the image
 * is sized to fit the column, and it round-trips as an inline drawing (with the
 * `insertion` mark applied in suggesting mode, via {@link insertImageNode}).
 *
 * By default the file is stored as a data URL for backward compatibility. If
 * `imageUploadHandler` is provided, the file is decoded through a temporary
 * object URL, uploaded before insertion, and only the returned opaque asset ID
 * is stored in the node. `onError` reports a failed read, decode, or upload,
 * and `onInserted` runs after the node lands (e.g. to refocus).
 *
 * @public
 */
export async function insertImageFromFile(
  view: EditorView,
  file: File,
  opts?: InsertImageFromFileOptions
): Promise<void> {
  const anchor = createImageUploadAnchor(view);

  try {
    await insertImageFromFileAtAnchor(view, file, opts, anchor);
  } catch (error) {
    opts?.onError?.(error);
  } finally {
    removeImageUploadAnchor(view, anchor);
  }
}

/**
 * Insert a clipboard batch at one mapped paste position while preserving file
 * order even when uploads are deferred.
 *
 * @internal
 */
export async function insertImageFilesAtSelection(
  view: EditorView,
  files: readonly File[],
  opts?: InsertImageFromFileOptions
): Promise<void> {
  const anchor = createImageUploadAnchor(view);
  try {
    for (const file of files) {
      try {
        const inserted = await insertImageFromFileAtAnchor(view, file, opts, anchor);
        if (!inserted && (view.isDestroyed || resolveImageUploadAnchor(view, anchor) === null)) {
          return;
        }
      } catch (error) {
        opts?.onError?.(error);
      }
    }
  } finally {
    removeImageUploadAnchor(view, anchor);
  }
}

/**
 * Change a floating image's wrap layout. `pos` is the PM document position of
 * the image node; `target` is either an OOXML wrap type (square / tight /
 * topAndBottom / behind / inFront / inline) or a directional convenience
 * choice (`squareLeft` / `squareRight`).
 *
 * `opts.initialPositionEmu` is used when promoting an inline image to an
 * anchor — the caller measures the image's current rendered offset relative
 * to the column origin in EMUs and passes it through, so the new float lands
 * exactly where the inline glyph used to sit (matches Word's behavior).
 */
export function setImageWrapType(
  pos: number,
  target: ImageLayoutTarget,
  opts?: SetImageWrapTypeOptions
): Command {
  // Resolve lazily to keep StarterKit -> ImagePasteExtension -> image commands
  // from reading the schema singleton while that same StarterKit is still
  // constructing it.
  return (state, dispatch, view) =>
    singletonManager.getCommands().setImageWrapType(pos, target, opts)(state, dispatch, view);
}

export type {
  AnchorWrapType,
  ImageLayoutTarget,
  SetImageWrapTypeOptions,
} from '../extensions/nodes/ImageExtension';

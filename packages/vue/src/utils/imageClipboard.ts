/**
 * Image clipboard + replace helpers. Pure DOM/PM operations — every
 * function takes the view as a parameter so they're testable and don't
 * close over Vue refs.
 *
 * `copyImageToClipboard` emits both `text/html` (so a subsequent paste
 * re-creates a PM image node via `pasteFromClipboard`) and a `text/plain`
 * fallback. `pasteFromClipboard` walks the clipboard items looking for
 * an image blob first, then an HTML payload carrying our custom data
 * attributes, then plain text.
 */

import type { EditorView } from 'prosemirror-view';
import { makeRevisionInfo } from '@eigenpal/docx-editor-core/prosemirror/plugins';
import { INSERT_IMAGE_MAX_WIDTH_PX } from '@eigenpal/docx-editor-core/prosemirror/commands';
import type { ImageUploadHandler } from '@eigenpal/docx-editor-core/prosemirror/commands';
import type { Node as PMNode } from 'prosemirror-model';
import type { Transaction } from 'prosemirror-state';

// Same cap as file-picker inserts (US Letter content width at 96dpi). Aliased
// from core rather than redeclared so the two cannot drift apart.
const MAX_PASTED_IMAGE_WIDTH = INSERT_IMAGE_MAX_WIDTH_PX;

interface PreparedImage {
  src: string;
  assetId: string | null;
  rId: string | null;
  naturalSize: { width: number; height: number };
  release: () => void;
}

/**
 * Apply the `insertion` mark to the just-replaced image when suggesting
 * mode is active, so clipboard-pasted images round-trip as tracked
 * additions. Run after replaceSelectionWith on the SAME tr.
 */
function tagPastedImageAsInsertion(view: EditorView, tr: Transaction, imageNode: PMNode): void {
  const info = makeRevisionInfo(view.state);
  const insertionType = view.state.schema.marks.insertion;
  if (!info || !insertionType) return;
  // After replaceSelectionWith, the image lives at the prior selection's
  // start position (mapped). The cursor sits just after it.
  const to = tr.selection.from;
  const from = to - imageNode.nodeSize;
  if (from < 0) return;
  tr.addMark(
    from,
    to,
    insertionType.create({
      revisionId: info.revisionId,
      author: info.author,
      date: info.date,
    })
  );
}

/**
 * `ClipboardItem.getBlob(type)` is non-standard but ships in current
 * Chromium/WebKit; the standard alternative is `getType(type)` (also
 * Promise<Blob>). Cast inline so this module stays in lockstep with the
 * pre-extraction call sites.
 */
function getBlob(item: ClipboardItem, type: string): Promise<Blob> {
  if (typeof item.getType === 'function') return item.getType(type);
  return (item as unknown as { getBlob: (t: string) => Promise<Blob> }).getBlob(type);
}

function createImageRelationshipId(): string {
  return `rId_img_${Date.now()}_${Math.round(Math.random() * 1e9)}`;
}

function imageFileFromBlob(blob: Blob, type: string): File {
  if (blob instanceof File) return blob;
  const subtype = type.split('/')[1]?.replace(/[^a-z0-9.+-]/gi, '') || 'bin';
  return new File([blob], `clipboard-image.${subtype}`, { type });
}

async function prepareImage(
  file: File,
  imageUploadHandler?: ImageUploadHandler
): Promise<PreparedImage> {
  let objectUrl: string | null = null;

  try {
    const source = imageUploadHandler
      ? (objectUrl = URL.createObjectURL(file))
      : await blobToDataUrl(file);
    const naturalSize = await loadImageDimensions(source);
    const upload = imageUploadHandler ? await imageUploadHandler(file, naturalSize) : null;

    if (upload && !upload.assetId.trim()) {
      throw new Error('Image upload returned an empty asset ID');
    }

    let released = false;
    return {
      src: upload ? '' : source,
      assetId: upload?.assetId ?? null,
      rId: upload ? null : createImageRelationshipId(),
      naturalSize,
      release: () => {
        if (released || !objectUrl) return;
        released = true;
        URL.revokeObjectURL(objectUrl);
      },
    };
  } catch (error) {
    if (objectUrl) URL.revokeObjectURL(objectUrl);
    throw error;
  }
}

/**
 * Whether a clipboard `img src` may be fetched to obtain bytes for upload.
 *
 * `data:` and `blob:` carry their own bytes, so reading them is local. A remote
 * URL would mean a network request triggered purely by pasting — a zero-click
 * SSRF / IP-leak / tracking-beacon vector — and cross-origin reads would fail
 * CORS regardless. Those are inserted by reference instead.
 */
function isFetchableImageSrc(src: string): boolean {
  const scheme = src.slice(0, src.indexOf(':') + 1).toLowerCase();
  return scheme === 'data:' || scheme === 'blob:';
}

/**
 * Fetch a local image src and run it through the upload handler.
 * Returns `null` when the bytes cannot be read or the upload fails, so the
 * caller can insert by reference instead of dropping the paste entirely.
 */
async function prepareImageFromSrc(
  src: string,
  imageUploadHandler: ImageUploadHandler
): Promise<PreparedImage | null> {
  try {
    const response = await fetch(src);
    if (!response.ok) return null;
    const blob = await response.blob();
    const file = imageFileFromBlob(blob, blob.type || 'image/png');
    return await prepareImage(file, imageUploadHandler);
  } catch {
    return null;
  }
}

/** Insert an image that keeps its original src, with no upload. */
function insertReferencedImage(
  view: EditorView,
  src: string,
  htmlWidth: number,
  htmlHeight: number
): void {
  const imageNode = view.state.schema.nodes.image.create({
    src,
    assetId: null,
    width: htmlWidth || 200,
    height: htmlHeight || 200,
    rId: createImageRelationshipId(),
    wrapType: 'inline',
    displayMode: 'inline',
  });
  const tr = view.state.tr.replaceSelectionWith(imageNode);
  tagPastedImageAsInsertion(view, tr, imageNode);
  view.dispatch(tr);
}

function fitWithinMaxWidth(
  dimensions: { width: number; height: number },
  maxWidth: number
): { width: number; height: number } {
  if (dimensions.width <= maxWidth) return dimensions;
  return {
    width: maxWidth,
    height: Math.max(1, Math.round(dimensions.height * (maxWidth / dimensions.width))),
  };
}

function serializeClipboardImage(node: PMNode): string {
  const image = document.createElement('img');
  const src = node.attrs.src as string;
  const assetId = node.attrs.assetId as string | null;

  // External assets deliberately omit src: serializing an empty attribute can
  // make consumers resolve it to the current page URL.
  if (src) image.setAttribute('src', src);
  if (assetId) image.setAttribute('data-asset-id', assetId);
  image.setAttribute('data-pm-image', 'true');
  image.setAttribute('data-width', String(node.attrs.width ?? ''));
  image.setAttribute('data-height', String(node.attrs.height ?? ''));
  image.setAttribute('data-wrap-type', String(node.attrs.wrapType ?? ''));
  image.setAttribute('data-display-mode', String(node.attrs.displayMode ?? ''));
  image.setAttribute('data-rid', String(node.attrs.rId ?? ''));

  // DOM serialization escapes every opaque/file-derived attribute without
  // evaluating clipboard HTML or interpolating untrusted values into markup.
  return new XMLSerializer().serializeToString(image);
}

interface ClipboardImageAttrs {
  src: string;
  assetId: string | null;
  width: number;
  height: number;
}

function decodeHtmlAttribute(value: string): string {
  return value.replace(/&(#x[\da-f]+|#\d+|amp|lt|gt|quot|apos);/gi, (entity, name: string) => {
    const normalized = name.toLowerCase();
    if (normalized === 'amp') return '&';
    if (normalized === 'lt') return '<';
    if (normalized === 'gt') return '>';
    if (normalized === 'quot') return '"';
    if (normalized === 'apos') return "'";

    const radix = normalized.startsWith('#x') ? 16 : 10;
    const digits = normalized.slice(radix === 16 ? 2 : 1);
    const codePoint = Number.parseInt(digits, radix);
    return Number.isFinite(codePoint) && codePoint <= 0x10ffff
      ? String.fromCodePoint(codePoint)
      : entity;
  });
}

function parseImageAttributes(html: string): Map<string, string> | null {
  const imageTag = html.match(/<img\b[^>]*>/i)?.[0];
  if (!imageTag) return null;

  const attributes = new Map<string, string>();
  const attributePattern = /([^\s=/>]+)\s*(?:=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'=<>`]+)))?/g;
  for (const match of imageTag.slice(4).matchAll(attributePattern)) {
    const name = match[1].toLowerCase();
    const value = match[2] ?? match[3] ?? match[4] ?? '';
    attributes.set(name, decodeHtmlAttribute(value));
  }
  return attributes;
}

function parseClipboardImage(html: string): ClipboardImageAttrs | null {
  // Read only the first image tag's attributes. Avoid constructing an HTML
  // document: detached HTML parsers may still initiate remote image loads.
  const attributes = parseImageAttributes(html);
  if (!attributes) return null;

  const rawAssetId = attributes.get('data-asset-id');
  const assetId = rawAssetId?.trim() ? rawAssetId : null;
  const src = assetId ? '' : (attributes.get('src') ?? '');
  if (!assetId && !src) return null;

  const rawWidth = Number(attributes.get('data-width'));
  const rawHeight = Number(attributes.get('data-height'));
  return {
    src,
    assetId,
    width: Number.isFinite(rawWidth) && rawWidth > 0 ? rawWidth : 0,
    height: Number.isFinite(rawHeight) && rawHeight > 0 ? rawHeight : 0,
  };
}

export function copyImageToClipboard(view: EditorView, pmPos: number): void {
  const node = view.state.doc.nodeAt(pmPos);
  if (!node || node.type.name !== 'image') return;

  const imgHtml = serializeClipboardImage(node);

  const clipboardItem = new ClipboardItem({
    'text/html': new Blob([imgHtml], { type: 'text/html' }),
    'text/plain': new Blob(['[image]'], { type: 'text/plain' }),
  });
  navigator.clipboard.write([clipboardItem]).catch(() => {
    // Fallback: at least copy as HTML
    const ta = document.createElement('textarea');
    ta.value = imgHtml;
    document.body.appendChild(ta);
    ta.select();
    document.execCommand('copy');
    document.body.removeChild(ta);
  });
}

export function blobToDataUrl(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as string);
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

export function loadImageDimensions(src: string): Promise<{ width: number; height: number }> {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => resolve({ width: img.naturalWidth, height: img.naturalHeight });
    img.onerror = () => resolve({ width: 200, height: 200 });
    img.src = src;
  });
}

export async function pasteFromClipboard(
  view: EditorView,
  imageUploadHandler?: ImageUploadHandler
): Promise<void> {
  try {
    const items = await navigator.clipboard.read();
    for (const item of items) {
      const imageType = item.types.find((t) => t.startsWith('image/'));
      if (imageType) {
        const blob = await getBlob(item, imageType);
        const file = imageFileFromBlob(blob, imageType);
        const prepared = await prepareImage(file, imageUploadHandler);
        try {
          if (view.isDestroyed) return;
          const dimensions = fitWithinMaxWidth(prepared.naturalSize, MAX_PASTED_IMAGE_WIDTH);
          const imageNode = view.state.schema.nodes.image.create({
            src: prepared.src,
            assetId: prepared.assetId,
            width: dimensions.width,
            height: dimensions.height,
            rId: prepared.rId,
            wrapType: 'inline',
            displayMode: 'inline',
          });
          const tr = view.state.tr.replaceSelectionWith(imageNode);
          tagPastedImageAsInsertion(view, tr, imageNode);
          view.dispatch(tr);
        } finally {
          prepared.release();
        }
        return;
      }

      if (item.types.includes('text/html')) {
        const htmlBlob = await getBlob(item, 'text/html');
        const html = await htmlBlob.text();
        const clipboardImage = parseClipboardImage(html);
        if (clipboardImage) {
          const { src, assetId, width: htmlWidth, height: htmlHeight } = clipboardImage;

          if (assetId) {
            const imageNode = view.state.schema.nodes.image.create({
              src: '',
              assetId,
              width: htmlWidth || 200,
              height: htmlHeight || 200,
              rId: null,
              wrapType: 'inline',
              displayMode: 'inline',
            });
            const tr = view.state.tr.replaceSelectionWith(imageNode);
            tagPastedImageAsInsertion(view, tr, imageNode);
            view.dispatch(tr);
          } else if (imageUploadHandler && isFetchableImageSrc(src)) {
            // Only same-document sources (data:/blob:) are fetched. A remote
            // src would be a zero-click external request on paste — an
            // SSRF/IP-leak/tracking-beacon vector — and would fail CORS
            // anyway. Anything else falls through to the reference-only insert
            // below rather than dropping the paste.
            const prepared = await prepareImageFromSrc(src, imageUploadHandler);
            if (!prepared) {
              insertReferencedImage(view, src, htmlWidth, htmlHeight);
              return;
            }
            try {
              if (view.isDestroyed) return;
              const dimensions =
                htmlWidth > 0 && htmlHeight > 0
                  ? { width: htmlWidth, height: htmlHeight }
                  : fitWithinMaxWidth(prepared.naturalSize, MAX_PASTED_IMAGE_WIDTH);
              const imageNode = view.state.schema.nodes.image.create({
                src: prepared.src,
                assetId: prepared.assetId,
                width: dimensions.width,
                height: dimensions.height,
                rId: prepared.rId,
                wrapType: 'inline',
                displayMode: 'inline',
              });
              const tr = view.state.tr.replaceSelectionWith(imageNode);
              tagPastedImageAsInsertion(view, tr, imageNode);
              view.dispatch(tr);
            } finally {
              prepared.release();
            }
          } else {
            insertReferencedImage(view, src, htmlWidth, htmlHeight);
          }
          return;
        }
      }

      if (item.types.includes('text/plain')) {
        const textBlob = await getBlob(item, 'text/plain');
        const text = await textBlob.text();
        if (text && text !== '[image]') {
          const { from } = view.state.selection;
          view.dispatch(view.state.tr.insertText(text, from));
        }
        return;
      }
    }
  } catch {
    // Fallback for browsers without clipboard API
    try {
      const text = await navigator.clipboard?.readText();
      if (text && text !== '[image]') {
        const { from } = view.state.selection;
        view.dispatch(view.state.tr.insertText(text, from));
      }
    } catch {
      // Clipboard access denied.
    }
  }
}

export function triggerReplaceImage(
  view: EditorView,
  pmPos: number,
  imageUploadHandler?: ImageUploadHandler
): void {
  const node = view.state.doc.nodeAt(pmPos);
  if (!node || node.type.name !== 'image') return;

  const input = document.createElement('input');
  input.type = 'file';
  input.accept = 'image/*';
  input.onchange = async () => {
    const file = input.files?.[0];
    if (!file) return;

    try {
      const prepared = await prepareImage(file, imageUploadHandler);
      try {
        if (view.isDestroyed || view.state.doc.nodeAt(pmPos) !== node) return;

        // Keep existing dimensions unless the aspect ratio is wildly different;
        // scale the new image to fit within the old bounding box.
        const oldW = (node.attrs.width as number) || prepared.naturalSize.width;
        const oldH = (node.attrs.height as number) || prepared.naturalSize.height;
        const scale = Math.min(
          oldW / prepared.naturalSize.width,
          oldH / prepared.naturalSize.height
        );
        const newW = Math.max(1, Math.round(prepared.naturalSize.width * scale));
        const newH = Math.max(1, Math.round(prepared.naturalSize.height * scale));

        const tr = view.state.tr.setNodeMarkup(pmPos, undefined, {
          ...node.attrs,
          src: prepared.src,
          assetId: prepared.assetId,
          width: newW,
          height: newH,
          rId: prepared.rId,
        });
        view.dispatch(tr);
      } finally {
        prepared.release();
      }
    } catch {
      // File decode/upload can fail, or the position may have changed.
    }
  };
  input.click();
}

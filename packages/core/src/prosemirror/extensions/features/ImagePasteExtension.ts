/**
 * Image Paste Extension — handles image files pasted from the clipboard
 *
 * When an image file is present on the clipboard, this intercepts the paste,
 * reads the image data, and inserts an image node instead of a file icon.
 */

import { Plugin } from 'prosemirror-state';
import { createExtension } from '../create';
import type { ExtensionRuntime } from '../types';
import { getClipboardImageFiles } from '../../../utils/clipboard';
import { insertImageFilesAtSelection, type ImageUploadHandler } from '../../commands/image';
import { imageUploadAnchorPlugin } from '../../commands/imageUploadAnchor';

type ImagePasteOptions = Record<string, unknown> & {
  imageUploadHandler?: ImageUploadHandler;
};

export const ImagePasteExtension = createExtension<ImagePasteOptions>({
  name: 'imagePaste',
  onSchemaReady(_ctx, options): ExtensionRuntime {
    const plugin = new Plugin({
      props: {
        handleDOMEvents: {
          paste(view, event) {
            const clipboardEvent = event as ClipboardEvent;
            const imageFiles = getClipboardImageFiles(clipboardEvent.clipboardData);

            if (imageFiles.length === 0) {
              return false;
            }

            if (!view.state.schema.nodes.image) {
              return false;
            }

            clipboardEvent.preventDefault();
            void insertImageFilesAtSelection(view, imageFiles, {
              imageUploadHandler: options.imageUploadHandler,
            });
            return true;
          },
        },
      },
    });

    return { plugins: [imageUploadAnchorPlugin, plugin] };
  },
});

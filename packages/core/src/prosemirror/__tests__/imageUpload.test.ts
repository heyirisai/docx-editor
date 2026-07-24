import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, beforeAll, describe, expect, mock, test } from 'bun:test';
import { EditorState, TextSelection, type Plugin, type Transaction } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { insertImageFromFile, type ImageUploadHandler } from '../commands/image';
import { ImagePasteExtension } from '../extensions/features/ImagePasteExtension';
import { singletonManager } from '../schema';

beforeAll(() => GlobalRegistrator.register());
afterAll(() => GlobalRegistrator.unregister());

const schema = singletonManager.getSchema();

function makeView(options?: { text?: string; selectionPos?: number; plugins?: Plugin[] }) {
  const paragraph = schema.nodes.paragraph.create(
    null,
    options?.text ? schema.text(options.text) : undefined
  );
  const doc = schema.nodes.doc.create(null, [paragraph]);
  const view = {
    state: EditorState.create({
      schema,
      doc,
      selection:
        options?.selectionPos === undefined
          ? undefined
          : TextSelection.create(doc, options.selectionPos),
      plugins: options?.plugins,
    }),
    dispatch(transaction: Transaction) {
      view.state = view.state.apply(transaction);
    },
    focus: mock(() => undefined),
    isDestroyed: false,
  };
  return view as unknown as EditorView & { state: EditorState };
}

function installImageDecoder() {
  const originalImage = Object.getOwnPropertyDescriptor(globalThis, 'Image');
  const originalCreateObjectUrl = Object.getOwnPropertyDescriptor(URL, 'createObjectURL');
  const originalRevokeObjectUrl = Object.getOwnPropertyDescriptor(URL, 'revokeObjectURL');
  const revokeObjectUrl = mock((_url: string) => undefined);

  class FakeImage {
    naturalWidth = 1200;
    naturalHeight = 600;
    onload: (() => void) | null = null;
    onerror: (() => void) | null = null;

    set src(_value: string) {
      queueMicrotask(() => this.onload?.());
    }
  }

  Object.defineProperty(globalThis, 'Image', {
    configurable: true,
    value: FakeImage,
  });
  Object.defineProperty(URL, 'createObjectURL', {
    configurable: true,
    value: () => 'blob:local-image',
  });
  Object.defineProperty(URL, 'revokeObjectURL', {
    configurable: true,
    value: revokeObjectUrl,
  });

  return {
    revokeObjectUrl,
    restore() {
      if (originalImage) Object.defineProperty(globalThis, 'Image', originalImage);
      if (originalCreateObjectUrl) {
        Object.defineProperty(URL, 'createObjectURL', originalCreateObjectUrl);
      }
      if (originalRevokeObjectUrl) {
        Object.defineProperty(URL, 'revokeObjectURL', originalRevokeObjectUrl);
      }
    },
  };
}

function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((done) => {
    resolve = done;
  });
  return { promise, resolve };
}

async function waitFor(predicate: () => boolean): Promise<void> {
  for (let attempt = 0; attempt < 20; attempt += 1) {
    if (predicate()) return;
    await new Promise<void>((resolve) => queueMicrotask(resolve));
  }
  throw new Error('Timed out waiting for deferred image insertion');
}

function imagePastePlugins(imageUploadHandler?: ImageUploadHandler): Plugin[] {
  return (
    ImagePasteExtension({ imageUploadHandler }).onSchemaReady({
      schema,
      manager: singletonManager,
    }).plugins ?? []
  );
}

function expectImageBetween(
  view: EditorView & { state: EditorState },
  before: string,
  after: string
) {
  const paragraph = view.state.doc.firstChild;
  expect(paragraph?.childCount).toBe(3);
  expect(paragraph?.child(0).text).toBe(before);
  expect(paragraph?.child(1).type.name).toBe('image');
  expect(paragraph?.child(2).text).toBe(after);
}

describe('insertImageFromFile external upload', () => {
  test('publishes only the opaque asset identity after upload succeeds', async () => {
    const decoder = installImageDecoder();

    try {
      const view = makeView({ plugins: imagePastePlugins() });
      const upload = mock<ImageUploadHandler>(async (_file, dimensions) => {
        expect(dimensions).toEqual({ width: 1200, height: 600 });
        return { assetId: 'opaque-asset-id' };
      });

      await insertImageFromFile(
        view,
        new File([Uint8Array.of(1, 2, 3)], 'logo.png', {
          type: 'image/png',
        }),
        { imageUploadHandler: upload }
      );

      const image = view.state.doc.nodeAt(1);
      expect(upload).toHaveBeenCalledTimes(1);
      expect(image?.type.name).toBe('image');
      expect(image?.attrs.src).toBe('');
      expect(image?.attrs.assetId).toBe('opaque-asset-id');
      expect(image?.attrs.rId).toBeNull();
      expect(image?.attrs.width).toBe(612);
      expect(image?.attrs.height).toBe(306);
      expect(decoder.revokeObjectUrl).toHaveBeenCalledWith('blob:local-image');
    } finally {
      decoder.restore();
    }
  });

  test('keeps a deferred upload at its mapped invocation position', async () => {
    const decoder = installImageDecoder();
    const uploadStarted = deferred<void>();
    const uploadResult = deferred<{ assetId: string }>();
    const upload = mock<ImageUploadHandler>(async () => {
      uploadStarted.resolve();
      return uploadResult.promise;
    });

    try {
      const view = makeView({
        text: 'AB',
        selectionPos: 2,
        plugins: imagePastePlugins(),
      });
      const insertion = insertImageFromFile(
        view,
        new File([Uint8Array.of(1)], 'deferred.png', { type: 'image/png' }),
        { imageUploadHandler: upload }
      );
      await uploadStarted.promise;

      const concurrent = view.state.tr.insertText('X', 1);
      concurrent.setSelection(TextSelection.atEnd(concurrent.doc));
      view.dispatch(concurrent);

      uploadResult.resolve({ assetId: 'deferred-asset' });
      await insertion;

      expectImageBetween(view, 'XA', 'B');
      expect(view.state.doc.firstChild?.child(1).attrs.assetId).toBe('deferred-asset');
    } finally {
      decoder.restore();
    }
  });

  test('keeps a deferred clipboard upload at the original paste position', async () => {
    const decoder = installImageDecoder();
    const uploadStarted = deferred<void>();
    const uploadResult = deferred<{ assetId: string }>();
    const upload = mock<ImageUploadHandler>(async () => {
      uploadStarted.resolve();
      return uploadResult.promise;
    });
    const plugins = imagePastePlugins(upload);

    try {
      const view = makeView({ text: 'AB', selectionPos: 2, plugins });
      const file = new File([Uint8Array.of(2)], 'clipboard.png', { type: 'image/png' });
      const pastePlugin = plugins.find((plugin) => plugin.props.handleDOMEvents?.paste);
      const pasteHandler = pastePlugin?.props.handleDOMEvents?.paste;
      const preventDefault = mock(() => undefined);
      const event = {
        clipboardData: {
          items: [
            {
              kind: 'file',
              type: 'image/png',
              getAsFile: () => file,
            },
          ],
          files: [],
        },
        preventDefault,
      } as unknown as ClipboardEvent;

      if (!pastePlugin || !pasteHandler) throw new Error('Image paste handler was not registered');
      expect(pasteHandler.call(pastePlugin, view, event)).toBe(true);
      await uploadStarted.promise;

      const concurrent = view.state.tr.insertText('X', 1);
      concurrent.setSelection(TextSelection.atEnd(concurrent.doc));
      view.dispatch(concurrent);

      uploadResult.resolve({ assetId: 'clipboard-asset' });
      await waitFor(() => view.state.doc.firstChild?.childCount === 3);

      expect(preventDefault).toHaveBeenCalledTimes(1);
      expectImageBetween(view, 'XA', 'B');
      expect(view.state.doc.firstChild?.child(1).attrs.assetId).toBe('clipboard-asset');
    } finally {
      decoder.restore();
    }
  });
});

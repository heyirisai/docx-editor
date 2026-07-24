import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { schema } from '@eigenpal/docx-editor-core/prosemirror';
import type { ImageUploadHandler } from '@eigenpal/docx-editor-core/prosemirror/commands';
import { afterAll, beforeAll, expect, mock, test } from 'bun:test';
import { NodeSelection, EditorState, type Transaction } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { copyImageToClipboard, pasteFromClipboard, triggerReplaceImage } from './imageClipboard';

beforeAll(() => GlobalRegistrator.register());
afterAll(() => GlobalRegistrator.unregister());

function makeImageView(attrs?: Record<string, unknown>) {
  const image = schema.nodes.image.create({
    src: '',
    assetId: 'old-asset',
    width: 300,
    height: 150,
    rId: null,
    wrapType: 'inline',
    displayMode: 'inline',
    ...attrs,
  });
  const doc = schema.nodes.doc.create(null, [schema.nodes.paragraph.create(null, [image])]);
  const view = {
    state: EditorState.create({
      schema,
      doc,
      selection: NodeSelection.create(doc, 1),
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

test('context-menu image paste publishes only the uploaded asset identity', async () => {
  const decoder = installImageDecoder();
  const originalClipboard = Object.getOwnPropertyDescriptor(navigator, 'clipboard');
  const blob = new Blob([Uint8Array.of(1, 2, 3)], { type: 'image/png' });
  Object.defineProperty(navigator, 'clipboard', {
    configurable: true,
    value: {
      read: async () => [
        {
          types: ['image/png'],
          getBlob: async () => blob,
        },
      ],
    },
  });
  const upload = mock<ImageUploadHandler>(async (_file, dimensions) => {
    expect(dimensions).toEqual({ width: 1200, height: 600 });
    return { assetId: 'new-asset' };
  });

  try {
    const view = makeImageView();
    await pasteFromClipboard(view, upload);

    const image = view.state.doc.nodeAt(1);
    expect(upload).toHaveBeenCalledTimes(1);
    expect(image?.attrs.src).toBe('');
    expect(image?.attrs.assetId).toBe('new-asset');
    expect(image?.attrs.rId).toBeNull();
    expect(image?.attrs.width).toBe(612);
    expect(image?.attrs.height).toBe(306);
    expect(decoder.revokeObjectUrl).toHaveBeenCalledWith('blob:local-image');
  } finally {
    decoder.restore();
    if (originalClipboard) {
      Object.defineProperty(navigator, 'clipboard', originalClipboard);
    } else {
      delete (navigator as { clipboard?: Clipboard }).clipboard;
    }
  }
});

test('copy and paste preserve an external asset identity without exposing a source URL', async () => {
  const originalClipboard = Object.getOwnPropertyDescriptor(navigator, 'clipboard');
  const writtenItems: ClipboardItem[] = [];
  Object.defineProperty(navigator, 'clipboard', {
    configurable: true,
    value: {
      read: async () => writtenItems,
      write: async (items: ClipboardItem[]) => {
        writtenItems.push(...items);
      },
    },
  });

  try {
    const source = makeImageView({
      src: '',
      assetId: 'opaque<&"asset',
    });
    copyImageToClipboard(source, 1);
    await Promise.resolve();

    expect(writtenItems).toHaveLength(1);
    const html = await (await writtenItems[0].getType('text/html')).text();
    expect(html).toContain('data-asset-id=');
    expect(html).not.toContain('src=');
    expect(html).not.toContain('blob:');
    expect(html).not.toContain('data:image');

    const target = makeImageView();
    const upload = mock<ImageUploadHandler>(async () => ({ assetId: 'unexpected-upload' }));
    await pasteFromClipboard(target, upload);

    const image = target.state.doc.nodeAt(1);
    expect(upload).not.toHaveBeenCalled();
    expect(image?.attrs.src).toBe('');
    expect(image?.attrs.assetId).toBe('opaque<&"asset');
    expect(image?.attrs.rId).toBeNull();
  } finally {
    if (originalClipboard) {
      Object.defineProperty(navigator, 'clipboard', originalClipboard);
    } else {
      delete (navigator as { clipboard?: Clipboard }).clipboard;
    }
  }
});

test('replace image swaps the old asset identity without retaining inline bytes', async () => {
  const decoder = installImageDecoder();
  const originalClick = HTMLInputElement.prototype.click;
  const file = new File([Uint8Array.of(4, 5, 6)], 'replacement.png', {
    type: 'image/png',
  });
  let change: Promise<void> | undefined;
  HTMLInputElement.prototype.click = function click() {
    Object.defineProperty(this, 'files', {
      configurable: true,
      value: [file],
    });
    change = this.onchange?.(new Event('change')) as Promise<void> | undefined;
  };
  const upload = mock<ImageUploadHandler>(async (_file, dimensions) => {
    expect(dimensions).toEqual({ width: 1200, height: 600 });
    return { assetId: 'replacement-asset' };
  });

  try {
    const view = makeImageView();
    triggerReplaceImage(view, 1, upload);
    await change;

    const image = view.state.doc.nodeAt(1);
    expect(upload).toHaveBeenCalledTimes(1);
    expect(image?.attrs.src).toBe('');
    expect(image?.attrs.assetId).toBe('replacement-asset');
    expect(image?.attrs.rId).toBeNull();
    expect(image?.attrs.width).toBe(300);
    expect(image?.attrs.height).toBe(150);
    expect(decoder.revokeObjectUrl).toHaveBeenCalledWith('blob:local-image');
  } finally {
    HTMLInputElement.prototype.click = originalClick;
    decoder.restore();
  }
});

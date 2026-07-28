import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, beforeAll, beforeEach, describe, expect, test } from 'bun:test';
import { Schema } from 'prosemirror-model';
import { EditorState } from 'prosemirror-state';
import { toFlowBlocks } from '../../layout-bridge/toFlowBlocks';
import type { ImageBlock, ImageFragment, ImageMeasure, Page } from '../../layout-engine/types';
import { buildBlockLookup } from '../index';
import { LazyImageAssetLoader } from '../imageAssets';
import { renderPages } from '../renderPage';

type ObserverCallback = ConstructorParameters<typeof IntersectionObserver>[0];

const schema = new Schema({
  nodes: {
    doc: { content: 'block+' },
    paragraph: { content: 'text*', group: 'block' },
    image: {
      group: 'block',
      attrs: {
        src: { default: '' },
        assetId: { default: null },
        width: { default: 100 },
        height: { default: 100 },
        wrapType: { default: null },
      },
    },
    text: {},
  },
});

let originalIntersectionObserver: typeof globalThis.IntersectionObserver | undefined;
let observerCallbacks: ObserverCallback[] = [];

beforeAll(() => {
  GlobalRegistrator.register();
  originalIntersectionObserver = globalThis.IntersectionObserver;
  globalThis.IntersectionObserver = class {
    constructor(callback: ObserverCallback) {
      observerCallbacks.push(callback);
    }
    observe() {}
    unobserve() {}
    disconnect() {}
    takeRecords() {
      return [];
    }
  } as unknown as typeof globalThis.IntersectionObserver;
});

beforeEach(() => {
  observerCallbacks = [];
  document.body.innerHTML = '';
});

afterAll(() => {
  globalThis.IntersectionObserver = originalIntersectionObserver!;
  GlobalRegistrator.unregister();
});

function pageFor(block: ImageBlock): Page {
  const fragment: ImageFragment = {
    kind: 'image',
    blockId: block.id,
    x: 0,
    y: 0,
    width: block.width,
    height: block.height,
    pmStart: block.pmStart,
    pmEnd: block.pmEnd,
  };
  return {
    number: 1,
    fragments: [fragment],
    margins: { top: 0, right: 0, bottom: 0, left: 0 },
    size: { w: 816, h: 1056 },
  };
}

function intersectImage(image: HTMLImageElement): void {
  observerCallbacks[0]?.(
    [
      {
        target: image,
        isIntersecting: true,
        intersectionRatio: 1,
      } as unknown as IntersectionObserverEntry,
    ],
    {} as IntersectionObserver
  );
}

async function nextMicrotask(): Promise<void> {
  await Promise.resolve();
  await Promise.resolve();
}

describe('external image layout stability', () => {
  test('typing after an image preserves its rendered DOM and resolver lease', async () => {
    const doc = schema.node('doc', null, [
      schema.node('image', {
        assetId: 'asset-1',
        width: 120,
        height: 80,
      }),
      schema.node('paragraph', null, [schema.text('Edit me')]),
    ]);
    const state = EditorState.create({ doc });
    const beforeImage = toFlowBlocks(state.doc)[0] as ImageBlock;
    const measure: ImageMeasure = { kind: 'image', width: 120, height: 80 };
    let resolveCalls = 0;
    const loader = new LazyImageAssetLoader(async () => {
      resolveCalls++;
      return { src: 'blob:asset-1' };
    });
    const container = document.createElement('div');
    document.body.appendChild(container);

    renderPages([pageFor(beforeImage)], container, {
      document,
      forcePageVirtualization: true,
      imageAssetLoader: loader,
      blockLookup: buildBlockLookup([beforeImage], [measure]),
    });
    const renderedBefore = container.querySelector<HTMLImageElement>(
      'img[data-asset-id="asset-1"]'
    );
    expect(renderedBefore).not.toBeNull();
    intersectImage(renderedBefore!);
    await nextMicrotask();
    expect(resolveCalls).toBe(1);

    const afterState = state.apply(state.tr.insertText('x', state.doc.content.size - 1));
    const afterImage = toFlowBlocks(afterState.doc)[0] as ImageBlock;
    renderPages([pageFor(afterImage)], container, {
      document,
      forcePageVirtualization: true,
      imageAssetLoader: loader,
      blockLookup: buildBlockLookup([afterImage], [measure]),
    });
    const renderedAfter = container.querySelector<HTMLImageElement>('img[data-asset-id="asset-1"]');

    expect(afterImage.id).toBe(beforeImage.id);
    expect(renderedAfter).toBe(renderedBefore);
    expect(resolveCalls).toBe(1);
    loader.dispose();
  });
});

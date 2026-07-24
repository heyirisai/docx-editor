import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, beforeAll, beforeEach, describe, expect, test } from 'bun:test';
import { LazyImageAssetLoader, setImageAssetSource, type ImageAssetResolver } from '../imageAssets';

type ObserverCallback = ConstructorParameters<typeof IntersectionObserver>[0];

let originalIntersectionObserver: typeof globalThis.IntersectionObserver | undefined;
let observerCallback: ObserverCallback | undefined;

beforeAll(() => {
  GlobalRegistrator.register();
  originalIntersectionObserver = globalThis.IntersectionObserver;
  globalThis.IntersectionObserver = class {
    constructor(callback: ObserverCallback) {
      observerCallback = callback;
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
  observerCallback = undefined;
  document.body.innerHTML = '';
});

afterAll(() => {
  globalThis.IntersectionObserver = originalIntersectionObserver!;
  GlobalRegistrator.unregister();
});

function intersect(images: HTMLImageElement[]): void {
  observerCallback?.(
    images.map(
      (target) =>
        ({
          target,
          isIntersecting: true,
          intersectionRatio: 1,
        }) as unknown as IntersectionObserverEntry
    ),
    {} as IntersectionObserver
  );
}

async function nextMicrotask(): Promise<void> {
  await Promise.resolve();
  await Promise.resolve();
}

describe('LazyImageAssetLoader', () => {
  test('does not resolve or assign a source until the image intersects', async () => {
    let calls = 0;
    const resolver: ImageAssetResolver = async () => {
      calls++;
      return { src: 'blob:asset-1' };
    };
    const loader = new LazyImageAssetLoader(resolver);
    const img = document.createElement('img');

    setImageAssetSource(img, { assetId: 'asset-1', width: 120, height: 80 }, loader);

    expect(calls).toBe(0);
    expect(img.hasAttribute('src')).toBe(false);
    expect(img.dataset.assetId).toBe('asset-1');

    intersect([img]);
    await nextMicrotask();

    expect(calls).toBe(1);
    expect(img.src).toContain('blob:asset-1');
    loader.dispose();
  });

  test('limits concurrent resolve and decode work to four images', async () => {
    let active = 0;
    let peak = 0;
    const pending: Array<() => void> = [];
    const resolver: ImageAssetResolver = ({ assetId }) =>
      new Promise((resolve) => {
        active++;
        peak = Math.max(peak, active);
        pending.push(() => {
          active--;
          resolve({ src: `blob:${assetId}` });
        });
      });
    const loader = new LazyImageAssetLoader(resolver, { maxConcurrent: 4 });
    const images = Array.from({ length: 12 }, (_, index) => {
      const img = document.createElement('img');
      setImageAssetSource(img, { assetId: `asset-${index}`, width: 100, height: 100 }, loader);
      return img;
    });

    intersect(images);
    await nextMicrotask();
    expect(active).toBe(4);
    expect(peak).toBe(4);

    while (pending.length > 0) {
      pending.shift()?.();
      await nextMicrotask();
    }

    expect(peak).toBe(4);
    expect(images.every((img) => img.hasAttribute('src'))).toBe(true);
    loader.dispose();
  });

  test('aborts unresolved work and releases resolved leases when a subtree is removed', async () => {
    let aborted = false;
    let released = 0;
    const resolver: ImageAssetResolver = ({ assetId, signal }) => {
      if (assetId === 'pending') {
        return new Promise((_resolve, reject) => {
          signal.addEventListener('abort', () => {
            aborted = true;
            reject(new DOMException('Aborted', 'AbortError'));
          });
        });
      }
      return Promise.resolve({
        src: `blob:${assetId}`,
        release: () => {
          released++;
        },
      });
    };
    const loader = new LazyImageAssetLoader(resolver);
    const root = document.createElement('div');
    const resolvedImage = document.createElement('img');
    const pendingImage = document.createElement('img');
    root.append(resolvedImage, pendingImage);
    setImageAssetSource(resolvedImage, { assetId: 'resolved' }, loader);
    setImageAssetSource(pendingImage, { assetId: 'pending' }, loader);

    intersect([resolvedImage]);
    await nextMicrotask();
    intersect([pendingImage]);
    await nextMicrotask();
    loader.releaseSubtree(root);
    await nextMicrotask();

    expect(aborted).toBe(true);
    expect(released).toBe(1);
    expect(resolvedImage.hasAttribute('src')).toBe(false);
    loader.dispose();
  });

  test('eagerly resolves every observed image before a print snapshot', async () => {
    const releases: Array<() => void> = [];
    const resolver: ImageAssetResolver = ({ assetId }) =>
      new Promise((resolve) => {
        releases.push(() => resolve({ src: `blob:${assetId}` }));
      });
    const loader = new LazyImageAssetLoader(resolver, { maxConcurrent: 2 });
    const root = document.createElement('div');
    const images = ['asset-1', 'asset-2'].map((assetId) => {
      const image = document.createElement('img');
      root.appendChild(image);
      setImageAssetSource(image, { assetId }, loader);
      return image;
    });

    const ready = loader.resolveSubtree(root);
    await nextMicrotask();

    expect(releases).toHaveLength(2);
    expect(images.every((image) => !image.hasAttribute('src'))).toBe(true);
    for (const release of releases) release();
    await ready;

    expect(images.map((image) => image.getAttribute('src'))).toEqual([
      'blob:asset-1',
      'blob:asset-2',
    ]);
    loader.dispose();
  });

  test('retains resolved leases until a print consumer releases its hold', async () => {
    let released = 0;
    const loader = new LazyImageAssetLoader(async () => ({
      src: 'blob:asset-1',
      release: () => {
        released++;
      },
    }));
    const root = document.createElement('div');
    const image = document.createElement('img');
    root.appendChild(image);
    setImageAssetSource(image, { assetId: 'asset-1' }, loader);

    const releaseHold = loader.holdLeases();
    await loader.resolveSubtree(root);
    loader.releaseSubtree(root);

    expect(released).toBe(0);
    releaseHold();
    expect(released).toBe(1);
    loader.dispose();
  });

  test('waits for a fresh settlement when print retries an errored image', async () => {
    let calls = 0;
    let resolveRetry: ((value: { src: string }) => void) | undefined;
    const loader = new LazyImageAssetLoader(async () => {
      calls++;
      if (calls === 1) throw new Error('temporary failure');
      return new Promise<{ src: string }>((resolve) => {
        resolveRetry = resolve;
      });
    });
    const root = document.createElement('div');
    const image = document.createElement('img');
    root.appendChild(image);
    setImageAssetSource(image, { assetId: 'asset-1' }, loader);

    await loader.resolveSubtree(root);
    expect(image.dataset.assetState).toBe('error');

    let retrySettled = false;
    const retry = loader.resolveSubtree(root).then(() => {
      retrySettled = true;
    });
    await nextMicrotask();

    expect(calls).toBe(2);
    expect(retrySettled).toBe(false);

    resolveRetry?.({ src: 'blob:asset-1-retry' });
    await retry;
    expect(retrySettled).toBe(true);
    expect(image.dataset.assetState).toBe('ready');
    expect(image.getAttribute('src')).toBe('blob:asset-1-retry');
    loader.dispose();
  });

  test('does not requeue a ready image or overwrite its lease on repeated print', async () => {
    let calls = 0;
    let released = 0;
    const loader = new LazyImageAssetLoader(async () => {
      calls++;
      return {
        src: 'blob:asset-1',
        release: () => {
          released++;
        },
      };
    });
    const root = document.createElement('div');
    const image = document.createElement('img');
    root.appendChild(image);
    setImageAssetSource(image, { assetId: 'asset-1' }, loader);

    await loader.resolveSubtree(root);
    await loader.resolveSubtree(root);

    expect(calls).toBe(1);
    loader.releaseSubtree(root);
    expect(released).toBe(1);
    loader.dispose();
  });
});

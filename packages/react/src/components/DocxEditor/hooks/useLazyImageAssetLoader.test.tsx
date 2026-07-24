import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, describe, expect, mock, test } from 'bun:test';
import type { ImageAssetResolver } from '@eigenpal/docx-editor-core/layout-painter';
import { cleanup, render } from '@testing-library/react';
import { StrictMode } from 'react';
import { useLazyImageAssetLoader } from './useLazyImageAssetLoader';

let originalIntersectionObserver: typeof globalThis.IntersectionObserver | undefined;

beforeAll(() => {
  GlobalRegistrator.register();
  originalIntersectionObserver = globalThis.IntersectionObserver;
  globalThis.IntersectionObserver = undefined as never;
});

afterAll(() => {
  globalThis.IntersectionObserver = originalIntersectionObserver!;
  GlobalRegistrator.unregister();
});

afterEach(() => {
  cleanup();
});

describe('useLazyImageAssetLoader', () => {
  test('keeps the loader active through the StrictMode effect replay', async () => {
    const resolver = mock<ImageAssetResolver>(async ({ assetId }) => ({
      src: `blob:${assetId}`,
    }));
    let loader: ReturnType<typeof useLazyImageAssetLoader>;

    function Harness() {
      loader = useLazyImageAssetLoader(resolver);
      return null;
    }

    render(
      <StrictMode>
        <Harness />
      </StrictMode>
    );

    const image = document.createElement('img');
    loader!.observe(image, { assetId: 'cover-image' });
    await Promise.resolve();
    await Promise.resolve();

    expect(resolver).toHaveBeenCalledTimes(1);
    expect(image.dataset.assetState).toBe('ready');
  });

  test('releases the active image lease when the component unmounts', async () => {
    const release = mock(() => {});
    const resolver = mock<ImageAssetResolver>(async ({ assetId }) => ({
      src: `blob:${assetId}`,
      release,
    }));
    let loader: ReturnType<typeof useLazyImageAssetLoader>;

    function Harness() {
      loader = useLazyImageAssetLoader(resolver);
      return null;
    }

    const { unmount } = render(<Harness />);
    const image = document.createElement('img');
    loader!.observe(image, { assetId: 'cover-image' });
    await Promise.resolve();
    await Promise.resolve();

    expect(image.dataset.assetState).toBe('ready');

    unmount();
    await Promise.resolve();

    expect(release).toHaveBeenCalledTimes(1);
    expect(image.hasAttribute('src')).toBe(false);
  });

  test('releases the old loader when the resolver changes and keeps the new loader active', async () => {
    const oldRelease = mock(() => {});
    const oldResolver = mock<ImageAssetResolver>(async ({ assetId }) => ({
      src: `blob:old-${assetId}`,
      release: oldRelease,
    }));
    const newResolver = mock<ImageAssetResolver>(async ({ assetId }) => ({
      src: `blob:new-${assetId}`,
    }));
    let loader: ReturnType<typeof useLazyImageAssetLoader>;

    function Harness({ resolver }: { resolver: ImageAssetResolver }) {
      loader = useLazyImageAssetLoader(resolver);
      return null;
    }

    const { rerender, unmount } = render(<Harness resolver={oldResolver} />);
    const oldLoader = loader!;
    const oldImage = document.createElement('img');
    oldLoader.observe(oldImage, { assetId: 'cover-image' });
    await Promise.resolve();
    await Promise.resolve();

    expect(oldImage.dataset.assetState).toBe('ready');

    rerender(<Harness resolver={newResolver} />);
    const newLoader = loader!;
    expect(newLoader).not.toBe(oldLoader);
    await Promise.resolve();

    expect(oldRelease).toHaveBeenCalledTimes(1);
    expect(oldImage.hasAttribute('src')).toBe(false);

    const newImage = document.createElement('img');
    newLoader.observe(newImage, { assetId: 'inline-image' });
    await Promise.resolve();
    await Promise.resolve();

    expect(newResolver).toHaveBeenCalledTimes(1);
    expect(newImage.dataset.assetState).toBe('ready');

    unmount();
    await Promise.resolve();
  });
});

import { afterEach, describe, expect, mock, test } from 'bun:test';
import { createCachedImageAssetResolver } from './cachedImageAssetResolver';

const originalFetch = globalThis.fetch;
const originalCreateObjectURL = URL.createObjectURL;
const originalRevokeObjectURL = URL.revokeObjectURL;

afterEach(() => {
  globalThis.fetch = originalFetch;
  URL.createObjectURL = originalCreateObjectURL;
  URL.revokeObjectURL = originalRevokeObjectURL;
});

describe('createCachedImageAssetResolver', () => {
  test('shares one fetch and keeps the asset warm across released leases', async () => {
    const fetchAsset = mock(async () => new Response(new Blob(['image'])));
    const createObjectURL = mock(() => 'blob:asset-1');
    const revokeObjectURL = mock(() => {});
    globalThis.fetch = fetchAsset as unknown as typeof fetch;
    URL.createObjectURL = createObjectURL;
    URL.revokeObjectURL = revokeObjectURL;

    const cached = createCachedImageAssetResolver((assetId) => `/assets/${assetId}`);
    const [first, second] = await Promise.all([
      cached.resolver({
        assetId: 'asset-1',
        signal: new AbortController().signal,
      }),
      cached.resolver({
        assetId: 'asset-1',
        signal: new AbortController().signal,
      }),
    ]);

    expect(fetchAsset).toHaveBeenCalledTimes(1);
    expect(createObjectURL).toHaveBeenCalledTimes(1);
    expect(first.src).toBe('blob:asset-1');
    expect(second.src).toBe(first.src);

    first.release?.();
    second.release?.();
    const third = await cached.resolver({
      assetId: 'asset-1',
      signal: new AbortController().signal,
    });

    expect(fetchAsset).toHaveBeenCalledTimes(1);
    expect(third.src).toBe(first.src);
    third.release?.();
    cached.dispose();
    expect(revokeObjectURL).toHaveBeenCalledTimes(1);
  });

  test('defers URL revocation until the final active lease is released', async () => {
    globalThis.fetch = mock(
      async () => new Response(new Blob(['image']))
    ) as unknown as typeof fetch;
    URL.createObjectURL = mock(() => 'blob:leased-asset');
    const revokeObjectURL = mock(() => {});
    URL.revokeObjectURL = revokeObjectURL;

    const cached = createCachedImageAssetResolver((assetId) => `/assets/${assetId}`);
    const resolved = await cached.resolver({
      assetId: 'leased-asset',
      signal: new AbortController().signal,
    });

    cached.dispose();
    expect(revokeObjectURL).not.toHaveBeenCalled();
    resolved.release?.();
    resolved.release?.();
    expect(revokeObjectURL).toHaveBeenCalledTimes(1);
  });
});

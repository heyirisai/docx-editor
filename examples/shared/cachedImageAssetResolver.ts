interface CachedImageAsset {
  controller: AbortController;
  promise: Promise<string>;
  src?: string;
  leases: number;
  revoked: boolean;
}

export interface CachedImageAssetResolver {
  resolver: (request: {
    assetId: string;
    width?: number;
    height?: number;
    signal: AbortSignal;
  }) => Promise<{ src: string; release?: () => void }>;
  dispose: () => void;
}

function abortError(): DOMException {
  return new DOMException('Image asset resolution was aborted', 'AbortError');
}

function waitForAsset(promise: Promise<string>, signal: AbortSignal): Promise<string> {
  if (signal.aborted) return Promise.reject(abortError());

  return new Promise((resolve, reject) => {
    const handleAbort = () => {
      reject(abortError());
    };
    signal.addEventListener('abort', handleAbort, { once: true });
    promise.then(
      (src) => {
        signal.removeEventListener('abort', handleAbort);
        if (signal.aborted) {
          reject(abortError());
        } else {
          resolve(src);
        }
      },
      (error: unknown) => {
        signal.removeEventListener('abort', handleAbort);
        reject(error);
      }
    );
  });
}

/**
 * Cache external image fetches and object URLs for the lifetime of one editor.
 *
 * Each asset has one shared fetch promise. Individual image elements receive a
 * reference-counted lease, while zero-lease entries remain warm so replacing
 * page DOM does not immediately fetch the same immutable asset again.
 */
export function createCachedImageAssetResolver(
  getAssetUrl: (assetId: string) => string
): CachedImageAssetResolver {
  const cache = new Map<string, CachedImageAsset>();
  let disposed = false;

  const revoke = (entry: CachedImageAsset): void => {
    if (!entry.src || entry.revoked) return;
    URL.revokeObjectURL(entry.src);
    entry.revoked = true;
  };

  const resolver: CachedImageAssetResolver['resolver'] = async ({ assetId, signal }) => {
    if (disposed || signal.aborted) throw abortError();

    let entry = cache.get(assetId);
    if (!entry) {
      const controller = new AbortController();
      const created: CachedImageAsset = {
        controller,
        promise: Promise.resolve(''),
        leases: 0,
        revoked: false,
      };
      created.promise = fetch(getAssetUrl(assetId), { signal: controller.signal })
        .then(async (response) => {
          if (!response.ok) {
            throw new Error(`Failed to load external image asset ${assetId}`);
          }
          const src = URL.createObjectURL(await response.blob());
          created.src = src;
          if (disposed) {
            revoke(created);
            throw abortError();
          }
          return src;
        })
        .catch((error: unknown) => {
          if (cache.get(assetId) === created) {
            cache.delete(assetId);
          }
          throw error;
        });
      entry = created;
      cache.set(assetId, entry);
    }

    const src = await waitForAsset(entry.promise, signal);
    if (disposed) throw abortError();

    entry.leases++;
    let released = false;
    const resolved = {
      src,
      release: () => {
        if (released) return;
        released = true;
        entry.leases = Math.max(0, entry.leases - 1);
        if (disposed && entry.leases === 0) {
          revoke(entry);
        }
      },
    };
    return resolved;
  };

  return {
    resolver,
    dispose: () => {
      if (disposed) return;
      disposed = true;
      for (const entry of cache.values()) {
        entry.controller.abort();
        if (entry.leases === 0) {
          revoke(entry);
        }
      }
      cache.clear();
    },
  };
}

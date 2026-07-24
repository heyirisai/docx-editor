/**
 * Viewport-gated image loading for externally stored DOCX media.
 *
 * The host owns fetching, deduplication, object-URL creation, and cache
 * lifetime. The painter owns visibility, bounded resolve/decode concurrency,
 * and releasing the host lease when rendered DOM is discarded.
 */

export interface ImageAssetRequest {
  assetId: string;
  width?: number;
  height?: number;
  signal: AbortSignal;
}

export interface ResolvedImageAsset {
  src: string;
  release?: () => void;
}

export type ImageAssetResolver = (request: ImageAssetRequest) => Promise<ResolvedImageAsset>;

export interface LazyImageAssetLoaderOptions {
  maxConcurrent?: number;
  rootMargin?: string;
}

interface ImageAssetEntry {
  image: HTMLImageElement;
  assetId: string;
  width?: number;
  height?: number;
  controller?: AbortController;
  release?: () => void;
  queued: boolean;
  loading: boolean;
  ready: boolean;
  released: boolean;
  settled: Promise<void>;
  settle: () => void;
  settledDone: boolean;
}

const DEFAULT_MAX_CONCURRENT = 4;
const DEFAULT_ROOT_MARGIN = '800px 0px 800px 0px';

/**
 * Coordinates all externally resolved images for one editor instance.
 */
export class LazyImageAssetLoader {
  private readonly resolver: ImageAssetResolver;
  private readonly maxConcurrent: number;
  private readonly observer?: IntersectionObserver;
  private readonly entries = new Map<HTMLImageElement, ImageAssetEntry>();
  private readonly queue: ImageAssetEntry[] = [];
  private readonly deferredReleases: Array<() => void> = [];
  private active = 0;
  private leaseHolds = 0;
  private disposed = false;

  constructor(resolver: ImageAssetResolver, options: LazyImageAssetLoaderOptions = {}) {
    this.resolver = resolver;
    this.maxConcurrent = Math.max(1, Math.floor(options.maxConcurrent ?? DEFAULT_MAX_CONCURRENT));

    if (typeof IntersectionObserver !== 'undefined') {
      this.observer = new IntersectionObserver(
        (entries) => {
          for (const intersection of entries) {
            if (!intersection.isIntersecting) continue;
            const image = intersection.target as HTMLImageElement;
            const entry = this.entries.get(image);
            if (!entry) continue;
            this.observer?.unobserve(image);
            this.enqueue(entry);
          }
        },
        {
          root: null,
          rootMargin: options.rootMargin ?? DEFAULT_ROOT_MARGIN,
        }
      );
    }
  }

  observe(image: HTMLImageElement, request: Omit<ImageAssetRequest, 'signal'>): void {
    this.releaseImage(image);
    if (this.disposed) return;

    const entry: ImageAssetEntry = {
      image,
      assetId: request.assetId,
      width: request.width,
      height: request.height,
      queued: false,
      loading: false,
      ready: false,
      released: false,
      settled: Promise.resolve(),
      settle: () => {},
      settledDone: true,
    };
    this.resetSettlement(entry);
    this.entries.set(image, entry);
    image.dataset.assetId = request.assetId;
    image.dataset.assetState = 'waiting';

    if (this.observer) {
      this.observer.observe(image);
    } else {
      this.enqueue(entry);
    }
  }

  releaseSubtree(root: ParentNode): void {
    if (root instanceof HTMLImageElement) {
      this.releaseImage(root);
    }
    for (const image of root.querySelectorAll<HTMLImageElement>('img[data-asset-id]')) {
      this.releaseImage(image);
    }
  }

  /**
   * Eagerly resolve and decode every observed external image in a subtree.
   * Used by print/DOM snapshot callers after virtual pages are materialized.
   */
  async resolveSubtree(root: ParentNode): Promise<void> {
    if (this.disposed) return;
    const images =
      root instanceof HTMLImageElement
        ? [root, ...root.querySelectorAll<HTMLImageElement>('img[data-asset-id]')]
        : [...root.querySelectorAll<HTMLImageElement>('img[data-asset-id]')];
    const pending: Promise<void>[] = [];
    for (const image of images) {
      const entry = this.entries.get(image);
      if (!entry) continue;
      this.observer?.unobserve(image);
      this.enqueue(entry);
      pending.push(entry.settled);
    }
    await Promise.all(pending);
  }

  /**
   * Keep resolved object-URL leases alive while another document (for
   * example a print window) consumes cloned image sources.
   */
  holdLeases(): () => void {
    if (this.disposed) return () => {};
    this.leaseHolds++;
    let released = false;
    return () => {
      if (released) return;
      released = true;
      this.leaseHolds = Math.max(0, this.leaseHolds - 1);
      if (this.leaseHolds !== 0) return;
      for (const release of this.deferredReleases.splice(0)) release();
    };
  }

  dispose(): void {
    if (this.disposed) return;
    this.disposed = true;
    this.observer?.disconnect();
    for (const image of [...this.entries.keys()]) {
      this.releaseImage(image);
    }
    this.queue.length = 0;
  }

  private enqueue(entry: ImageAssetEntry): void {
    if (entry.released || entry.ready || entry.queued || entry.loading) return;
    this.resetSettlement(entry);
    entry.queued = true;
    entry.image.dataset.assetState = 'queued';
    this.queue.push(entry);
    this.drain();
  }

  private drain(): void {
    while (!this.disposed && this.active < this.maxConcurrent && this.queue.length > 0) {
      const entry = this.queue.shift();
      if (!entry || entry.released) continue;
      entry.queued = false;
      entry.loading = true;
      this.active++;
      void this.resolveAndDecode(entry).finally(() => {
        entry.loading = false;
        this.active--;
        this.drain();
      });
    }
  }

  private async resolveAndDecode(entry: ImageAssetEntry): Promise<void> {
    const controller = new AbortController();
    entry.controller = controller;
    entry.ready = false;
    entry.image.dataset.assetState = 'loading';

    try {
      const resolved = await this.resolver({
        assetId: entry.assetId,
        width: entry.width,
        height: entry.height,
        signal: controller.signal,
      });

      if (entry.released || controller.signal.aborted || this.disposed) {
        resolved.release?.();
        return;
      }

      entry.release = resolved.release;
      entry.image.src = resolved.src;
      entry.image.dataset.assetState = 'decoding';

      if (typeof entry.image.decode === 'function') {
        try {
          await entry.image.decode();
        } catch {
          // A decode rejection can mean the browser decoded during assignment,
          // the DOM was removed, or the asset is invalid. Keep the assigned
          // source so normal image error/render behavior remains available.
        }
      }

      if (!entry.released) {
        entry.ready = true;
        entry.image.dataset.assetState = 'ready';
      }
    } catch (error) {
      if (!entry.released && !(error instanceof DOMException && error.name === 'AbortError')) {
        entry.image.dataset.assetState = 'error';
      }
    } finally {
      entry.controller = undefined;
      entry.settle();
    }
  }

  private resetSettlement(entry: ImageAssetEntry): void {
    if (!entry.settledDone) return;

    entry.settledDone = false;
    entry.settled = new Promise<void>((resolve) => {
      entry.settle = () => {
        if (entry.settledDone) return;
        entry.settledDone = true;
        resolve();
      };
    });
  }

  private releaseImage(image: HTMLImageElement): void {
    const entry = this.entries.get(image);
    if (!entry) return;

    entry.released = true;
    this.observer?.unobserve(image);
    entry.controller?.abort();
    if (entry.release) {
      if (this.leaseHolds > 0) {
        this.deferredReleases.push(entry.release);
      } else {
        entry.release();
      }
    }
    entry.settle();
    entry.release = undefined;
    image.removeAttribute('src');
    delete image.dataset.assetState;
    this.entries.delete(image);
  }
}

export interface ImageAssetSource {
  assetId?: string;
  src?: string;
  width?: number;
  height?: number;
}

/**
 * Assign a legacy inline source immediately, or register an external asset
 * without creating a network request until it nears the viewport.
 */
export function setImageAssetSource(
  image: HTMLImageElement,
  source: ImageAssetSource,
  loader?: LazyImageAssetLoader
): void {
  if (source.assetId) {
    image.dataset.assetId = source.assetId;
    if (loader) {
      loader.observe(image, {
        assetId: source.assetId,
        width: source.width,
        height: source.height,
      });
      return;
    }
  }

  if (source.src) {
    image.src = source.src;
  }
}

import {
  type ImageAssetResolver,
  LazyImageAssetLoader,
} from '@eigenpal/docx-editor-core/layout-painter';
import { useEffect, useMemo, useRef } from 'react';

export function useLazyImageAssetLoader(
  resolver: ImageAssetResolver | undefined
): LazyImageAssetLoader | undefined {
  const loader = useMemo(
    () => (resolver ? new LazyImageAssetLoader(resolver, { maxConcurrent: 4 }) : undefined),
    [resolver]
  );
  const effectGenerationRef = useRef(0);
  const currentLoaderRef = useRef(loader);
  currentLoaderRef.current = loader;

  useEffect(() => {
    const effectGeneration = ++effectGenerationRef.current;
    return () => {
      queueMicrotask(() => {
        if (
          effectGenerationRef.current === effectGeneration ||
          currentLoaderRef.current !== loader
        ) {
          loader?.dispose();
        }
      });
    };
  }, [loader]);

  return loader;
}

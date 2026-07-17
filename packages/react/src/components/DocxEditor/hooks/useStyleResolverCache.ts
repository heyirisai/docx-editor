import { useCallback, useRef } from 'react';
import { createStyleResolver } from '@eigenpal/docx-editor-core/prosemirror';

/**
 * Caches the style resolver by styles-object identity so it is not rebuilt
 * on every selection change. Extracted verbatim from DocxEditor.tsx.
 */
export function useStyleResolverCache(): (
  styles: Parameters<typeof createStyleResolver>[0]
) => ReturnType<typeof createStyleResolver> {
  const cacheRef = useRef<{
    styles: unknown;
    resolver: ReturnType<typeof createStyleResolver>;
  } | null>(null);

  return useCallback((styles: Parameters<typeof createStyleResolver>[0]) => {
    const cached = cacheRef.current;
    if (cached && cached.styles === styles) {
      return cached.resolver;
    }
    const resolver = createStyleResolver(styles);
    cacheRef.current = { styles, resolver };
    return resolver;
  }, []);
}

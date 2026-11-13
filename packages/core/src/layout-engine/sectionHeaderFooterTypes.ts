/**
 * Per-section header/footer types carried by the layout engine.
 * @packageDocumentation
 * @public
 */

/**
 * Header/footer part references (`r:id`s) of one section, by variant. Even-page
 * parts are not modelled: `w:evenAndOddHeaders` is unsupported.
 */
export type SectionHeaderFooterRefs = {
  headerDefault?: string;
  headerFirst?: string;
  footerDefault?: string;
  footerFirst?: string;
};

/**
 * Header/footer settings a section states in its own `w:sectPr`. Omitted
 * references inherit from the previous section (ECMA-376 §17.10.5); the
 * consumer resolves that chain.
 */
export type SectionHeaderFooterSettings = {
  refs?: SectionHeaderFooterRefs;
  /** `w:titlePg` — the section's first page uses the `first` stories. */
  titlePg?: boolean;
  /** `w:pgMar/@w:header` in px, when stated. */
  headerDistance?: number;
  /** `w:pgMar/@w:footer` in px, when stated. */
  footerDistance?: number;
};

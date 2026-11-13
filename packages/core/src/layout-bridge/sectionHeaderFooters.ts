/**
 * Per-section header/footer resolution (ECMA-376 §17.10).
 *
 * Every section names its own header/footer parts. A page shows the stories
 * of the section it belongs to (`Page.sectionIndex`): the `first` pair on the
 * section's first page when the section sets `w:titlePg`, the `default` pair
 * everywhere else. A section that omits a reference of some type inherits
 * that reference from the previous section; the first section has nothing to
 * inherit, so a variant it does not reference is simply empty. A `first`
 * reference is never shown without `w:titlePg`. `even` references are ignored
 * (`w:evenAndOddHeaders` is unsupported).
 *
 * Section settings travel on the `sectionBreak` flow blocks (so they follow
 * edits to the PM doc) plus the body-level `w:sectPr` for the last section.
 */

import type {
  FlowBlock,
  Page,
  SectionBreakBlock,
  SectionHeaderFooterRefs,
  SectionHeaderFooterSettings,
} from '../layout-engine/types';
import type { Node as PMNode } from 'prosemirror-model';
import type { DocxPackage, HeaderFooter, SectionProperties } from '../types/document';
import { twipsToPixels } from '../utils/units';

/** Header/footer settings a `w:sectPr` states itself (no inheritance). */
export function sectionHeaderFooterSettings(
  sp: SectionProperties | null | undefined
): SectionHeaderFooterSettings | undefined {
  if (!sp) return undefined;
  const refs: SectionHeaderFooterRefs = {};
  for (const ref of sp.headerReferences ?? []) {
    if (!ref.rId) continue;
    if (ref.type === 'default') refs.headerDefault = ref.rId;
    else if (ref.type === 'first') refs.headerFirst = ref.rId;
  }
  for (const ref of sp.footerReferences ?? []) {
    if (!ref.rId) continue;
    if (ref.type === 'default') refs.footerDefault = ref.rId;
    else if (ref.type === 'first') refs.footerFirst = ref.rId;
  }
  const settings: SectionHeaderFooterSettings = {};
  if (Object.keys(refs).length > 0) settings.refs = refs;
  if (sp.titlePg !== undefined) settings.titlePg = sp.titlePg;
  if (sp.headerDistance != null) settings.headerDistance = twipsToPixels(sp.headerDistance);
  if (sp.footerDistance != null) settings.footerDistance = twipsToPixels(sp.footerDistance);
  return settings;
}

/** One section's header/footer stories, after inheritance. */
export interface ResolvedSectionHeaderFooter {
  /** Effective part references (own, else inherited). */
  refs: SectionHeaderFooterRefs;
  header: HeaderFooter | null;
  footer: HeaderFooter | null;
  firstHeader: HeaderFooter | null;
  firstFooter: HeaderFooter | null;
  titlePg: boolean;
  headerDistance?: number;
  footerDistance?: number;
}

const REF_KEYS: Array<keyof SectionHeaderFooterRefs> = [
  'headerDefault',
  'headerFirst',
  'footerDefault',
  'footerFirst',
];

/**
 * Resolve the header/footer stories of every section, in order: one entry per
 * `sectionBreak` block, then the final (body-level) section. The index of an
 * entry is the `Page.sectionIndex` of the pages that show it.
 */
export function resolveSectionHeaderFooters(
  blocks: readonly FlowBlock[],
  finalSectionProperties: SectionProperties | null | undefined,
  headers: Map<string, HeaderFooter> | undefined,
  footers: Map<string, HeaderFooter> | undefined
): ResolvedSectionHeaderFooter[] {
  const own: Array<SectionHeaderFooterSettings | undefined> = [];
  for (const block of blocks) {
    if (block.kind === 'sectionBreak') own.push((block as SectionBreakBlock).headerFooter);
  }
  own.push(sectionHeaderFooterSettings(finalSectionProperties));

  const out: ResolvedSectionHeaderFooter[] = [];
  let prior: ResolvedSectionHeaderFooter | undefined;
  for (const settings of own) {
    const refs: SectionHeaderFooterRefs = {};
    for (const key of REF_KEYS) {
      const rId = settings?.refs?.[key] ?? prior?.refs[key];
      if (rId) refs[key] = rId;
    }
    const lookup = (map: Map<string, HeaderFooter> | undefined, rId: string | undefined) =>
      (rId && map?.get(rId)) || null;
    const resolved: ResolvedSectionHeaderFooter = {
      refs,
      header: lookup(headers, refs.headerDefault),
      footer: lookup(footers, refs.footerDefault),
      firstHeader: lookup(headers, refs.headerFirst),
      firstFooter: lookup(footers, refs.footerFirst),
      // Inherited like the references, matching `applySectionInheritance`.
      titlePg: settings?.titlePg ?? prior?.titlePg ?? false,
      headerDistance: settings?.headerDistance,
      footerDistance: settings?.footerDistance,
    };
    out.push(resolved);
    prior = resolved;
  }
  return out;
}

/**
 * Whether `page` shows its section's first-page stories. Pages laid out
 * before `firstPageOfSection` existed fall back to "page 1".
 */
export function pageUsesFirstPageHeaderFooter(
  page: Readonly<Pick<Page, 'number' | 'firstPageOfSection'>>,
  titlePg: boolean | undefined
): boolean {
  if (titlePg !== true) return false;
  return page.firstPageOfSection ?? page.number === 1;
}

/**
 * Record on every page the header/footer parts it shows
 * (`Page.headerFooterRefs`: only the variant in play — `headerFirst` /
 * `footerFirst` on a `titlePg` first page, `headerDefault` / `footerDefault`
 * elsewhere), so hosts can tell which part a click on a page's header or
 * footer should edit.
 */
export function assignPageHeaderFooterRefs(
  pages: Page[],
  sections: readonly ResolvedSectionHeaderFooter[]
): void {
  if (sections.length === 0) return;
  for (const page of pages) {
    const section = sections[Math.min(page.sectionIndex ?? 0, sections.length - 1)];
    const refs: SectionHeaderFooterRefs = {};
    if (pageUsesFirstPageHeaderFooter(page, section.titlePg)) {
      if (section.firstHeader && section.refs.headerFirst)
        refs.headerFirst = section.refs.headerFirst;
      if (section.firstFooter && section.refs.footerFirst)
        refs.footerFirst = section.refs.footerFirst;
    } else {
      if (section.header && section.refs.headerDefault)
        refs.headerDefault = section.refs.headerDefault;
      if (section.footer && section.refs.footerDefault)
        refs.footerDefault = section.refs.footerDefault;
    }
    page.headerFooterRefs = refs;
  }
}

/** A header/footer part a page shows. */
export interface HeaderFooterPartOnPage {
  rId: string;
  /** Shown as the section's `titlePg` first-page variant. */
  isFirstPage: boolean;
}

/** The section a clicked page belongs to (`Page.sectionIndex`). */
export interface HeaderFooterPageSection {
  index: number;
  /** The page is its section's first page (shows `first` parts under `w:titlePg`). */
  firstPageOfSection: boolean;
}

/**
 * Host callback for a double-click on a page's header/footer: the slot, the
 * page number, the part that page shows (null when it shows none), and the
 * page's section (to create a part there when it shows none).
 */
export type HeaderFooterDoubleClickHandler = (
  position: 'header' | 'footer',
  pageNumber?: number,
  shownPart?: HeaderFooterPartOnPage | null,
  section?: HeaderFooterPageSection
) => void;

/**
 * The header or footer part (`r:id`) `page` shows, if any — see
 * {@link assignPageHeaderFooterRefs}.
 */
export function pageHeaderFooterRId(
  page: Readonly<Pick<Page, 'headerFooterRefs'>> | null | undefined,
  position: 'header' | 'footer'
): HeaderFooterPartOnPage | null {
  const refs = page?.headerFooterRefs;
  if (!refs) return null;
  const first = position === 'header' ? refs.headerFirst : refs.footerFirst;
  if (first) return { rId: first, isFirstPage: true };
  const def = position === 'header' ? refs.headerDefault : refs.footerDefault;
  return def ? { rId: def, isFirstPage: false } : null;
}

/**
 * The section a header/footer double-click edits, for a page that shows no
 * part (see {@link HeaderFooterDoubleClickHandler}).
 */
export interface HeaderFooterSectionTarget {
  /** Position of the section-break paragraph holding the sectPr; null for the body-level one. */
  pos: number | null;
  sectionProperties: SectionProperties | undefined;
  hdrFtrType: 'default' | 'first';
  isFirstPage: boolean;
  /** The part this section uses for that variant after inheritance, if any. */
  inheritedRId?: string;
}

/**
 * Resolve the clicked page's own section (ECMA-376 §17.10): its `w:sectPr`
 * is a `_sectionProperties` paragraph attr in the body PM doc, or the
 * body-level one for the last section (also the fallback without page info).
 * References and `w:titlePg` inherit from earlier sections.
 */
export function resolveHeaderFooterSectionTarget(
  pmDoc: PMNode | null | undefined,
  finalSectionProperties: SectionProperties | null | undefined,
  position: 'header' | 'footer',
  section?: HeaderFooterPageSection,
  pageNumber?: number
): HeaderFooterSectionTarget {
  const sections: { sp: SectionProperties | undefined; pos: number | null }[] = [];
  pmDoc?.descendants((node, pos) => {
    if (node.type.name === 'table') return false;
    const sp = node.attrs._sectionProperties as SectionProperties | null | undefined;
    if (node.type.name === 'paragraph' && sp) sections.push({ sp, pos });
    return true;
  });
  sections.push({ sp: finalSectionProperties ?? undefined, pos: null });
  const index = Math.min(section?.index ?? sections.length - 1, sections.length - 1);

  const refKey = position === 'header' ? 'headerReferences' : 'footerReferences';
  let titlePg = false;
  for (let i = 0; i <= index; i++) titlePg = sections[i]!.sp?.titlePg ?? titlePg;
  const isFirstPage = titlePg && (section ? section.firstPageOfSection : (pageNumber ?? 1) === 1);
  const hdrFtrType = isFirstPage ? 'first' : 'default';
  let inheritedRId: string | undefined;
  for (let i = 0; i <= index; i++) {
    const ref = sections[i]!.sp?.[refKey]?.find((r) => r.type === hdrFtrType);
    if (ref?.rId) inheritedRId = ref.rId;
  }
  const { pos, sp } = sections[index]!;
  return { pos, sectionProperties: sp, hdrFtrType, isFirstPage, inheritedRId };
}

/**
 * A new empty header/footer part for `pkg`: its `r:id`, the part map and the
 * relationships with it registered (so the serializer wires up content types
 * and document rels), plus `sectionProperties` with a reference to it.
 */
export function createHeaderFooterPart(
  pkg: DocxPackage,
  sectionProperties: SectionProperties,
  position: 'header' | 'footer',
  hdrFtrType: 'default' | 'first'
): {
  rId: string;
  part: HeaderFooter;
  map: Map<string, HeaderFooter>;
  relationships: NonNullable<DocxPackage['relationships']>;
  sectionProperties: SectionProperties;
} {
  const mapKey = position === 'header' ? 'headers' : 'footers';
  const existingRels = pkg.relationships;
  let rId = `rId_new_${position}_${hdrFtrType}`;
  for (let n = 2; existingRels?.has(rId) || pkg[mapKey]?.has(rId); n++) {
    rId = `rId_new_${position}_${hdrFtrType}_${n}`;
  }
  const part: HeaderFooter = {
    type: position,
    hdrFtrType,
    content: [{ type: 'paragraph', content: [] }],
  };
  const map = new Map(pkg[mapKey] ?? []);
  map.set(rId, part);

  const usedTargets = new Set<string>();
  for (const rel of existingRels?.values() ?? []) {
    if (rel.target) usedTargets.add(rel.target);
  }
  let targetNum = 1;
  while (usedTargets.has(`${position}${targetNum}.xml`)) targetNum++;
  const relationships = new Map(existingRels);
  relationships.set(rId, {
    id: rId,
    type: `http://schemas.openxmlformats.org/officeDocument/2006/relationships/${position}`,
    target: `${position}${targetNum}.xml`,
  });

  const refKey = position === 'header' ? 'headerReferences' : 'footerReferences';
  return {
    rId,
    part,
    map,
    relationships,
    sectionProperties: {
      ...sectionProperties,
      [refKey]: [...(sectionProperties[refKey] ?? []), { type: hdrFtrType, rId }],
    },
  };
}

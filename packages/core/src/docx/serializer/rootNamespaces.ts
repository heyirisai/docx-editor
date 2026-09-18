/**
 * Root namespace declarations, derived from one table.
 *
 * Each part Word writes declares its own ordered subset of prefixes, so these
 * are prefix LISTS — the URIs all come from `OOXML_NAMESPACE_URIS` in
 * `xmlParser`, which is also what re-declares a preserved fragment's inherited
 * prefixes. Several hand-maintained prefix→URI maps is how a root ends up
 * declaring a different URI than the fragment it has to bind, and
 * `namespace-tables.test.ts` keeps the relation enforced rather than commented.
 */

import { OOXML_NAMESPACE_URIS } from '../xmlParser';
import { escapeXml } from './xmlUtils';

/**
 * `document.xml` and `header*.xml`/`footer*.xml`. Both capture their source
 * root's declarations on parse, so this only has to cover what a rebuilt part
 * emits from scratch — plus the modern Word extension prefixes, because
 * verbatim passthrough (a captured `w:sdtPr`) can echo `w16*`/`wne` children
 * and an undeclared prefix makes Word offer to repair the file.
 */
export const CAPTURING_ROOT_PREFIXES = [
  'wpc',
  'mc',
  'o',
  'r',
  'm',
  'v',
  'wp14',
  'wp',
  'w10',
  'w',
  'w14',
  'w15',
  'w16se',
  'w16cid',
  'w16',
  'w16cex',
  'w16sdtdh',
  'wne',
  'wpg',
  'wps',
];

/**
 * `comments.xml` and friends, `footnotes.xml`, `endnotes.xml` — parts with no
 * per-document capture, so their root has to declare everything up front.
 */
export const FIXED_ROOT_PREFIXES = [
  'wpc',
  'cx',
  'cx1',
  'cx2',
  'cx3',
  'cx4',
  'cx5',
  'cx6',
  'cx7',
  'cx8',
  'mc',
  'aink',
  'am3d',
  'o',
  'oel',
  'r',
  'm',
  'v',
  'wp14',
  'wp',
  'w10',
  'w',
  'w14',
  'w15',
  'w16cex',
  'w16cid',
  'w16',
  'w16du',
  'w16sdtdh',
  'w16sdtfl',
  'w16se',
  'wpg',
  'wpi',
  'wne',
  'wps',
];

/**
 * `prefixes`, plus any the source root declared that the list lacks — a prefix
 * preserved markup inherited but the root never declares exports as invalid XML.
 * Returns the declaration string and the prefixes it actually emitted, which is
 * what `mc:Ignorable` has to be intersected with.
 */
export function buildRootNamespaces(
  prefixes: string[],
  captured?: Record<string, string>
): { decl: string; prefixes: string[] } {
  const merged: Record<string, string> = {};
  for (const prefix of prefixes) merged[prefix] = OOXML_NAMESPACE_URIS[prefix];
  for (const [prefix, uri] of Object.entries(captured ?? {})) {
    if (!(prefix in merged)) merged[prefix] = uri;
  }
  return {
    decl: Object.entries(merged)
      .map(([prefix, uri]) => `xmlns:${escapeXml(prefix)}="${escapeXml(uri)}"`)
      .join(' '),
    prefixes: Object.keys(merged),
  };
}

/**
 * Full OOXML namespace block matching Word output.
 *
 * Shared by the parts that emit a standalone WordprocessingML root with the
 * complete namespace set — `comments.xml` (and its companions) and
 * `footnotes.xml` / `endnotes.xml`.
 */
export const OOXML_NAMESPACES = buildRootNamespaces(FIXED_ROOT_PREFIXES).decl;

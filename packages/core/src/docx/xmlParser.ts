/**
 * XML Parser Utilities for OOXML
 *
 * Provides helper functions for parsing Office Open XML (OOXML) content
 * with proper namespace handling.
 *
 * OOXML uses many namespaces:
 * - w:  WordprocessingML (main document content)
 * - a:  DrawingML (graphics)
 * - r:  Relationships
 * - wp: Word Drawing positioning
 * - wps: Word Drawing shapes
 * - wpc: Word Drawing canvas
 * - wpg: Word Drawing group
 * - m:  Math
 * - mc: Markup Compatibility
 * - v:  VML (legacy vector graphics)
 * - o:  Office (extensions)
 * - pic: Pictures
 */

import { xml2js, js2xml, type Element as XmlElement } from 'xml-js';
import { MAX_ELEMENT_DEPTH, scanXmlFragment } from './xmlFragmentScanner';

// Re-export Element type for consumers
export type { Element as XmlElement } from 'xml-js';

/**
 * Common OOXML namespace URIs
 */
export const NAMESPACES = {
  // Main namespaces
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',

  // Drawing namespaces
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  wp14: 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
  wps: 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
  wpc: 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
  wpg: 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',

  // Picture namespace
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',

  // Math namespace
  m: 'http://schemas.openxmlformats.org/officeDocument/2006/math',

  // Markup Compatibility
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',

  // Legacy VML
  v: 'urn:schemas-microsoft-com:vml',
  o: 'urn:schemas-microsoft-com:office:office',

  // Other
  w14: 'http://schemas.microsoft.com/office/word/2010/wordml',
  w15: 'http://schemas.microsoft.com/office/word/2012/wordml',

  // Content Types
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',

  // Relationships
  pr: 'http://schemas.openxmlformats.org/package/2006/relationships',
} as const;

// Matches a literal `&` that isn't the start of a valid XML entity reference
// (named, decimal, or hex). Real-world DOCX files produced by non-Word tools
// occasionally leave stray ampersands unescaped; xml-js's sax parser is strict
// and rejects the whole document with "Invalid character in entity name".
const STRAY_AMPERSAND_RE = /&(?!(?:[A-Za-z][A-Za-z0-9]*|#[0-9]+|#x[0-9a-fA-F]+);)/g;

function escapeStrayAmpersands(xml: string): string {
  return xml.replace(STRAY_AMPERSAND_RE, '&amp;');
}

/**
 * Parse XML string into element tree
 *
 * @param xml - XML string to parse
 * @returns Parsed element tree
 */
export function parseXml(xml: string): XmlElement {
  const sanitized = escapeStrayAmpersands(xml);

  try {
    return xml2js(sanitized, {
      compact: false,
      ignoreComment: true,
      ignoreInstruction: true,
      ignoreDoctype: true,
      alwaysArray: false,
      // IMPORTANT: Do NOT trim whitespace - it strips significant spaces
      // around hyperlinks and other inline elements. DOCX uses xml:space="preserve"
      // to indicate significant whitespace, but we need to preserve all text as-is.
      trim: false,
      // IMPORTANT: Without this, xml-js silently drops whitespace-only text nodes
      // (e.g. <w:t xml:space="preserve"> </w:t> loses the space).
      captureSpacesBetweenElements: true,
      attributesKey: 'attributes',
      textKey: 'text',
    }) as XmlElement;
  } catch (error) {
    if (!(error instanceof Error)) throw error;

    // xml-js error messages include "Line: N\nColumn: M\nChar: c" but no
    // surrounding context. Append a short snippet of the input around the
    // offending column so the next reader can see what byte broke the parse.
    const colMatch = error.message.match(/Column:\s*(\d+)/);
    if (colMatch) {
      const col = parseInt(colMatch[1], 10);
      const start = Math.max(0, col - 30);
      const snippet = JSON.stringify(sanitized.slice(start, col + 30));
      const wrapped = new Error(`${error.message}\nNear: ${snippet}`);
      wrapped.stack = error.stack;
      throw wrapped;
    }
    throw error;
  }
}

/**
 * Serialize an XmlElement back to an XML string
 */
export function elementToXml(element: XmlElement): string {
  return js2xml({ elements: [element] }, { compact: false, spaces: 0 });
}

/**
 * The `xmlns:` bindings declared on an element, so a rebuilt part can re-declare
 * whatever its preserved markup inherited from the original root.
 */
export function rootNamespaceDeclarations(element: XmlElement): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [key, value] of Object.entries(element.attributes ?? {})) {
    if (key.startsWith('xmlns:')) out[key.slice(6)] = String(value);
  }
  return out;
}

/**
 * The prefixes the root listed in `mc:Ignorable` — the other half of the
 * markup-compatibility contract. Declaring a namespace only says what a prefix
 * means; `mc:Ignorable` is what lets a consumer skip an element in it. A
 * captured prefix that never makes the list leaves preserved markup a consumer
 * has to reject rather than ignore.
 */
export function rootIgnorablePrefixes(element: XmlElement): string[] {
  for (const [key, value] of Object.entries(element.attributes ?? {})) {
    if (getLocalName(key) !== 'Ignorable') continue;
    return String(value).split(/\s+/).filter(Boolean);
  }
  return [];
}

/**
 * Prefix -> URI for every namespace Word writes. THE table: every root the
 * serializers emit is a derived view of this one (see
 * `serializer/rootNamespaces.ts`), and `namespace-tables.test.ts` asserts the
 * superset relation rather than leaving it to a comment.
 *
 * A captured fragment is re-declared wherever it lands, and
 * `footnotes.xml`/`endnotes.xml`/`comments.xml` use a fixed block with no
 * per-document capture — so a prefix missing here (`a16:creationId` appears in
 * nearly every Word drawing) exports unbound and voids the part.
 */
export const OOXML_NAMESPACE_URIS: Record<string, string> = {
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  a14: 'http://schemas.microsoft.com/office/drawing/2010/main',
  a16: 'http://schemas.microsoft.com/office/drawing/2014/main',
  adec: 'http://schemas.microsoft.com/office/drawing/2017/decorative',
  aink: 'http://schemas.microsoft.com/office/drawing/2016/ink',
  am3d: 'http://schemas.microsoft.com/office/drawing/2017/model3d',
  asvg: 'http://schemas.microsoft.com/office/drawing/2016/SVG/main',
  c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
  cx: 'http://schemas.microsoft.com/office/drawing/2014/chartex',
  cx1: 'http://schemas.microsoft.com/office/drawing/2015/9/8/chartex',
  cx2: 'http://schemas.microsoft.com/office/drawing/2015/10/21/chartex',
  cx3: 'http://schemas.microsoft.com/office/drawing/2016/5/9/chartex',
  cx4: 'http://schemas.microsoft.com/office/drawing/2016/5/10/chartex',
  cx5: 'http://schemas.microsoft.com/office/drawing/2016/5/11/chartex',
  cx6: 'http://schemas.microsoft.com/office/drawing/2016/5/12/chartex',
  cx7: 'http://schemas.microsoft.com/office/drawing/2016/5/13/chartex',
  cx8: 'http://schemas.microsoft.com/office/drawing/2016/5/14/chartex',
  dgm: 'http://schemas.openxmlformats.org/drawingml/2006/diagram',
  m: 'http://schemas.openxmlformats.org/officeDocument/2006/math',
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  o: 'urn:schemas-microsoft-com:office:office',
  oel: 'http://schemas.microsoft.com/office/2019/extlst',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  v: 'urn:schemas-microsoft-com:vml',
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  w10: 'urn:schemas-microsoft-com:office:word',
  w14: 'http://schemas.microsoft.com/office/word/2010/wordml',
  w15: 'http://schemas.microsoft.com/office/word/2012/wordml',
  w16: 'http://schemas.microsoft.com/office/word/2018/wordml',
  w16cex: 'http://schemas.microsoft.com/office/word/2018/wordml/cex',
  w16cid: 'http://schemas.microsoft.com/office/word/2016/wordml/cid',
  w16du: 'http://schemas.microsoft.com/office/word/2023/wordml/word16du',
  w16sdtdh: 'http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash',
  w16sdtfl: 'http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock',
  w16se: 'http://schemas.microsoft.com/office/word/2015/wordml/symex',
  wne: 'http://schemas.microsoft.com/office/word/2006/wordml',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  wp14: 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
  wpc: 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
  wpg: 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
  wpi: 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk',
  wps: 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
};

function prefixOf(name: string): string | null {
  const i = name.indexOf(':');
  return i > 0 ? name.slice(0, i) : null;
}

/**
 * Collect prefixes used without a binding in scope. `inScope` holds only what
 * ancestors WITHIN the fragment declared, never what the original root did.
 * `state.truncated` reports that the depth bound cut the walk short — the
 * caller has an INCOMPLETE prefix set and must not treat it as the whole
 * picture.
 */
function collectUnbound(
  el: XmlElement,
  inScope: Set<string>,
  unbound: Set<string>,
  state: { truncated: boolean },
  depth = 0
): void {
  // Both callers walk attacker-controlled XML (a preserved part on parse, a
  // pasted fragment at the trust boundary), so the walk must be bounded.
  if (depth >= MAX_ELEMENT_DEPTH) {
    state.truncated = true;
    return;
  }
  let scope = inScope;
  for (const key of Object.keys(el.attributes ?? {})) {
    if (!key.startsWith('xmlns:')) continue;
    if (scope === inScope) scope = new Set(inScope);
    scope.add(key.slice(6));
  }

  const own = el.name ? prefixOf(el.name) : null;
  if (own && !scope.has(own)) unbound.add(own);

  for (const key of Object.keys(el.attributes ?? {})) {
    if (key.startsWith('xmlns:') || key === 'xmlns') continue;
    const p = prefixOf(key);
    // `xml:` is bound by the spec and never declared.
    if (p && p !== 'xml' && !scope.has(p)) unbound.add(p);
  }

  for (const child of el.elements ?? []) collectUnbound(child, scope, unbound, state, depth + 1);
}

/**
 * Elements a PASTED fragment may carry. A subset of what `w:r` allows
 * (ECMA-376 §17.3.2): block-level markup (`w:p`, `w:tbl`, `w:sectPr`) would be
 * emitted inside a run and make the file unopenable, and the run-level
 * constructs that CARRY BEHAVIOUR are excluded too —
 * `w:instrText`/`w:delInstrText`/`w:fldChar` (a field instruction Word acts on,
 * e.g. `DDEAUTO`, `INCLUDETEXT`) and `w:object` (an OLE embed). Those still
 * round-trip fine when the PARSER produces them; nothing from the clipboard
 * gets to introduce one.
 */
const RUN_CONTENT_ELEMENTS = new Set([
  'mc:AlternateContent',
  'w:annotationRef',
  'w:br',
  'w:commentReference',
  'w:continuationSeparator',
  'w:cr',
  'w:delText',
  'w:drawing',
  'w:endnoteRef',
  'w:endnoteReference',
  'w:footnoteRef',
  'w:footnoteReference',
  'w:lastRenderedPageBreak',
  'w:noBreakHyphen',
  'w:pict',
  'w:ptab',
  'w:ruby',
  'w:separator',
  'w:softHyphen',
  'w:sym',
  'w:t',
  'w:tab',
]);

const WELL_FORMED_CACHE_MAX = 512;
const wellFormedCache = new Map<string, boolean>();

/**
 * Whether `xml` is a single well-formed element using only known prefixes.
 * Preserved markup is written into the saved package verbatim, so anything
 * failing this is dropped rather than left to corrupt the part — Word refuses
 * to open a document.xml with unbalanced tags or an unbound prefix.
 *
 * {@link scanXmlFragment} IS the guarantee and runs in every runtime;
 * `DOMParser` runs on top of it wherever the host has one.
 *
 * `requireBoundPrefixes` additionally rejects a prefix that neither the table
 * nor the fragment itself binds, and `runContentOnly` rejects a root element
 * that cannot legally sit inside a `w:r`. Parsed-from-source markup relies on
 * the destination root (`rootNamespaces`) for prefixes and is already run
 * content by construction, so only the paste boundary sets either.
 */
export function isWellFormedXmlElement(
  xml: string,
  options?: { requireBoundPrefixes?: boolean; runContentOnly?: boolean }
): boolean {
  if (typeof xml !== 'string' || xml.trim().length === 0) return false;
  if (/]]>/.test(xml)) return false;

  // A save re-validates every preserved node, and a TOC alone produces hundreds
  // of repeats of the same short string. Fixed-length flag prefix so the key
  // stays unambiguous whatever the markup contains.
  const key = `${options?.requireBoundPrefixes ? '1' : '0'}${options?.runContentOnly ? '1' : '0'}${xml}`;
  const cached = wellFormedCache.get(key);
  if (cached !== undefined) return cached;

  const verdict = checkWellFormed(xml, options);
  if (wellFormedCache.size >= WELL_FORMED_CACHE_MAX) wellFormedCache.clear();
  wellFormedCache.set(key, verdict);
  return verdict;
}

function checkWellFormed(
  xml: string,
  options?: { requireBoundPrefixes?: boolean; runContentOnly?: boolean }
): boolean {
  const scan = scanXmlFragment(xml);
  if (!scan.ok) return false;
  if (options?.runContentOnly && !RUN_CONTENT_ELEMENTS.has(scan.rootName ?? '')) return false;

  // Whether an exotic prefix resolves is the destination root's business (it
  // may be in the captured `rootNamespaces`), so only the paste boundary — the
  // one path with no such root behind it — insists on a known binding.
  const bindings = new Map(Object.entries(OOXML_NAMESPACE_URIS));
  for (const prefix of scan.unbound) {
    if (bindings.has(prefix)) continue;
    if (options?.requireBoundPrefixes) return false;
    bindings.set(prefix, `urn:ep-unbound:${prefix}`);
  }

  const Parser = (globalThis as { DOMParser?: typeof DOMParser }).DOMParser;
  if (!Parser) return true;
  const declarations = [...bindings].map(([prefix, uri]) => `xmlns:${prefix}="${uri}"`).join(' ');
  try {
    const doc = new Parser().parseFromString(
      `<epRoot ${declarations}>${xml}</epRoot>`,
      'application/xml'
    );
    if (doc.getElementsByTagName('parsererror').length > 0) return false;
    return (doc.documentElement?.children?.length ?? 0) === 1;
  } catch {
    return false;
  }
}

/**
 * Serialize an element for re-insertion elsewhere, declaring any prefix it uses
 * but inherits from its original root — otherwise the part is invalid XML.
 */
export function elementToSelfContainedXml(element: XmlElement): string {
  const unbound = new Set<string>();
  const state = { truncated: false };
  collectUnbound(element, new Set(), unbound, state);

  // Past the depth bound the prefix set is incomplete, so declaring only what
  // was found would leave a deeper one unbound and void the part. Declare the
  // whole table instead; markup that deep using a prefix outside it is rejected
  // by `isWellFormedXmlElement` before it can reach the package.
  const missing = state.truncated
    ? Object.keys(OOXML_NAMESPACE_URIS)
    : [...unbound].filter((p) => p in OOXML_NAMESPACE_URIS);
  if (missing.length === 0) return elementToXml(element);

  const withDecls: XmlElement = {
    ...element,
    attributes: {
      ...(element.attributes ?? {}),
      ...Object.fromEntries(missing.map((p) => [`xmlns:${p}`, OOXML_NAMESPACE_URIS[p]])),
    },
  };
  return elementToXml(withDecls);
}

/**
 * Parse XML string to a more convenient format
 */
export function parseXmlDocument(xml: string): XmlElement | null {
  try {
    const parsed = parseXml(xml);

    // The root is typically the declaration + elements array
    if (parsed.elements && parsed.elements.length > 0) {
      // Return the first real element (skip declarations)
      return parsed.elements.find((e) => e.type === 'element') ?? null;
    }

    return parsed;
  } catch (error) {
    console.warn('Failed to parse XML:', error);
    return null;
  }
}

/**
 * Get local name from a prefixed element name
 * e.g., "w:p" -> "p", "a:graphic" -> "graphic"
 */
export function getLocalName(name: string): string {
  const colonIndex = name.indexOf(':');
  return colonIndex >= 0 ? name.substring(colonIndex + 1) : name;
}

/**
 * Get namespace prefix from an element name
 * e.g., "w:p" -> "w", "a:graphic" -> "a"
 */
export function getNamespacePrefix(name: string): string | null {
  const colonIndex = name.indexOf(':');
  return colonIndex >= 0 ? name.substring(0, colonIndex) : null;
}

/**
 * Check if an element matches a given namespaced name
 *
 * @param element - Element to check
 * @param namespace - Namespace prefix (e.g., "w", "a")
 * @param localName - Local element name (e.g., "p", "r")
 */
export function matchesName(element: XmlElement, namespace: string, localName: string): boolean {
  if (!element.name) return false;

  const fullName = `${namespace}:${localName}`;
  if (element.name === fullName) return true;

  // Also check just the local name if no namespace prefix in element
  if (getLocalName(element.name) === localName) return true;

  return false;
}

/**
 * Find first child element matching the given namespaced name
 *
 * @param parent - Parent element
 * @param namespace - Namespace prefix (e.g., "w")
 * @param localName - Local element name (e.g., "p")
 * @returns First matching child or null
 */
export function findChild(
  parent: XmlElement | null | undefined,
  namespace: string,
  localName: string
): XmlElement | null {
  if (!parent || !parent.elements) return null;

  const fullName = `${namespace}:${localName}`;

  for (const child of parent.elements) {
    if (child.type !== 'element') continue;

    if (child.name === fullName) {
      return child;
    }

    // Check local name match
    if (getLocalName(child.name || '') === localName) {
      return child;
    }
  }

  return null;
}

/**
 * Find all child elements matching the given namespaced name
 *
 * @param parent - Parent element
 * @param namespace - Namespace prefix
 * @param localName - Local element name
 * @returns Array of matching children
 */
export function findChildren(
  parent: XmlElement | null | undefined,
  namespace: string,
  localName: string
): XmlElement[] {
  if (!parent || !parent.elements) return [];

  const fullName = `${namespace}:${localName}`;
  const results: XmlElement[] = [];

  for (const child of parent.elements) {
    if (child.type !== 'element') continue;

    if (child.name === fullName || getLocalName(child.name || '') === localName) {
      results.push(child);
    }
  }

  return results;
}

/**
 * Find first child element by local name only (ignoring namespace)
 *
 * @param parent - Parent element
 * @param localName - Local element name
 * @returns First matching child or null
 */
export function findChildByLocalName(
  parent: XmlElement | null | undefined,
  localName: string
): XmlElement | null {
  if (!parent || !parent.elements) return null;

  for (const child of parent.elements) {
    if (child.type !== 'element') continue;

    if (getLocalName(child.name || '') === localName) {
      return child;
    }
  }

  return null;
}

/**
 * Find all child elements by local name only
 *
 * @param parent - Parent element
 * @param localName - Local element name
 * @returns Array of matching children
 */
export function findChildrenByLocalName(
  parent: XmlElement | null | undefined,
  localName: string
): XmlElement[] {
  if (!parent || !parent.elements) return [];

  return parent.elements.filter(
    (child) => child.type === 'element' && getLocalName(child.name || '') === localName
  );
}

/**
 * Find first child element by full name (including namespace prefix)
 *
 * @param parent - Parent element
 * @param fullName - Full element name with namespace prefix (e.g., 'wp:extent')
 * @returns First matching child or null
 */
export function findByFullName(
  parent: XmlElement | null | undefined,
  fullName: string
): XmlElement | null {
  if (!parent || !parent.elements) return null;

  for (const child of parent.elements) {
    if (child.type !== 'element') continue;
    if (child.name === fullName) return child;
  }

  return null;
}

/**
 * Get all child elements (excludes text nodes, etc.)
 *
 * @param parent - Parent element
 * @returns Array of child elements
 */
export function getChildElements(parent: XmlElement | null | undefined): XmlElement[] {
  if (!parent || !parent.elements) return [];
  return parent.elements.filter((child) => child.type === 'element');
}

/**
 * Get an attribute value from an element
 *
 * @param element - Element to get attribute from
 * @param namespace - Namespace prefix for the attribute (or null for no namespace)
 * @param name - Attribute name
 * @returns Attribute value or null if not found
 */
export function getAttribute(
  element: XmlElement | null | undefined,
  namespace: string | null,
  name: string
): string | null {
  if (!element || !element.attributes) return null;

  const attrs = element.attributes as Record<string, string>;

  // Try with namespace prefix first
  if (namespace) {
    const prefixedName = `${namespace}:${name}`;
    if (prefixedName in attrs) {
      return attrs[prefixedName];
    }
  }

  // Try without namespace
  if (name in attrs) {
    return attrs[name];
  }

  return null;
}

/**
 * Get an attribute value, trying multiple possible names
 *
 * @param element - Element to get attribute from
 * @param names - Array of possible attribute names (with or without namespace)
 * @returns First found attribute value or null
 */
export function getAttributeAny(
  element: XmlElement | null | undefined,
  names: string[]
): string | null {
  if (!element || !element.attributes) return null;

  const attrs = element.attributes as Record<string, string>;

  for (const name of names) {
    if (name in attrs) {
      return attrs[name];
    }
  }

  return null;
}

/**
 * Get all attributes from an element
 *
 * @param element - Element to get attributes from
 * @returns Record of attribute name -> value
 */
export function getAttributes(element: XmlElement | null | undefined): Record<string, string> {
  if (!element || !element.attributes) return {};
  return element.attributes as Record<string, string>;
}

/**
 * Get the text content of an element (concatenates all text nodes)
 *
 * @param element - Element to get text from
 * @returns Text content or empty string
 */
export function getTextContent(element: XmlElement | null | undefined): string {
  if (!element) return '';

  // Check for direct text property
  if ('text' in element && typeof element.text === 'string') {
    return element.text;
  }

  // Check elements array for text nodes
  if (!element.elements) return '';

  let text = '';
  for (const child of element.elements) {
    if (child.type === 'text' && 'text' in child) {
      text += child.text ?? '';
    } else if (child.type === 'element') {
      // Recurse into child elements
      text += getTextContent(child);
    }
  }

  return text;
}

/**
 * Check if an element has a specific attribute with value "true" or "1"
 *
 * @param element - Element to check
 * @param namespace - Attribute namespace
 * @param name - Attribute name
 * @returns true if attribute exists and is truthy
 */
export function hasFlag(
  element: XmlElement | null | undefined,
  namespace: string | null,
  name: string
): boolean {
  const value = getAttribute(element, namespace, name);

  // In OOXML, presence of element often means true, absence means false
  // If value is null, check if the element itself exists
  if (value === null) {
    return false;
  }

  // Explicitly false
  if (value === '0' || value === 'false' || value === 'off') {
    return false;
  }

  // Any other value (including "1", "true", "on", or empty string) means true
  return true;
}

/**
 * Check if a child element exists (used for boolean flags in OOXML)
 *
 * @param parent - Parent element
 * @param namespace - Namespace prefix
 * @param localName - Local element name
 * @returns true if child element exists
 */
export function hasChild(
  parent: XmlElement | null | undefined,
  namespace: string,
  localName: string
): boolean {
  return findChild(parent, namespace, localName) !== null;
}

/**
 * Parse an OOXML color value
 *
 * @param element - Color element (e.g., w:color)
 * @returns Object with val, themeColor, themeTint, themeShade
 */
export function parseColorElement(element: XmlElement | null | undefined): {
  val?: string;
  themeColor?: string;
  themeTint?: string;
  themeShade?: string;
} | null {
  if (!element) return null;

  return {
    val: getAttribute(element, 'w', 'val') ?? undefined,
    themeColor: getAttribute(element, 'w', 'themeColor') ?? undefined,
    themeTint: getAttribute(element, 'w', 'themeTint') ?? undefined,
    themeShade: getAttribute(element, 'w', 'themeShade') ?? undefined,
  };
}

/**
 * Parse a numeric value from an attribute, with optional scale
 *
 * @param element - Element containing the attribute
 * @param namespace - Attribute namespace
 * @param name - Attribute name
 * @param scale - Optional scale factor (e.g., 20 for twips to points)
 * @returns Parsed number or undefined
 */
export function parseNumericAttribute(
  element: XmlElement | null | undefined,
  namespace: string | null,
  name: string,
  scale: number = 1
): number | undefined {
  const value = getAttribute(element, namespace, name);
  if (value === null) return undefined;

  const num = parseInt(value, 10);
  if (isNaN(num)) return undefined;

  return num * scale;
}

/**
 * Parse a boolean value from an attribute or element presence
 *
 * OOXML boolean conventions:
 * - Element presence with no val attribute = true
 * - w:val="true" or w:val="1" = true
 * - w:val="false" or w:val="0" = false
 *
 * @param element - Element to check
 * @param namespace - Namespace for val attribute
 * @returns boolean value
 */
export function parseBooleanElement(
  element: XmlElement | null | undefined,
  namespace: string = 'w'
): boolean {
  if (!element) return false;

  const val = getAttribute(element, namespace, 'val');

  // No val attribute = true (element presence implies true)
  if (val === null) return true;

  // Explicit false values
  if (val === '0' || val === 'false' || val === 'off') {
    return false;
  }

  return true;
}

/**
 * Deep find - search recursively for an element
 *
 * @param root - Root element to search from
 * @param namespace - Namespace prefix
 * @param localName - Local element name
 * @returns First matching element found or null
 */
export function findDeep(
  root: XmlElement | null | undefined,
  namespace: string,
  localName: string
): XmlElement | null {
  if (!root) return null;

  // Check if this element matches
  if (matchesName(root, namespace, localName)) {
    return root;
  }

  // Search children
  if (root.elements) {
    for (const child of root.elements) {
      if (child.type !== 'element') continue;

      const found = findDeep(child, namespace, localName);
      if (found) return found;
    }
  }

  return null;
}

/**
 * Find all elements matching name, searching recursively
 *
 * @param root - Root element to search from
 * @param namespace - Namespace prefix
 * @param localName - Local element name
 * @returns Array of all matching elements
 */
export function findAllDeep(
  root: XmlElement | null | undefined,
  namespace: string,
  localName: string
): XmlElement[] {
  const results: XmlElement[] = [];

  function search(element: XmlElement | null | undefined): void {
    if (!element) return;

    if (matchesName(element, namespace, localName)) {
      results.push(element);
    }

    if (element.elements) {
      for (const child of element.elements) {
        if (child.type === 'element') {
          search(child);
        }
      }
    }
  }

  search(root);
  return results;
}

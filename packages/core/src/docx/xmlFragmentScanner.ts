/**
 * Runtime-independent XML well-formedness scan for a preserved fragment.
 *
 * Split out of `xmlParser` because it is a parser in its own right: `xml-js` is
 * not validating (it silently closes `<w:t>unclosed`) and `DOMParser` does not
 * exist in a bare Node process, where `DocxReviewer`, server-side generation and
 * the `agents` package all run this code.
 */

/** Recursion bound for walks over file- or clipboard-derived markup. */
export const MAX_ELEMENT_DEPTH = 64;

/** Outcome of {@link scanXmlFragment}. `ok: false` means "do not write this". */
export type FragmentScan = {
  ok: boolean;
  rootName: string | null;
  /** Prefixes used without a binding declared inside the fragment itself. */
  unbound: Set<string>;
};

/** `&` must open one of the five predefined entities or a character reference. */
const ENTITY_RE = /^&(?:amp|lt|gt|quot|apos|#([0-9]+)|#x([0-9a-fA-F]+));/;

/**
 * XML NameChar, minus the unicode ranges: every non-ASCII codepoint is allowed
 * through (enumerating them buys nothing — Word's names are all ASCII) while
 * ASCII is held to the spec, so `<w:t?/>` is a name error rather than an
 * element called `w:t?`.
 */
const NAME_CHAR_RE = /[A-Za-z0-9_.:-]|[^\x00-\x7F]/;
const NAME_START_RE = /[A-Za-z_]|[^\x00-\x7F]/;

function isNameStart(ch: string): boolean {
  return NAME_START_RE.test(ch);
}

/**
 * A character reference has to name a codepoint XML permits: not NUL, not a
 * surrogate half, not a noncharacter, and inside the Unicode range.
 */
function isValidCharRef(decimal: string | undefined, hex: string | undefined): boolean {
  const code = decimal !== undefined ? Number(decimal) : parseInt(hex ?? '', 16);
  if (!Number.isFinite(code)) return false;
  if (code === 0x9 || code === 0xa || code === 0xd) return true;
  if (code >= 0x20 && code <= 0xd7ff) return true;
  if (code >= 0xe000 && code <= 0xfffd) return true;
  return code >= 0x10000 && code <= 0x10ffff;
}

/** `{prefix}` for `p:local`, `null` prefix when unprefixed, `false` if malformed. */
function splitName(name: string): { prefix: string | null } | false {
  const i = name.indexOf(':');
  if (i < 0) return { prefix: null };
  if (i === 0 || i === name.length - 1) return false;
  if (name.indexOf(':', i + 1) >= 0) return false;
  return { prefix: name.slice(0, i) };
}

/**
 * Well-formedness check that does not depend on a host `DOMParser`, so a
 * fragment gets the SAME verdict in a browser and in a bare Node process —
 * `xml-js` is not a validating parser (it silently closes `<w:t>unclosed`), and
 * the headless consumers (`DocxReviewer`, server-side generation) have no DOM.
 *
 * Deliberately strict about what it accepts: no DTD or processing instruction
 * (XXE), only the predefined entities and character references, and a hard
 * nesting bound. Everything Word emits inside a run passes; anything that does
 * not is dropped rather than left to void the part.
 */
export function scanXmlFragment(xml: string): FragmentScan {
  const unbound = new Set<string>();
  const fail: FragmentScan = { ok: false, rootName: null, unbound };
  // One entry per open element: the prefixes in scope, and the name the
  // closing tag has to match.
  const scopes: Set<string>[] = [];
  const openNames: string[] = [];
  let roots = 0;
  let rootName: string | null = null;
  let i = 0;

  /** Validate entity references in a span of character data or attribute value. */
  const scanText = (end: number): boolean => {
    while (i < end) {
      if (xml[i] === '&') {
        const m = ENTITY_RE.exec(xml.slice(i, i + 14));
        if (!m) return false;
        if ((m[1] !== undefined || m[2] !== undefined) && !isValidCharRef(m[1], m[2])) return false;
        i += m[0].length;
        continue;
      }
      i++;
    }
    return true;
  };

  const readName = (): string | null => {
    if (i >= xml.length || !isNameStart(xml[i])) return null;
    const start = i;
    while (i < xml.length && NAME_CHAR_RE.test(xml[i])) i++;
    return xml.slice(start, i);
  };

  const skipSpace = (): void => {
    while (i < xml.length && /\s/.test(xml[i])) i++;
  };

  /**
   * Character data between tags. Outside the root only whitespace is legal —
   * `garbage<w:t/>` written back verbatim puts text straight inside a `w:r`.
   */
  const scanCharData = (end: number): boolean => {
    if (openNames.length === 0 && xml.slice(i, end).trim().length > 0) return false;
    return scanText(end);
  };

  while (i < xml.length) {
    const lt = xml.indexOf('<', i);
    if (lt < 0) {
      if (!scanCharData(xml.length)) return fail;
      break;
    }
    if (!scanCharData(lt)) return fail;
    i = lt + 1;

    if (xml.startsWith('!--', i)) {
      const end = xml.indexOf('-->', i + 3);
      if (end < 0) return fail;
      const body = xml.slice(i + 3, end);
      // `--` may not appear inside a comment, and it may not end with `-`
      // either — `<!--x--->` reads as content `x-`.
      if (body.includes('--') || body.endsWith('-')) return fail;
      i = end + 3;
      continue;
    }
    // `<!DOCTYPE`/`<!ENTITY` (XXE, billion laughs), `<![CDATA[` and `<?pi?>`
    // have no place in a preserved run fragment.
    if (xml[i] === '!' || xml[i] === '?') return fail;

    if (xml[i] === '/') {
      i++;
      const closing = readName();
      if (closing === null || openNames.length === 0) return fail;
      skipSpace();
      if (xml[i] !== '>') return fail;
      i++;
      if (openNames.pop() !== closing) return fail;
      scopes.pop();
      continue;
    }

    const name = readName();
    if (name === null) return fail;
    if (openNames.length === 0) {
      roots++;
      rootName ??= name;
    }
    if (openNames.length >= MAX_ELEMENT_DEPTH) return fail;

    const inherited = scopes.length > 0 ? scopes[scopes.length - 1] : new Set<string>();
    let scope = inherited;
    const seen = new Set<string>();
    const used: string[] = [];

    for (;;) {
      const beforeSpace = i;
      skipSpace();
      if (i >= xml.length) return fail;
      if (xml[i] === '>' || xml.startsWith('/>', i)) break;
      // XML requires whitespace between attributes; `<w:t a="1"b="2"/>` is not
      // two attributes, it is a syntax error.
      if (i === beforeSpace) return fail;
      const attr = readName();
      if (attr === null || seen.has(attr)) return fail;
      seen.add(attr);
      skipSpace();
      if (xml[i] !== '=') return fail;
      i++;
      skipSpace();
      const quote = xml[i];
      if (quote !== '"' && quote !== "'") return fail;
      i++;
      const close = xml.indexOf(quote, i);
      if (close < 0) return fail;
      if (xml.slice(i, close).includes('<')) return fail;
      if (!scanText(close)) return fail;
      i = close + 1;

      if (attr === 'xmlns') continue;
      if (attr.startsWith('xmlns:')) {
        const declared = attr.slice(6);
        if (declared.length === 0) return fail;
        if (scope === inherited) scope = new Set(inherited);
        scope.add(declared);
        continue;
      }
      const split = splitName(attr);
      if (split === false) return fail;
      if (split.prefix) used.push(split.prefix);
    }

    const ownSplit = splitName(name);
    if (ownSplit === false) return fail;
    if (ownSplit.prefix) used.push(ownSplit.prefix);
    // `xml:` is bound by the spec and never declared.
    for (const prefix of used) {
      if (prefix !== 'xml' && !scope.has(prefix)) unbound.add(prefix);
    }

    if (xml.startsWith('/>', i)) {
      i += 2;
      continue;
    }
    i++;
    scopes.push(scope);
    openNames.push(name);
  }

  if (openNames.length !== 0 || roots !== 1) return fail;
  return { ok: true, rootName, unbound };
}

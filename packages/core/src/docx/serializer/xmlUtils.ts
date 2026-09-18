/**
 * Shared XML utility functions for serializers.
 */

export function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/**
 * Format a numeric value as an integer XML attribute.
 *
 * OOXML measure types (twips, EMU, half-points, eighths-of-point) are
 * integer-typed in the schema (xs:unsignedInt / xs:long / xs:int). Word
 * rejects floating-point values (e.g. `0.7 * 1440 === 1008.0000000000001`
 * surfacing from `inches * TWIPS_PER_INCH`), even though tolerant readers
 * accept them. Coerce to a finite integer at every serialization site.
 *
 * `NaN`/`Infinity`/`null`/`undefined` collapse to `'0'` rather than
 * leaking literal `"NaN"` or `"Infinity"` into the XML.
 */
export function intAttr(value: number | undefined | null): string {
  if (value == null || !Number.isFinite(value)) return '0';
  return String(Math.round(value));
}

/**
 * Extension prefixes a consumer may skip. Declaring a namespace only says what
 * a prefix means; `mc:Ignorable` is what makes an element in it skippable
 * rather than an error, so every root that declares these has to list them too.
 */
export const MC_IGNORABLE_PREFIXES = [
  'w14',
  'w15',
  'w16se',
  'w16cid',
  'w16',
  'w16cex',
  'w16sdtdh',
  'w16sdtfl',
  'w16du',
  'wp14',
];

/** `mc:Ignorable` attribute paired with {@link OOXML_NAMESPACES}. */
export const MC_IGNORABLE = `mc:Ignorable="${MC_IGNORABLE_PREFIXES.join(' ')}"`;

/**
 * `mc:Ignorable` for a root that declares `declared`, carrying over whatever
 * the source root listed. Naming a prefix the root never declares is itself
 * invalid, so the result is intersected with what is actually emitted.
 */
export function buildMcIgnorable(declared: Iterable<string>, captured?: string[]): string {
  const available = new Set(declared);
  const prefixes = [...MC_IGNORABLE_PREFIXES, ...(captured ?? [])].filter(
    (p, i, all) => available.has(p) && all.indexOf(p) === i
  );
  return prefixes.length > 0 ? ` mc:Ignorable="${escapeXml(prefixes.join(' '))}"` : '';
}

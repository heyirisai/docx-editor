/**
 * Prototype-pollution guards for attribute maps derived from untrusted input.
 *
 * Collaboration JSON and the fidelity sidecar both originate outside the trust
 * boundary (a peer's document state, relayed verbatim). Copying those keys onto
 * a plain object with spread or `Object.assign` lets `__proto__`,
 * `constructor`, or `prototype` reach `Object.prototype` and corrupt editor
 * state — or any downstream code that reads a plain object.
 *
 * These helpers drop the dangerous keys instead of throwing: a hostile peer
 * should not be able to break a document open for everyone else, and no
 * legitimate DOCX attribute uses those names.
 */

/** Keys that must never be copied onto a plain object. */
const UNSAFE_KEYS = new Set(['__proto__', 'constructor', 'prototype']);

export function isUnsafeAttrKey(key: string): boolean {
  return UNSAFE_KEYS.has(key);
}

/**
 * Shallow-copy an attribute map, omitting prototype-polluting keys.
 * Returns `undefined` when given `undefined` so optional attrs stay optional.
 */
export function copySafeAttrs<T>(
  attrs: Record<string, T> | undefined
): Record<string, T> | undefined {
  if (!attrs) return undefined;
  const safe: Record<string, T> = {};
  for (const [key, value] of Object.entries(attrs)) {
    if (isUnsafeAttrKey(key)) continue;
    safe[key] = value;
  }
  return safe;
}

/**
 * `Object.assign(target, source)` for untrusted sources, skipping
 * prototype-polluting keys. Mutates and returns `target`.
 */
export function assignSafeAttrs<T>(
  target: Record<string, T>,
  source: Record<string, T> | undefined
): Record<string, T> {
  if (!source) return target;
  for (const [key, value] of Object.entries(source)) {
    if (isUnsafeAttrKey(key)) continue;
    target[key] = value;
  }
  return target;
}

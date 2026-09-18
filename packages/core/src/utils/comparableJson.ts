/**
 * Key-order-independent deep normalizer for structural comparison.
 *
 * Lives here rather than beside its collaboration consumer because the DOCX
 * parser needs it too (`headerFooterSnapshot`), and importing it from
 * `collaboration/sourcePreservingExport` would drag that module's static
 * `jszip` import into the parse graph — making every read-only consumer of
 * `parseDocx` pay for the zip library at startup.
 */

/**
 * Attributes added solely to support collaboration are not serialized into
 * OOXML and therefore do not constitute a document edit.
 */
export function comparableJson(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(comparableJson);
  if (!value || typeof value !== 'object') return value;

  // Sort keys so the JSON.stringify comparison in valuesEqual() is
  // insertion-order independent. Yjs and the PM serializer can emit the same
  // attrs with different key order, and an order-sensitive compare would read
  // that as a real change and reject the export as a structural change.
  const result: Record<string, unknown> = {};
  for (const [key, child] of Object.entries(value).sort(([a], [b]) =>
    a < b ? -1 : a > b ? 1 : 0
  )) {
    if (key === 'collaborationId') continue;
    result[key] = comparableJson(child);
  }
  return result;
}

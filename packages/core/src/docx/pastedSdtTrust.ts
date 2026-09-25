/**
 * Trust boundary for content-control state arriving through PASTED HTML.
 *
 * The PM `sdt` / `blockSdt` nodes carry markup that the serializer writes into
 * `document.xml` verbatim — the captured `w:sdtPr` / `w:sdtEndPr`, and for a
 * legacy form field the captured `w:fldChar` run sequence. When that state
 * comes from the parser it is a byte-faithful capture of the source part; when
 * it comes from `parseDOM` it is whatever the clipboard HTML said, i.e.
 * attacker-controlled. A crafted `data-legacy-form-field` could otherwise plant
 * an arbitrary field (`DDEAUTO`, `INCLUDETEXT`) in the saved file, and a
 * crafted `data-raw-properties-xml` could close the `w:sdtPr` early and inject
 * runs.
 *
 * The attributes are still needed: copying a content control or a legacy
 * field inside the editor goes through ProseMirror's clipboard, which is
 * `toDOM` → HTML → `parseDOM`. So rather than drop them, each is accepted only
 * if it is exactly the shape the parser itself produces:
 *
 * - `w:sdtPr` / `w:sdtEndPr`: one well-formed element of that name, with no
 *   run content, field or embedded-object markup anywhere inside it.
 * - legacy form field: a well-formed `begin` (+ `w:ffData`) · `instrText` ·
 *   optional `separate` sequence whose instruction is exactly `FORMTEXT`,
 *   `FORMCHECKBOX` or `FORMDROPDOWN`, an `end` run, and an `w:ffData` of known
 *   children only (no `w:entryMacro` / `w:exitMacro`). The modeled state is
 *   re-derived from that `w:ffData` and the descriptor rebuilt from known keys.
 *
 * Anything else is rejected (`null`) and the node keeps only its modeled
 * attrs, which the serializer escapes.
 */

import type { LegacyFormField, LegacyFormFieldType, SdtProperties } from '../types/document';
import { legacyFormFieldTypeFor, parseFfData } from './legacyFormField';
import { getLocalName, isWellFormedXmlElement, parseXml, type XmlElement } from './xmlParser';

/**
 * Elements that must never appear inside pasted property markup: run and block
 * content, field machinery, embedded objects, and markup-compatibility
 * wrappers that could hide either.
 */
const FORBIDDEN_IN_PROPERTIES = new Set([
  'p',
  'r',
  't',
  'tbl',
  'sdt',
  'sdtContent',
  'fldChar',
  'fldSimple',
  'instrText',
  'delInstrText',
  'ffData',
  'hyperlink',
  'object',
  'drawing',
  'pict',
  'AlternateContent',
  'Choice',
  'Fallback',
  'subDoc',
  'altChunk',
]);

/** Child elements of `el` (skipping whitespace text), or null on character data. */
function elementChildren(el: XmlElement): XmlElement[] | null {
  const out: XmlElement[] = [];
  for (const node of el.elements ?? []) {
    if (node.type === 'element') out.push(node);
    else if (node.type === 'text' || node.type === 'cdata') {
      if (String(node.text ?? node.cdata ?? '').trim() !== '') return null;
    }
  }
  return out;
}

/** Whether `el` or any descendant is a forbidden element (depth is scanner-bounded). */
function containsForbidden(el: XmlElement): boolean {
  for (const child of el.elements ?? []) {
    if (child.type !== 'element') continue;
    if (FORBIDDEN_IN_PROPERTIES.has(getLocalName(child.name ?? ''))) return true;
    if (containsForbidden(child)) return true;
  }
  return false;
}

/**
 * Parse a pasted fragment of one or more sibling elements. Null unless the
 * whole thing is well-formed, uses only bound prefixes, and has no character
 * data between the elements.
 */
function parseFragment(xml: string): XmlElement[] | null {
  if (typeof xml !== 'string' || xml.trim() === '') return null;
  const wrapped = `<epPasted>${xml}</epPasted>`;
  if (!isWellFormedXmlElement(wrapped, { requireBoundPrefixes: true })) return null;
  let root: XmlElement | undefined;
  try {
    root = parseXml(wrapped).elements?.find((e) => e.type === 'element');
  } catch {
    return null;
  }
  return root ? elementChildren(root) : null;
}

/**
 * The pasted `w:sdtPr` (or `w:sdtEndPr`) markup if it is safe to write back
 * verbatim, else `null`.
 */
export function trustedPastedSdtPropertiesXml(
  xml: string | null | undefined,
  element: 'sdtPr' | 'sdtEndPr'
): string | null {
  if (!xml) return null;
  const top = parseFragment(xml);
  if (!top || top.length !== 1) return null;
  if (top[0].name !== `w:${element}`) return null;
  return containsForbidden(top[0]) ? null : xml;
}

const LOCKS: ReadonlySet<string> = new Set([
  'sdtLocked',
  'contentLocked',
  'sdtContentLocked',
  'unlocked',
]);

/** A pasted `data-lock` value, if it is one of the `ST_Lock` values. */
export function trustedPastedLock(value: string | null | undefined): SdtProperties['lock'] | null {
  return value && LOCKS.has(value) ? (value as SdtProperties['lock']) : null;
}

// ── legacy form fields ──────────────────────────────────────────────────────

/** `w:ffData` children Word defines, minus the macro hooks (§17.16.17). */
const FFDATA_CHILDREN = new Set([
  'name',
  'label',
  'tabIndex',
  'calcOnExit',
  'enabled',
  'helpText',
  'statusText',
  'checkBox',
  'ddList',
  'textInput',
]);

/** Elements allowed anywhere below those children (checkBox/ddList/textInput content). */
const FFDATA_DESCENDANTS = new Set([
  'size',
  'sizeAuto',
  'default',
  'checked',
  'result',
  'listEntry',
  'type',
  'maxLength',
  'format',
]);

function isWElement(el: XmlElement, local: string): boolean {
  return el.name === `w:${local}`;
}

function fldCharType(el: XmlElement): string | undefined {
  const v = el.attributes?.['w:fldCharType'];
  return typeof v === 'string' ? v : undefined;
}

/** Every descendant of an ffData child is a known, text-free element. */
function ffDataSubtreeOk(el: XmlElement, depth = 0): boolean {
  if (depth > 4) return false;
  const kids = elementChildren(el);
  if (!kids) return false;
  return kids.every(
    (k) =>
      k.name?.startsWith('w:') &&
      FFDATA_DESCENDANTS.has(getLocalName(k.name)) &&
      ffDataSubtreeOk(k, depth + 1)
  );
}

function ffDataOk(ffData: XmlElement): boolean {
  const kids = elementChildren(ffData);
  if (!kids) return false;
  return kids.every(
    (k) =>
      k.name?.startsWith('w:') && FFDATA_CHILDREN.has(getLocalName(k.name)) && ffDataSubtreeOk(k)
  );
}

/** `w:rPr` holding run properties only. */
function rPrOk(rPr: XmlElement): boolean {
  return elementChildren(rPr) !== null && !containsForbidden(rPr);
}

type FieldItem =
  | { kind: 'begin'; ffData: XmlElement }
  | { kind: 'instr'; text: string }
  | { kind: 'separate' }
  | { kind: 'end' };

/**
 * The structural items of a `w:r` in a field sequence — an optional leading
 * `w:rPr`, then only `w:instrText` / `w:fldChar` children, in order (Word
 * usually writes one per run, but may pack an instruction and its separator
 * together). An rPr-only run yields `[]`. Null for anything else.
 */
function runItems(run: XmlElement): FieldItem[] | null {
  if (!isWElement(run, 'r')) return null;
  const kids = elementChildren(run);
  if (!kids) return null;
  let rest = kids;
  if (rest[0] && isWElement(rest[0], 'rPr')) {
    if (!rPrOk(rest[0])) return null;
    rest = rest.slice(1);
  }
  const items: FieldItem[] = [];
  for (const el of rest) {
    if (isWElement(el, 'instrText')) {
      let text = '';
      for (const node of el.elements ?? []) {
        if (node.type === 'text') text += String(node.text ?? '');
        else return null;
      }
      items.push({ kind: 'instr', text });
      continue;
    }
    if (!isWElement(el, 'fldChar')) return null;
    const type = fldCharType(el);
    const fcKids = elementChildren(el);
    if (!fcKids) return null;
    if (type === 'begin') {
      if (fcKids.length !== 1 || !isWElement(fcKids[0], 'ffData') || !ffDataOk(fcKids[0])) {
        return null;
      }
      items.push({ kind: 'begin', ffData: fcKids[0] });
    } else if (fcKids.length === 0 && (type === 'separate' || type === 'end')) {
      items.push({ kind: type });
    } else {
      return null;
    }
  }
  return items;
}

const EXACT_INSTRUCTION = /^\s*(FORMTEXT|FORMCHECKBOX|FORMDROPDOWN)\s*$/i;

/**
 * Validate the captured prefix (`begin` … `separate`, plus any folded empty
 * result runs). Returns the field's ffData element and whether a `separate`
 * was present, or null.
 */
function validatePrefix(
  prefixXml: string,
  instruction: string
): { ffData: XmlElement; hasSeparate: boolean } | null {
  const runs = parseFragment(prefixXml);
  if (!runs || runs.length === 0) return null;
  const items: FieldItem[] = [];
  for (const run of runs) {
    const runFieldItems = runItems(run);
    if (!runFieldItems) return null;
    items.push(...runFieldItems);
  }
  const first = items[0];
  if (!first || first.kind !== 'begin') return null;
  let instr = '';
  let hasSeparate = false;
  for (const item of items.slice(1)) {
    if (item.kind === 'instr' && !hasSeparate) instr += item.text;
    else if (item.kind === 'separate' && !hasSeparate) hasSeparate = true;
    else return null; // a second begin, an end inside the prefix, text after separate
  }
  if (!EXACT_INSTRUCTION.test(instr) || instr.trim() !== instruction.trim()) return null;
  return { ffData: first.ffData, hasSeparate };
}

function validateSuffix(suffixXml: string): boolean {
  const runs = parseFragment(suffixXml);
  if (!runs || runs.length !== 1) return false;
  const items = runItems(runs[0]);
  return !!items && items.length === 1 && items[0].kind === 'end';
}

const FIELD_TYPES: ReadonlySet<string> = new Set(['dropdown', 'checkbox', 'text']);

function isRecord(v: unknown): v is Record<string, unknown> {
  return typeof v === 'object' && v !== null && !Array.isArray(v);
}

/**
 * A pasted `data-legacy-form-field` value, re-serialized from a descriptor
 * rebuilt out of validated parts, or `null` when it is not exactly a legacy
 * `FORMTEXT` / `FORMCHECKBOX` / `FORMDROPDOWN` capture.
 */
export function trustedPastedLegacyFormField(json: string | null | undefined): string | null {
  if (!json) return null;
  let raw: unknown;
  try {
    raw = JSON.parse(json);
  } catch {
    return null;
  }
  if (!isRecord(raw) || raw.kind !== 'legacy') return null;
  const { fieldType, instruction, ffDataXml, rawPrefixXml, rawSuffixXml } = raw;
  if (typeof fieldType !== 'string' || !FIELD_TYPES.has(fieldType)) return null;
  if (typeof instruction !== 'string' || !EXACT_INSTRUCTION.test(instruction)) return null;
  if (legacyFormFieldTypeFor(instruction) !== fieldType) return null;
  if (typeof ffDataXml !== 'string' || typeof rawPrefixXml !== 'string') return null;
  if (typeof rawSuffixXml !== 'string' || !validateSuffix(rawSuffixXml)) return null;

  const prefix = validatePrefix(rawPrefixXml, instruction);
  if (!prefix) return null;
  // The setters patch `ffDataXml` inside the prefix by substring, so it must be
  // the prefix's own `w:ffData`, not a free-standing string.
  const ffTop = parseFragment(ffDataXml);
  if (!ffTop || ffTop.length !== 1 || !isWElement(ffTop[0], 'ffData')) return null;
  if (!rawPrefixXml.includes(ffDataXml)) return null;

  // Re-derive the modeled state from the validated ffData rather than trusting
  // the JSON's copy of it.
  const projection = parseFfData(prefix.ffData);
  if (!projection || projection.fieldType !== fieldType) return null;

  const hasResult = raw.hasResult === true;
  const value = typeof raw.value === 'string' ? raw.value : undefined;
  const field: LegacyFormField = {
    kind: 'legacy',
    fieldType: fieldType as LegacyFormFieldType,
    ...(projection.name != null ? { name: projection.name } : {}),
    ...(projection.options ? { options: projection.options } : {}),
    ...(projection.selectedIndex != null ? { selectedIndex: projection.selectedIndex } : {}),
    ...(fieldType === 'checkbox' || value === undefined ? {} : { value }),
    ...(projection.checked != null ? { checked: projection.checked } : {}),
    ...(projection.sizeAuto ? { sizeAuto: true } : {}),
    ...(projection.defaultText != null ? { defaultText: projection.defaultText } : {}),
    instruction: instruction.trim(),
    ffDataXml,
    rawPrefixXml,
    rawSuffixXml,
    hasSeparate: prefix.hasSeparate,
    hasResult,
  };
  return JSON.stringify(field);
}

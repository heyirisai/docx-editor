/**
 * Legacy Word form fields (`w:fldChar` + `w:ffData`) — detection, projection
 * onto the content-control model, targeted `w:ffData` patching, and
 * capture-and-replay serialization.
 *
 * A legacy form field is not an element but a **run sequence**: a `begin`
 * field char carrying `<w:ffData>`, one or more `w:instrText` runs holding
 * ` FORMDROPDOWN ` / ` FORMCHECKBOX ` / ` FORMTEXT `, an optional `separate`,
 * the displayed result runs, and an `end`. Word's own form-protection UI (and
 * every requirements matrix authored before content controls existed) uses
 * these, so Iris has to stage and answer them exactly as it does `w:sdt`
 * controls.
 *
 * Rather than add a parallel node type through the parser → PM → painter →
 * serializer chain, the parser projects the sequence onto an {@link InlineSdt}
 * whose `properties.legacyFormField` carries the legacy state. Everything
 * downstream — `findContentControls`, `setContentControlValue`, the PM `sdt`
 * node, the painter's inline widget, `ContentControlWidgets` — then works
 * unchanged, and callers get one list with a `source` discriminator.
 *
 * Round-trip contract (same as `w:sdtPr`): the `begin…separate` runs and the
 * `end` run are captured verbatim and replayed on save, with only the
 * `<w:ffData>` region rewritten by a targeted string patch. Nothing is
 * re-synthesized, so `w:enabled`, `w:calcOnExit`, macros, help text and
 * `w:textInput` formats survive untouched.
 */

import type {
  InlineSdt,
  LegacyFormField,
  LegacyFormFieldType,
  Run,
  SdtProperties,
  SdtType,
} from '../types/document';
import { elementToXml, findChild, getAttribute, getLocalName, type XmlElement } from './xmlParser';

// ============================================================================
// DETECTION
// ============================================================================

/** Field instruction name → legacy form-field type. */
const INSTRUCTION_TO_FIELD_TYPE: Record<string, LegacyFormFieldType> = {
  FORMDROPDOWN: 'dropdown',
  FORMCHECKBOX: 'checkbox',
  FORMTEXT: 'text',
};

/** `w:ffData` child element → legacy form-field type. */
const FFDATA_CHILD_TO_FIELD_TYPE: Record<string, LegacyFormFieldType> = {
  ddList: 'dropdown',
  checkBox: 'checkbox',
  textInput: 'text',
};

/**
 * Which legacy form field an instruction denotes, or `null` if it is an
 * ordinary field (PAGE, REF, MERGEFIELD, …).
 */
export function legacyFormFieldTypeFor(instruction: string): LegacyFormFieldType | null {
  const name = instruction.trim().toUpperCase().split(/\s+/)[0] ?? '';
  return INSTRUCTION_TO_FIELD_TYPE[name] ?? null;
}

/**
 * The `<w:ffData>` element carried by a run's `begin` field char, if any.
 * Returns `null` for every other run — including a `begin` char with no
 * `w:ffData`, which is an ordinary complex field.
 */
export function findFfDataElement(runElement: XmlElement): XmlElement | null {
  for (const child of runElement.elements ?? []) {
    if (child.type !== 'element' || getLocalName(child.name ?? '') !== 'fldChar') continue;
    if (getAttribute(child, 'w', 'fldCharType') !== 'begin') continue;
    return findChild(child, 'w', 'ffData');
  }
  return null;
}

/** Control-type projection for each legacy field, so callers can filter by `sdtType`. */
export const LEGACY_FIELD_SDT_TYPE: Record<LegacyFormFieldType, SdtType> = {
  dropdown: 'dropDownList',
  checkbox: 'checkbox',
  // OOXML has no `text` SDT marker; `plainText` is the modeled equivalent and
  // keeps `isTextReplaceable` (and therefore setContentControlContent) working.
  text: 'plainText',
};

/** Word's rendered glyphs for an unchecked / checked legacy checkbox. */
export const LEGACY_CHECKBOX_GLYPHS = { checked: '☒', unchecked: '☐' } as const;

/** The glyph Word paints for a legacy checkbox in the given state. */
export function legacyCheckboxGlyph(checked: boolean | undefined): string {
  return checked ? LEGACY_CHECKBOX_GLYPHS.checked : LEGACY_CHECKBOX_GLYPHS.unchecked;
}

/**
 * The blank Word shows for a `FORMTEXT` field with neither a result nor a
 * `w:textInput/w:default`: five en spaces (U+2002), which Word also writes as
 * the field's initial result when it creates one. En spaces are not
 * line-break opportunities, so the blank paints as one shaded-width gap.
 */
export const LEGACY_TEXT_PLACEHOLDER = '     ';

/** The state a display run is derived from (see {@link legacyFormFieldDisplayText}). */
export type LegacyFormFieldDisplayState = Pick<
  LegacyFormField,
  'fieldType' | 'options' | 'selectedIndex' | 'checked' | 'defaultText'
>;

/**
 * What Word displays for a legacy form field that carries no result of its
 * own (ECMA-376 §17.16.19 / §17.16.7 / §17.16.33 and Word's behaviour):
 *
 * - dropdown → the `w:listEntry` at `w:ddList/w:result` (default index 0)
 * - checkbox → ☒/☐ from `w:checked`, falling back to `w:default`
 * - text     → `w:textInput/w:default`, else the five-space blank
 *
 * The parser synthesizes a run with this text so the painted page matches Word,
 * and the serializer uses the same function to recognise (and drop) that run on
 * save so an untouched field round-trips byte for byte.
 */
export function legacyFormFieldDisplayText(field: LegacyFormFieldDisplayState): string {
  switch (field.fieldType) {
    case 'dropdown':
      return field.options?.[field.selectedIndex ?? 0] ?? '';
    case 'checkbox':
      return legacyCheckboxGlyph(field.checked);
    case 'text':
      return field.defaultText || LEGACY_TEXT_PLACEHOLDER;
  }
}

// ============================================================================
// w:ffData PROJECTION
// ============================================================================

/** `CT_OnOff`-style read: a present element with no `w:val` means "on". */
function onOff(el: XmlElement | null): boolean | undefined {
  if (!el) return undefined;
  const v = getAttribute(el, 'w', 'val');
  return v == null || !/^(0|false|off)$/i.test(v);
}

/** Read a `CT_DecimalNumber`'s `w:val` as a non-negative integer. */
function decimal(el: XmlElement | null): number | undefined {
  if (!el) return undefined;
  const raw = getAttribute(el, 'w', 'val');
  if (raw == null) return undefined;
  const n = parseInt(raw, 10);
  return Number.isFinite(n) && n >= 0 ? n : undefined;
}

/** The first `w:ffData` child that identifies the field type, if present. */
function ffDataTypeElement(
  ffData: XmlElement
): { type: LegacyFormFieldType; element: XmlElement } | null {
  for (const child of ffData.elements ?? []) {
    if (child.type !== 'element') continue;
    const mapped = FFDATA_CHILD_TO_FIELD_TYPE[getLocalName(child.name ?? '')];
    if (mapped) return { type: mapped, element: child };
  }
  return null;
}

/** The modeled projection of a `<w:ffData>` element (no raw capture). */
export interface FfDataProjection {
  fieldType: LegacyFormFieldType;
  name?: string;
  options?: string[];
  selectedIndex?: number;
  checked?: boolean;
  sizeAuto?: boolean;
  defaultText?: string;
}

/**
 * Project a `<w:ffData>` element onto its modeled fields. Returns `null` when
 * the element carries no recognizable form-field child, so the caller can fall
 * back to the opaque complex-field passthrough.
 */
export function parseFfData(ffData: XmlElement): FfDataProjection | null {
  const typed = ffDataTypeElement(ffData);
  if (!typed) return null;

  const out: FfDataProjection = { fieldType: typed.type };
  const name = getAttribute(findChild(ffData, 'w', 'name'), 'w', 'val');
  if (name != null && name !== '') out.name = name;

  if (typed.type === 'dropdown') {
    const options: string[] = [];
    for (const child of typed.element.elements ?? []) {
      if (child.type === 'element' && getLocalName(child.name ?? '') === 'listEntry') {
        options.push(getAttribute(child, 'w', 'val') ?? '');
      }
    }
    out.options = options;
    // `w:result` is the current selection; `w:default` the fallback index.
    out.selectedIndex =
      decimal(findChild(typed.element, 'w', 'result')) ??
      decimal(findChild(typed.element, 'w', 'default')) ??
      0;
  } else if (typed.type === 'checkbox') {
    // Word writes `w:checked` only once the box has been ticked; before that
    // `w:default` is the state, so read checked ?? default (§17.16.7/§17.16.9).
    out.checked =
      onOff(findChild(typed.element, 'w', 'checked')) ??
      onOff(findChild(typed.element, 'w', 'default')) ??
      false;
    out.sizeAuto = findChild(typed.element, 'w', 'sizeAuto') != null;
  } else if (typed.type === 'text') {
    // `w:textInput/w:default` is what Word shows until the field is answered.
    const def = getAttribute(findChild(typed.element, 'w', 'default'), 'w', 'val');
    if (def != null && def !== '') out.defaultText = def;
  }

  return out;
}

// ============================================================================
// PARSE → InlineSdt
// ============================================================================

/** Plain text of a run's text/tab/symbol content (the field's displayed result). */
function runText(run: Run): string {
  let text = '';
  for (const content of run.content) {
    if (content.type === 'text') text += content.text;
    else if (content.type === 'tab') text += '\t';
    else if (content.type === 'symbol') text += content.char;
  }
  return text;
}

/** The pieces of a parsed legacy form-field run sequence. */
export interface LegacyFormFieldSequence {
  /** The `<w:ffData>` element from the `begin` field char. */
  ffData: XmlElement;
  /** Accumulated `w:instrText` content (untrimmed). */
  instruction: string;
  /** Verbatim XML of the runs from `begin` through `separate`. */
  prefixXml: string;
  /** Verbatim XML of the `end` run. */
  suffixXml: string;
  /** Whether the sequence carried a `separate` field char. */
  hasSeparate: boolean;
  /** Parsed result runs between `separate` and `end`. */
  resultRuns: Run[];
  /**
   * Verbatim XML of {@link resultRuns}. When they carry no text (e.g. a lone
   * `<w:r><w:rPr/></w:r>`), they are folded into the replayed prefix so the
   * synthesized display run can take their place without changing the bytes.
   */
  resultXml?: string;
  /**
   * Run formatting of the field's structural runs (the `w:rPr` on the run
   * carrying `fldChar begin`, or the first field run that has one). A
   * synthesized display run inherits it, as Word formats the result from the
   * field code's run properties.
   */
  formatting?: Run['formatting'];
}

/**
 * Build the {@link InlineSdt} projection of a legacy form-field run sequence,
 * or `null` when the sequence is not a (well-formed) legacy form field — the
 * caller then keeps its existing opaque `complexField` passthrough, so a
 * malformed or unknown field is never worsened by this path.
 *
 * The control's `tag` and `alias` are projected from `w:ffData/w:name`, which
 * is what Word shows in its form-field dialog and the only stable identifier a
 * legacy field has — so `findContentControls({ tag })` addresses them too.
 */
export function buildLegacyFormFieldSdt(seq: LegacyFormFieldSequence): InlineSdt | null {
  const instruction = seq.instruction.trim();
  const byInstruction = legacyFormFieldTypeFor(instruction);
  if (!byInstruction) return null;
  const projection = parseFfData(seq.ffData);
  // Trust the instruction, but refuse a sequence whose ffData contradicts it
  // (e.g. ` FORMTEXT ` over a `w:ddList`): patching it would corrupt the field.
  if (!projection || projection.fieldType !== byInstruction) return null;
  if (!seq.suffixXml) return null; // no `end` run captured — not a closed field

  // A result exists only if it shows something. Result runs with no text
  // (Word and other generators emit a bare `<w:r><w:rPr/></w:r>` there) are
  // replayed verbatim as part of the prefix and the display is synthesized.
  const resultText = seq.resultRuns.map(runText).join('');
  const hasResult = resultText !== '';
  const prefixXml = hasResult ? seq.prefixXml : seq.prefixXml + (seq.resultXml ?? '');

  const displayState: LegacyFormFieldDisplayState = {
    fieldType: projection.fieldType,
    ...(projection.options ? { options: projection.options } : {}),
    ...(projection.selectedIndex != null ? { selectedIndex: projection.selectedIndex } : {}),
    ...(projection.checked != null ? { checked: projection.checked } : {}),
    ...(projection.defaultText != null ? { defaultText: projection.defaultText } : {}),
  };

  // `value` is the field's answer: the selected entry, the result text, or —
  // for an unanswered text field — its default (never the blank placeholder).
  const value =
    projection.fieldType === 'dropdown'
      ? legacyFormFieldDisplayText(displayState)
      : hasResult
        ? resultText
        : (projection.defaultText ?? '');

  const legacyFormField: LegacyFormField = {
    kind: 'legacy',
    fieldType: projection.fieldType,
    ...(projection.name != null ? { name: projection.name } : {}),
    ...(projection.options ? { options: projection.options } : {}),
    ...(projection.selectedIndex != null ? { selectedIndex: projection.selectedIndex } : {}),
    ...(projection.fieldType === 'checkbox' ? {} : { value }),
    ...(projection.checked != null ? { checked: projection.checked } : {}),
    ...(projection.sizeAuto ? { sizeAuto: true } : {}),
    ...(projection.defaultText != null ? { defaultText: projection.defaultText } : {}),
    instruction,
    ffDataXml: elementToXml(seq.ffData),
    rawPrefixXml: prefixXml,
    rawSuffixXml: seq.suffixXml,
    hasSeparate: seq.hasSeparate,
    hasResult,
  };

  const properties: SdtProperties = {
    sdtType: LEGACY_FIELD_SDT_TYPE[projection.fieldType],
    ...(projection.name != null ? { tag: projection.name, alias: projection.name } : {}),
    ...(projection.options
      ? { listItems: projection.options.map((o) => ({ displayText: o, value: o })) }
      : {}),
    ...(projection.checked != null ? { checked: projection.checked } : {}),
    legacyFormField,
  };

  // With no result of its own, Word still displays something — the current
  // list entry, the checkbox glyph, the text default or the blank. Synthesize
  // that run (in the field's run formatting) so the painted page and the
  // clickable inline widget, which attaches to a text run, match Word. The
  // serializer drops it again while `hasResult` is false and the text still
  // matches, keeping the saved bytes identical.
  const content: InlineSdt['content'] = hasResult
    ? seq.resultRuns
    : [
        {
          type: 'run',
          ...(seq.formatting ? { formatting: seq.formatting } : {}),
          content: [{ type: 'text', text: legacyFormFieldDisplayText(displayState) }],
        },
      ];

  return { type: 'inlineSdt', properties, content };
}

// ============================================================================
// w:ffData PATCHING (targeted string edits, like the raw w:sdtPr contract)
// ============================================================================

/** Escape a qualified XML name (an NCName, possibly prefixed) for use in a RegExp. */
const escapeRe = (s: string): string => s.replace(/[.*+?^${}()|[\]\\-]/g, '\\$&');

/**
 * The qualified names a captured `ffData` uses, read from its own bytes so a
 * patch speaks the file's dialect. The parser accepts the elements by *local*
 * name, so a file binding the WordprocessingML namespace to another prefix (or
 * to the default namespace) projects exactly like a `w:` one — patching with
 * literal `w:*` names would then silently miss, leaving the modeled state and
 * the replayed bytes disagreeing after save.
 */
interface FfDataNames {
  /** `local` → the element's qualified name in this ffData (`w:local`, `x:local`, `local`). */
  el: (local: string) => string;
  /** Qualified name of the `val` attribute as this ffData writes it (`w:val` normally). */
  val: string;
}

function ffDataNames(ffDataXml: string): FfDataNames {
  const prefix = /^<(?:([^\s<>/:]+):)?ffData\b/.exec(ffDataXml)?.[1] ?? '';
  const el = (local: string): string => (prefix ? `${prefix}:${local}` : local);
  // Attributes carry their own prefix (or none); follow the one already in
  // use — `w:name w:val` is all but universal — and fall back to the element's.
  const val = /\s((?:[^\s<>/:=]+:)?val)=/.exec(ffDataXml)?.[1] ?? el('val');
  return { el, val };
}

/** `<qname …/>` or `<qname …>` matcher for a known-safe qualified name. */
function elementOpenTag(xml: string, qname: string): RegExpExecArray | null {
  return new RegExp(`<${escapeRe(qname)}\\b[^>]*?(/?)>`).exec(xml);
}

/**
 * Set the `val` attribute on the first `<local>` element inside `xml`, or
 * insert the element (via `insert`) when it is absent.
 */
function setDecimalOrOnOff(
  xml: string,
  names: FfDataNames,
  local: string,
  value: string,
  insert: (x: string) => string
): string {
  const m = elementOpenTag(xml, names.el(local));
  if (!m) return insert(xml);
  const open = m[0];
  const selfClose = m[1] === '/';
  const body = open.slice(1, selfClose ? -2 : -1);
  const valAttr = new RegExp(`(^|\\s)${escapeRe(names.val)}="[^"]*"`);
  const next = valAttr.test(body)
    ? body.replace(valAttr, (_, lead: string) => `${lead}${names.val}="${value}"`)
    : `${body} ${names.val}="${value}"`;
  return xml.replace(open, () => `<${next}${selfClose ? '/>' : '>'}`);
}

/** Insert `fragment` immediately after the opening tag of `container`. */
function insertAfterOpenTag(xml: string, container: string, fragment: string): string {
  const m = new RegExp(`<${escapeRe(container)}\\b[^>]*>`).exec(xml);
  if (!m || m[0].endsWith('/>')) return xml;
  return xml.slice(0, m.index + m[0].length) + fragment + xml.slice(m.index + m[0].length);
}

/** Insert `fragment` immediately before the closing tag of `container`. */
function insertBeforeCloseTag(xml: string, container: string, fragment: string): string {
  const close = `</${container}>`;
  const at = xml.lastIndexOf(close);
  if (at < 0) return xml;
  return xml.slice(0, at) + fragment + xml.slice(at);
}

/**
 * Insert `fragment` immediately before the first `<before>` element inside
 * `container`, or before `</container>` when there is none.
 */
function insertBeforeChildOrClose(
  xml: string,
  container: string,
  before: string,
  fragment: string
): string {
  const open = new RegExp(`<${escapeRe(container)}\\b[^>]*>`).exec(xml);
  const close = xml.lastIndexOf(`</${container}>`);
  if (open && close > open.index) {
    const bodyStart = open.index + open[0].length;
    const child = new RegExp(`<${escapeRe(before)}[\\s/>]`).exec(xml.slice(bodyStart, close));
    if (child) {
      const at = bodyStart + child.index;
      return xml.slice(0, at) + fragment + xml.slice(at);
    }
  }
  return insertBeforeCloseTag(xml, container, fragment);
}

/**
 * Rewrite a legacy field's `<w:ffData>` and keep the captured prefix XML in
 * step. Only the ffData region changes; every other captured byte is replayed.
 */
function withFfData(field: LegacyFormField, nextFfDataXml: string): LegacyFormField {
  return {
    ...field,
    ffDataXml: nextFfDataXml,
    rawPrefixXml: field.rawPrefixXml.replace(field.ffDataXml, () => nextFfDataXml),
  };
}

/**
 * Select a dropdown entry by 0-based index, writing `w:ddList/w:result`.
 * `CT_FFDDList` is `result?, default?, listEntry*`, so a missing `w:result` is
 * inserted directly after the `<w:ddList>` open tag to stay sequence-valid.
 */
export function setLegacyDropdownIndex(field: LegacyFormField, index: number): LegacyFormField {
  const n = ffDataNames(field.ffDataXml);
  const ff = setDecimalOrOnOff(field.ffDataXml, n, 'result', String(index), (xml) =>
    insertAfterOpenTag(xml, n.el('ddList'), `<${n.el('result')} ${n.val}="${index}"/>`)
  );
  return {
    ...withFfData(field, ff),
    selectedIndex: index,
    value: field.options?.[index] ?? '',
    // The caller writes the entry as the result run; it is a real result now.
    hasResult: true,
  };
}

/**
 * Set a legacy checkbox's state. Word stores the *current* state in
 * `w:checkBox/w:checked` and the initial state in `w:default`; it writes both
 * when the user ticks a protected form, and a file with only `w:default`
 * re-renders from that on open — so both are written here.
 * `CT_FFCheckBox` is `(size|sizeAuto), default?, checked?`: an absent
 * `w:default` goes after the size choice and before any existing `w:checked`
 * (Word writes `w:checked` without `w:default`, and a `default` appended after
 * it breaks the sequence), and an absent `w:checked` is appended last.
 */
export function setLegacyCheckbox(field: LegacyFormField, checked: boolean): LegacyFormField {
  const n = ffDataNames(field.ffDataXml);
  const val = checked ? '1' : '0';
  let ff = setDecimalOrOnOff(field.ffDataXml, n, 'default', val, (xml) =>
    insertBeforeChildOrClose(
      xml,
      n.el('checkBox'),
      n.el('checked'),
      `<${n.el('default')} ${n.val}="${val}"/>`
    )
  );
  ff = setDecimalOrOnOff(ff, n, 'checked', val, (xml) =>
    insertBeforeCloseTag(xml, n.el('checkBox'), `<${n.el('checked')} ${n.val}="${val}"/>`)
  );
  return { ...withFfData(field, ff), checked };
}

/** Set the displayed text of a legacy `FORMTEXT` field (`w:ffData` is untouched). */
export function setLegacyText(field: LegacyFormField, text: string): LegacyFormField {
  return { ...field, value: text, hasResult: true };
}

/**
 * Bring a legacy `FORMTEXT` descriptor into step with content written through
 * the *generic* paths — `setContentControlContent`, or typing into the field in
 * the editor — which replace the control's runs without calling
 * {@link setLegacyText}. `value` mirrors the new text and `hasResult` flips on
 * once the text is no longer the parser-synthesized display, so
 * `findContentControls` reports the same answer the saved file will carry.
 *
 * Every other case returns `field` itself: dropdown and checkbox state lives in
 * `w:ffData` and moves only through the typed setters, and a display that is
 * still the synthesized default is not an answer (leaving `hasResult` false is
 * what keeps an untouched field byte-identical on save).
 */
export function syncLegacyTextField(field: LegacyFormField, text: string): LegacyFormField {
  if (field.fieldType !== 'text') return field;
  const unchanged = field.hasResult
    ? field.value === text
    : text === legacyFormFieldDisplayText(field);
  return unchanged ? field : setLegacyText(field, text);
}

/**
 * {@link syncLegacyTextField} over an inline control's content model. Content
 * that is not plain runs (a hyperlink, a nested field or control) is left
 * alone — it is not a text answer the descriptor could mirror.
 */
export function syncLegacyFormFieldContent(
  props: SdtProperties,
  content: InlineSdt['content']
): SdtProperties {
  const field = props.legacyFormField;
  if (!field || field.fieldType !== 'text') return props;
  const text = inlineRunsText(content);
  if (text === null) return props;
  const next = syncLegacyTextField(field, text);
  return next === field ? props : { ...props, legacyFormField: next };
}

// ============================================================================
// SERIALIZATION
// ============================================================================

const SEPARATE_RUN = '<w:r><w:fldChar w:fldCharType="separate"/></w:r>';

/**
 * Plain text of an inline control's content, or `null` when it holds anything
 * other than runs (a hyperlink, field or nested control is never the
 * parser-synthesized display run).
 */
function inlineRunsText(content: InlineSdt['content']): string | null {
  let text = '';
  for (const item of content) {
    if (item.type !== 'run') return null;
    text += runText(item);
  }
  return text;
}

/**
 * Whether the control's content is still the display run the parser
 * synthesized for a result-less field — in which case it must not be written
 * back as a real result. A checkbox never gets one: Word derives the glyph from
 * `w:checked`, and a text run there would double it. For a dropdown/text
 * field the run is dropped only while its text matches what Word would show
 * anyway; content typed over it in the editor is a real answer and is kept.
 */
function isSynthesizedDisplay(
  field: LegacyFormField,
  content: InlineSdt['content'] | undefined
): boolean {
  if (field.hasResult) return false;
  if (field.fieldType === 'checkbox' || content === undefined) return true;
  return inlineRunsText(content) === legacyFormFieldDisplayText(field);
}

/**
 * Replay a legacy form field: the captured `begin…separate` runs verbatim, the
 * (possibly rewritten) result runs, then the captured `end` run.
 *
 * A field whose source carried no result emits none while its content is still
 * the parser-synthesized display run (see {@link isSynthesizedDisplay}) —
 * which is what makes an untouched field round-trip byte for byte.
 *
 * @param field - the legacy field state (already patched by any value setter)
 * @param resultXml - serialized result runs from the control's content
 * @param content - the control's content model, used to recognise the
 *   synthesized display run; omitted = treat a result-less field as untouched
 */
export function serializeLegacyFormField(
  field: LegacyFormField,
  resultXml: string,
  content?: InlineSdt['content']
): string {
  const result = isSynthesizedDisplay(field, content) ? '' : resultXml;
  // A source without `separate` has no place to put a result; open one only
  // when there is actually something to show.
  const separate = !field.hasSeparate && result ? SEPARATE_RUN : '';
  return `${field.rawPrefixXml}${separate}${result}${field.rawSuffixXml}`;
}

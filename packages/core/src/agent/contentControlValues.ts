/**
 * Typed value setters for block-level content controls — set a dropdown
 * selection, toggle a checkbox, or set a date. These produce both the visible
 * content (the run text Word shows) and the structured state inside the
 * captured raw `w:sdtPr` (dropdown `w:lastValue`, `w14:checked`, `w:date`'s
 * `w:fullDate`), patched in place so the rest of the control round-trips
 * verbatim. Use these instead of {@link setContentControlContent} for typed
 * controls, which that function refuses by design.
 *
 * Raw `w:sdtPr` is patched with targeted string edits (not a full re-serialize)
 * to preserve the `CT_SdtPr` element order and any unmodeled properties — the
 * same capture-and-replay contract used everywhere else for SDTs.
 */

import type {
  Document,
  BlockContent,
  SdtProperties,
  Run,
  InlineSdt,
  LegacyFormField,
} from '../types/document';
import {
  legacyCheckboxGlyph,
  setLegacyCheckbox,
  setLegacyDropdownIndex,
  setLegacyText,
} from '../docx/legacyFormField';
import {
  ContentControlLockedError,
  ContentControlBoundError,
  isContentLocked,
  isDataBound,
  clearShowingPlaceholderXml,
  applyControlMutation,
  type ContentControlFilter,
  type BlockControlOp,
  type InlineControlOp,
} from './contentControls';

/** A typed value to apply to a content control. */
export type ContentControlValue =
  | { kind: 'dropdown'; value: string }
  | { kind: 'checkbox'; checked: boolean }
  | { kind: 'date'; date: string }
  /**
   * Free text. Accepted by legacy `FORMTEXT` fields and by free-form
   * (`richText` / `plainText`) content controls — the typed controls reject it,
   * as they do any other mismatched value kind.
   */
  | { kind: 'text'; text: string };

/** The control doesn't support the requested value kind, or the value is invalid. */
export class ContentControlValueError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'ContentControlValueError';
  }
}

// ── raw w:sdtPr string patching ─────────────────────────────────────────────

/** Read an attribute from the first `<prefix:local ...>` element in `xml`. */
function readAttr(xml: string, element: string, attr: string): string | undefined {
  const el = new RegExp(`<${element}\\b[^>]*>`).exec(xml);
  if (!el) return undefined;
  const m = new RegExp(`\\b${attr}="([^"]*)"`).exec(el[0]);
  return m ? m[1] : undefined;
}

/**
 * Set (or add) an attribute on the first `<prefix:local ...>` element in `xml`.
 * Returns the patched string; if the element isn't present, `xml` is unchanged.
 */
function setAttr(xml: string, element: string, attr: string, value: string): string {
  const tag = new RegExp(`<${element}\\b[^>]*?(/?)>`);
  const m = tag.exec(xml);
  if (!m) return xml;
  const open = m[0];
  const selfClose = m[1] === '/';
  const body = open.slice(1, selfClose ? -2 : -1); // strip "<" and "/>"/">"
  const hasAttr = new RegExp(`\\b${attr}="[^"]*"`).test(body);
  // Use a replacement *function* so `$`-sequences in `value` aren't interpreted.
  const newBody = hasAttr
    ? body.replace(new RegExp(`\\b${attr}="[^"]*"`), () => `${attr}="${value}"`)
    : `${body} ${attr}="${value}"`;
  return xml.replace(open, () => `<${newBody}${selfClose ? '/>' : '>'}`);
}

/** Escape a string for safe interpolation into an XML attribute value. */
function escapeXmlAttr(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** Hex code-point string (e.g. "2612") → the character it denotes. */
function codePointChar(hex: string | undefined, fallback: string): string {
  if (!hex) return fallback;
  const n = parseInt(hex, 16);
  return Number.isNaN(n) ? fallback : String.fromCodePoint(n);
}

// ── date formatting (minimal OOXML w:dateFormat support) ────────────────────

const MONTHS = [
  'January',
  'February',
  'March',
  'April',
  'May',
  'June',
  'July',
  'August',
  'September',
  'October',
  'November',
  'December',
];

/** Format an ISO date (yyyy-mm-dd) with a subset of OOXML date tokens. */
export function formatSdtDate(iso: string, pattern?: string): string {
  const [y, m, d] = iso
    .slice(0, 10)
    .split('-')
    .map((n) => parseInt(n, 10));
  if (!y || !m || !d) return iso;
  const fmt = pattern && pattern.trim() ? pattern : 'M/d/yyyy';
  const pad = (n: number) => String(n).padStart(2, '0');
  // Single pass so an emitted month name (e.g. "March") isn't re-scanned by a
  // later, shorter token like `M` — which would corrupt it to "3arch".
  const tokens: Record<string, string> = {
    yyyy: String(y),
    yy: String(y).slice(-2),
    MMMM: MONTHS[m - 1],
    MMM: MONTHS[m - 1].slice(0, 3),
    MM: pad(m),
    M: String(m),
    dd: pad(d),
    d: String(d),
  };
  return fmt.replace(/yyyy|yy|MMMM|MMM|MM|M|dd|d/g, (t) => tokens[t]);
}

// ── value application ───────────────────────────────────────────────────────

/**
 * A one-run paragraph. `font` sets the run's font (for symbol glyphs);
 * `formatting` is the base run formatting the display run keeps (a legacy
 * field's own run properties), with `font` layered on top.
 */
function paragraph(text: string, font?: string, formatting?: Run['formatting']): BlockContent {
  if (!text) return { type: 'paragraph', content: [] };
  const merged: Run['formatting'] = font
    ? { ...formatting, fontFamily: { ascii: font, hAnsi: font, eastAsia: font, cs: font } }
    : formatting;
  const run: Run = {
    type: 'run',
    content: [{ type: 'text', text }],
    ...(merged && Object.keys(merged).length > 0 ? { formatting: merged } : {}),
  };
  return { type: 'paragraph', content: [run] };
}

/**
 * Formatting of the first run in an inline control's content — the run a
 * typed value replaces, whose formatting the new display run inherits (as
 * Word keeps a form field's result in the field's formatting).
 */
function firstRunFormatting(content: InlineSdt['content']): Run['formatting'] | undefined {
  for (const item of content) {
    if (item.type === 'run') return item.formatting;
  }
  return undefined;
}

/** Clear a control's placeholder state (real content is being written). */
function withoutPlaceholder(props: SdtProperties, nextRaw: string): SdtProperties {
  const cleaned = clearShowingPlaceholderXml(nextRaw);
  return {
    ...props,
    showingPlaceholder: false,
    rawPropertiesXml: (cleaned ?? nextRaw) || undefined,
  };
}

/**
 * Apply a typed value to a **legacy Word form field**. The `w:ffData` state is
 * patched with the same targeted-string contract used for raw `w:sdtPr`, and
 * the display content is rebuilt so the painted page matches Word:
 *
 * - dropdown → `w:ddList/w:result` index + the selected entry as the result run
 * - checkbox → `w:checkBox/w:default` **and** `w:checked` + the ☒/☐ glyph
 * - text     → the result run text (`w:ffData` is left alone)
 *
 * Everything else inside `w:ffData` (`w:enabled`, `w:calcOnExit`, help text,
 * macros, text-input formats) is replayed verbatim.
 *
 * The display run is built in `formatting` — the run formatting of the
 * content it replaces (the parser gives a result-less field's synthesized run
 * the field's own run properties). Without it a FORMDROPDOWN styled Arial 9pt
 * bold would fall back to the document default font once answered, and since
 * the answer is a real result (`hasResult`), Word would show it that way too.
 */
function applyLegacyFormFieldValue(
  props: SdtProperties,
  field: LegacyFormField,
  value: ContentControlValue,
  formatting: Run['formatting'] | undefined
): { properties: SdtProperties; content: BlockContent[] } {
  const done = (next: LegacyFormField, text: string) => ({
    properties: {
      ...props,
      ...(next.checked != null ? { checked: next.checked } : {}),
      legacyFormField: next,
    },
    content: [paragraph(text, undefined, formatting)],
  });

  switch (value.kind) {
    case 'dropdown': {
      if (field.fieldType !== 'dropdown') {
        throw new ContentControlValueError(
          `Legacy form field is a '${field.fieldType}' field, not a dropdown.`
        );
      }
      const options = field.options ?? [];
      const index = options.indexOf(value.value);
      if (index < 0) {
        throw new ContentControlValueError(
          `'${value.value}' is not one of the field's list entries.`
        );
      }
      const next = setLegacyDropdownIndex(field, index);
      return done(next, options[index]);
    }
    case 'checkbox': {
      if (field.fieldType !== 'checkbox') {
        throw new ContentControlValueError(
          `Legacy form field is a '${field.fieldType}' field, not a checkbox.`
        );
      }
      const next = setLegacyCheckbox(field, value.checked);
      return done(next, legacyCheckboxGlyph(value.checked));
    }
    case 'text': {
      if (field.fieldType !== 'text') {
        throw new ContentControlValueError(
          `Legacy form field is a '${field.fieldType}' field, not a text field.`
        );
      }
      return done(setLegacyText(field, value.text), value.text);
    }
    case 'date':
      throw new ContentControlValueError('Legacy form fields have no date type.');
  }
}

/**
 * Compute the new properties + display blocks for applying a typed value, without
 * touching a document. Shared by the headless setter and the editor (PM) path.
 * Throws {@link ContentControlValueError} on a type/value mismatch.
 *
 * Legacy Word form fields (`w:fldChar` + `w:ffData`) are routed to their own
 * applier — they carry no `w:sdtPr` to patch — so callers treat legacy fields
 * and `w:sdt` controls identically.
 *
 * `currentFormatting` is the run formatting of the control's current content
 * (see {@link firstRunFormatting}); a legacy field's new display run keeps it.
 */
export function applyContentControlValue(
  props: SdtProperties,
  value: ContentControlValue,
  currentFormatting?: Run['formatting']
): { properties: SdtProperties; content: BlockContent[] } {
  if (props.legacyFormField) {
    return applyLegacyFormFieldValue(props, props.legacyFormField, value, currentFormatting);
  }
  const raw = props.rawPropertiesXml ?? '';
  switch (value.kind) {
    case 'dropdown': {
      if (props.sdtType !== 'dropDownList' && props.sdtType !== 'comboBox') {
        throw new ContentControlValueError(
          `Control is '${props.sdtType}', not a dropdown/combo box.`
        );
      }
      const item = props.listItems?.find(
        (it) => it.value === value.value || it.displayText === value.value
      );
      if (!item) {
        throw new ContentControlValueError(
          `'${value.value}' is not one of the control's list items.`
        );
      }
      // Always write `w:lastValue` (the stored selection): a dropdown/combo box
      // that has never been picked has no such attribute, and without it Word
      // loses the structured selection on reload even though the display text is
      // right. `setAttr` adds it when absent and is a no-op if the element is
      // missing entirely (e.g. an empty raw), so this is safe.
      const element = props.sdtType === 'comboBox' ? 'w:comboBox' : 'w:dropDownList';
      const nextRaw = setAttr(raw, element, 'w:lastValue', escapeXmlAttr(item.value));
      return {
        properties: withoutPlaceholder(props, nextRaw),
        content: [paragraph(item.displayText)],
      };
    }
    case 'checkbox': {
      if (props.sdtType !== 'checkbox') {
        throw new ContentControlValueError(`Control is '${props.sdtType}', not a checkbox.`);
      }
      if (readAttr(raw, 'w14:checked', 'w14:val') == null) {
        throw new ContentControlValueError(
          'Checkbox control has no <w14:checked> state to update (not a Word checkbox).'
        );
      }
      const stateEl = value.checked ? 'w14:checkedState' : 'w14:uncheckedState';
      const char = codePointChar(readAttr(raw, stateEl, 'w14:val'), value.checked ? '☒' : '☐');
      // The glyph renders in the state's symbol font (e.g. MS Gothic).
      const font = readAttr(raw, stateEl, 'w14:font');
      const nextRaw = setAttr(raw, 'w14:checked', 'w14:val', value.checked ? '1' : '0');
      return {
        properties: { ...withoutPlaceholder(props, nextRaw), checked: value.checked },
        content: [paragraph(char, font)],
      };
    }
    case 'date': {
      if (props.sdtType !== 'date') {
        throw new ContentControlValueError(`Control is '${props.sdtType}', not a date control.`);
      }
      const iso = value.date.slice(0, 10);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(iso)) {
        throw new ContentControlValueError(`Date must be ISO yyyy-mm-dd, got '${value.date}'.`);
      }
      // Local-floating (no trailing Z): a UTC midnight would render as the
      // previous day in timezones behind UTC after Word's local conversion.
      const fullDate = `${iso}T00:00:00`;
      const nextRaw = setAttr(raw, 'w:date', 'w:fullDate', fullDate);
      const pattern = readAttr(raw, 'w:dateFormat', 'w:val');
      return {
        properties: withoutPlaceholder(props, nextRaw),
        content: [paragraph(formatSdtDate(iso, pattern))],
      };
    }
    case 'text': {
      // Free text only fits a free-form control; a typed one would desync its
      // structured state from the visible value.
      if (
        props.sdtType !== 'richText' &&
        props.sdtType !== 'plainText' &&
        props.sdtType !== 'unknown'
      ) {
        throw new ContentControlValueError(
          `Control is '${props.sdtType}'; free text would desync its typed state.`
        );
      }
      return {
        properties: withoutPlaceholder(props, raw),
        content: [paragraph(value.text)],
      };
    }
  }
}

/**
 * Set a typed value (dropdown selection / checkbox / date) on the first control
 * matching `filter` — **block-level OR inline** (inline includes controls inside
 * table cells, and with `includeHeadersFooters: true`, headers/footers) — returning a new
 * {@link Document}. Updates both the visible content and the structured raw
 * state (dropdown `w:lastValue`, `w14:checked`, `w:date/@w:fullDate`), so the
 * result round-trips and Word shows the new value.
 *
 * Pass `{ all: true }` to set the value on **every** control matching `filter`
 * (a value shared across duplicated controls) instead of just the first.
 *
 * Throws `ContentControlNotFoundError` if nothing matches,
 * {@link ContentControlLockedError} if content-locked,
 * {@link ContentControlBoundError} if data-bound (the store would override the
 * write), and {@link ContentControlValueError} if the value doesn't fit the
 * control type. The lock/bound guards are overridable with `{ force: true }`.
 */
export function setContentControlValue(
  doc: Document,
  filter: ContentControlFilter,
  value: ContentControlValue,
  options: { force?: boolean; includeHeadersFooters?: boolean; all?: boolean } = {}
): Document {
  const guard = (props: SdtProperties): void => {
    if (!options.force && isContentLocked(props.lock)) {
      throw new ContentControlLockedError(props.lock, 'edit');
    }
    if (!options.force && isDataBound(props)) {
      throw new ContentControlBoundError();
    }
  };
  const blockOp: BlockControlOp = (control) => {
    guard(control.properties);
    const { properties, content } = applyContentControlValue(control.properties, value);
    return [{ ...control, properties, content }];
  };
  const inlineOp: InlineControlOp = (control) => {
    guard(control.properties);
    // The typed setters render their display value as a single paragraph of
    // runs; lift those runs into the inline control's inline content (mirroring
    // how setContentControlContent fills an inline control).
    const { properties, content } = applyContentControlValue(
      control.properties,
      value,
      firstRunFormatting(control.content)
    );
    const display = content[0];
    const inlineContent = (
      display && display.type === 'paragraph' ? display.content : []
    ) as InlineSdt['content'];
    return [{ ...control, properties, content: inlineContent }];
  };
  return applyControlMutation(
    doc,
    filter,
    blockOp,
    inlineOp,
    options.includeHeadersFooters ?? false,
    undefined,
    options.all ?? false
  );
}

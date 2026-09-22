/**
 * Structured Document Tags / content controls (`w:sdt`) — inline and
 * block variants, plus properties (alias, tag, lock, list items,
 * checkbox state) for the supported SDT types.
 *
 * Legacy Word form fields (`w:fldChar` + `w:ffData`) are projected onto the same
 * shapes — see {@link LegacyFormField} and `SdtProperties.legacyFormField`.
 */

import type { Run } from './run';
import type { Hyperlink, SimpleField, ComplexField } from './link';
import type { MathEquation } from './math';
import type { BlockContent } from './section';

/**
 * SDT type (content control type).
 *
 * Values mirror the `w:sdtPr` type-marker element names from ECMA-376
 * §17.5.2 (`CT_SdtPr`), with two deliberate exceptions:
 * - `checkbox` is the `w14:checkbox` (Office 2010) extension, not a base
 *   OOXML type marker.
 * - `buildingBlockGallery` covers both `w:docPartObj` and `w:docPartList`.
 *
 * A `w:sdtPr` with no type marker means `richText` (the spec default). A
 * type marker the parser does not model maps to `unknown` — it is never
 * coerced to `richText`, so the projection stays honest. Round-trip
 * fidelity does not depend on this enum: the raw `w:sdtPr` is replayed
 * verbatim (see `rawPropertiesXml`).
 */
export type SdtType =
  | 'richText'
  | 'plainText'
  | 'date'
  | 'dropDownList'
  | 'comboBox'
  | 'checkbox'
  | 'picture'
  | 'buildingBlockGallery'
  | 'group'
  | 'equation'
  | 'citation'
  | 'bibliography'
  | 'unknown';

/**
 * XML data binding (`w:dataBinding`) — links a content control to a node in a
 * Custom XML data store. Modeled read-only; the binding round-trips verbatim
 * via `rawPropertiesXml` (this projection is for inspection, e.g. "which
 * controls are bound, and to what XPath"). The editor does not resolve or
 * sync bound values.
 */
export interface SdtDataBinding {
  /** XPath into the bound Custom XML part (`w:xpath`). */
  xpath?: string;
  /** Target Custom XML store id (`w:storeItemID`). */
  storeItemID?: string;
  /** Namespace prefix mappings used by the XPath (`w:prefixMappings`). */
  prefixMappings?: string;
}

/**
 * Which legacy Word form field a {@link LegacyFormField} describes. Mirrors the
 * three `w:ffData` children (`w:ddList` / `w:checkBox` / `w:textInput`) and the
 * matching `w:instrText` code (`FORMDROPDOWN` / `FORMCHECKBOX` / `FORMTEXT`).
 */
export type LegacyFormFieldType = 'dropdown' | 'checkbox' | 'text';

/**
 * A legacy Word form field (ECMA-376 §17.16.17, `CT_FFData`) — the pre-content-
 * control protection-based form widget still used by most requirements
 * matrices and RFP response grids.
 *
 * On the wire it is not an element but a *run sequence*:
 *
 * ```xml
 * <w:r><w:fldChar w:fldCharType="begin"><w:ffData>…</w:ffData></w:fldChar></w:r>
 * <w:r><w:instrText xml:space="preserve"> FORMDROPDOWN </w:instrText></w:r>
 * <w:r><w:fldChar w:fldCharType="separate"/></w:r>
 * <w:r><w:t>Never</w:t></w:r>            <!-- the displayed result -->
 * <w:r><w:fldChar w:fldCharType="end"/></w:r>
 * ```
 *
 * The parser projects that sequence onto an {@link InlineSdt} so legacy fields
 * and modern `w:sdt` content controls share one discovery/edit/render path (see
 * `SdtProperties.legacyFormField`); the modeled fields below are a read-only
 * projection, exactly as for `w:sdtPr`. Serialization is capture-and-replay:
 * `rawPrefixXml` (begin → separate) and `rawSuffixXml` (end) are echoed
 * verbatim, with only the `w:ffData` region rewritten by a targeted string
 * patch, so every unmodeled feature (`w:enabled`, `w:calcOnExit`, macros,
 * `w:helpText`, `w:textInput` formats) survives untouched.
 */
export interface LegacyFormField {
  /** Discriminator against modern (`w:sdt`) controls and glyph checkboxes. */
  kind: 'legacy';
  /** Which form field this is. */
  fieldType: LegacyFormFieldType;
  /** Bookmark-style field name (`w:ffData/w:name@w:val`). */
  name?: string;
  /** Dropdown entries in document order (`w:ddList/w:listEntry@w:val`). */
  options?: string[];
  /** Selected dropdown index (`w:ddList/w:result@w:val`), 0-based. */
  selectedIndex?: number;
  /** Current displayed value: the selected option, or the result run's text. */
  value?: string;
  /** Checkbox state — `w:checkBox/w:checked` when present, else `w:default`. */
  checked?: boolean;
  /** `w:checkBox/w:sizeAuto` (size the box from the surrounding font). */
  sizeAuto?: boolean;
  /**
   * Default text of a `FORMTEXT` field (`w:textInput/w:default@w:val`). Word
   * displays it while the field has no result of its own.
   */
  defaultText?: string;
  /** Field instruction, trimmed (e.g. `FORMDROPDOWN`). */
  instruction: string;
  /** The `<w:ffData>` element serialized verbatim, for lossless round-trip. */
  ffDataXml: string;
  /**
   * Verbatim XML of the runs from `fldChar begin` through `fldChar separate`
   * (or through the last field-code run when the source has no separator).
   * Contains {@link ffDataXml}; the two are patched together.
   */
  rawPrefixXml: string;
  /** Verbatim XML of the `fldChar end` run. */
  rawSuffixXml: string;
  /** Whether the source sequence carried a `fldChar separate` run. */
  hasSeparate: boolean;
  /**
   * Whether the source carried a non-empty result between `separate` and
   * `end`. Word writes none for `FORMCHECKBOX` (it draws the box from
   * `w:checked`) and many generators write none for `FORMDROPDOWN` /
   * `FORMTEXT` either — Word then displays the current list entry, the
   * `w:textInput` default, or its blank placeholder. The parser synthesizes
   * that display run (with the field's run formatting) and the serializer drops
   * it again while it still matches, keeping the saved bytes identical to the
   * source. Set to `true` once a value is written through the value setters.
   */
  hasResult: boolean;
}

/**
 * SDT properties (`w:sdtPr`).
 *
 * The modeled fields are a **read-only projection** for downstream tooling
 * (tag/alias addressing, template extraction). They are NOT the
 * serialization source: the original `w:sdtPr` is captured verbatim in
 * `rawPropertiesXml` and replayed on save, which preserves element order
 * (`CT_SdtPr` is an `xsd:sequence`), avoids double-emission, and keeps
 * unmodeled features (data binding, `w15:*`, `@lastValue`) lossless.
 */
export interface SdtProperties {
  /** SDT type (projection; see {@link SdtType}). */
  sdtType: SdtType;
  /** Unique numeric id (`w:id`, signed). */
  id?: number;
  /** Alias (friendly name, `w:alias`). */
  alias?: string;
  /** Tag (developer identifier, `w:tag`). */
  tag?: string;
  /** Lock setting (`w:lock`). */
  lock?: 'sdtLocked' | 'contentLocked' | 'sdtContentLocked' | 'unlocked';
  /**
   * Placeholder building-block name (`w:placeholder/w:docPart@w:val`).
   * This is a reference to a glossary docPart that supplies the placeholder
   * content — NOT the literal placeholder text.
   */
  placeholder?: string;
  /** Whether the control is currently showing its placeholder (`w:showingPlcHdr`). */
  showingPlaceholder?: boolean;
  /** Date display format for date controls (`w:date/w:dateFormat@w:val`). */
  dateFormat?: string;
  /** Dropdown/combobox list items. */
  listItems?: { displayText: string; value: string }[];
  /** Checkbox checked state (`w14:checkbox`). */
  checked?: boolean;
  /** XML data binding (`w:dataBinding`), if the control is bound. */
  dataBinding?: SdtDataBinding;
  /**
   * The original `<w:sdtPr>` serialized verbatim as an XML string, captured
   * at parse time. Replayed unchanged on save so the properties block
   * round-trips losslessly. Stored as a string (not an `XmlElement`) so the
   * types layer stays free of the parser/`xml-js` dependency. Absent for
   * SDTs created programmatically — the serializer then synthesizes a
   * minimal, sequence-valid `w:sdtPr` from the modeled fields.
   */
  rawPropertiesXml?: string;
  /** The original `<w:sdtEndPr>` serialized verbatim, if present. */
  rawEndPropertiesXml?: string;
  /**
   * Present when this control is not a real `w:sdt` at all but a **legacy Word
   * form field** (`w:fldChar` + `w:ffData`) projected onto the SDT model so it
   * shares the content-control discovery/edit/render path. When set,
   * `rawPropertiesXml` is absent and serialization goes through
   * {@link LegacyFormField}'s captured run XML instead of `w:sdtPr`.
   */
  legacyFormField?: LegacyFormField;
}

/**
 * Inline SDT (content control within a paragraph)
 */
export interface InlineSdt {
  type: 'inlineSdt';
  /** SDT properties */
  properties: SdtProperties;
  /**
   * Inline content held inside the control. OOXML allows runs,
   * hyperlinks, simple/complex fields, nested SDTs, and math at this
   * level; the renderer must descend into all of them so docProps-bound
   * fields and similar template content survive paged rendering.
   */
  content: (Run | Hyperlink | SimpleField | ComplexField | InlineSdt | MathEquation)[];
}

/**
 * Block-level SDT (content control wrapping block content).
 *
 * `content` is `BlockContent[]` (not just paragraphs/tables) so a nested
 * block SDT survives the round trip. `CT_SdtContentBlock` also permits
 * run-level content (bookmarks, etc.); that is carried through the same
 * block-content parsing as elsewhere in the document.
 */
export interface BlockSdt {
  type: 'blockSdt';
  /** SDT properties */
  properties: SdtProperties;
  /** Block content inside the control */
  content: BlockContent[];
}

/**
 * The shape markup we keep verbatim, checked at both boundaries it crosses.
 *
 * `wps:bodyPr` and the unmodelled `wps:spPr` children (`a:ln`, `a:effectLst`)
 * are captured from the source and replayed on save, because the model holds
 * only a subset of them — see `ShapeTextBody.bodyPrXml`. Captured from a
 * `.docx` the shape of that markup is fixed by construction, but the same
 * strings ride `data-body-pr-xml` / `data-sp-pr-extra-xml` through the DOM,
 * so pasted HTML can supply anything. Well-formedness alone is not enough:
 * it would let arbitrary elements land inside `wps:spPr` on the next save.
 *
 * So the element name and its children are checked against the schema's own
 * lists before the string is stored or written. Below that first level,
 * {@link isWellFormedXmlElement} (no DTD, no PI, no CDATA, bounded depth,
 * bound prefixes) is the guarantee — the content there is DrawingML-shaped
 * either way, and what governs where an injection can land is the root.
 */

import {
  elementToXml,
  getChildElements,
  isWellFormedXmlElement,
  parseXml,
  type XmlElement,
} from './xmlParser';

/** `a:bodyPr` children — ECMA-376 CT_TextBodyProperties. */
const BODY_PR_CHILDREN = new Set([
  'a:prstTxWarp',
  'a:noAutofit',
  'a:normAutofit',
  'a:spAutoFit',
  'a:scene3d',
  'a:sp3d',
  'a:flatTx',
  'a:extLst',
]);

/** `a:ln` children — CT_LineProperties. */
const LINE_CHILDREN = new Set([
  'a:noFill',
  'a:solidFill',
  'a:gradFill',
  'a:pattFill',
  'a:prstDash',
  'a:custDash',
  'a:round',
  'a:bevel',
  'a:miter',
  'a:headEnd',
  'a:tailEnd',
  'a:extLst',
]);

/** `a:effectLst` children — CT_EffectList. */
const EFFECT_CHILDREN = new Set([
  'a:blur',
  'a:fillOverlay',
  'a:glow',
  'a:innerShdw',
  'a:outerShdw',
  'a:prstShdw',
  'a:reflection',
  'a:softEdge',
]);

export interface PreservedXmlOptions {
  /** Attributes to write over the source element's own. */
  overrides?: Record<string, string>;
  /**
   * Reject a prefix that neither the namespace table nor the fragment itself
   * binds. The paste boundary sets it; markup captured from a source part gets
   * its prefixes from the destination root instead (see `rootNamespaces`), so
   * requiring them here would drop a valid `a:extLst` extension on save.
   */
  requireBoundPrefixes?: boolean;
}

/**
 * The fragment's top-level elements, or null if it is not markup we will write.
 * A fragment can hold two siblings, so it is wrapped before validating — the
 * well-formedness check wants exactly one root.
 */
function topLevelElements(
  xml: string | null | undefined,
  options: PreservedXmlOptions
): XmlElement[] | null {
  if (typeof xml !== 'string' || xml.trim().length === 0) return null;
  const wrapped = `<epPreserved>${xml}</epPreserved>`;
  if (!isWellFormedXmlElement(wrapped, { requireBoundPrefixes: options.requireBoundPrefixes })) {
    return null;
  }

  const root = parseXml(wrapped).elements?.[0];
  if (!root) return null;
  // Character data around the root is already rejected; inside it is legal XML
  // but not something we replay. The parser keeps the whitespace between
  // elements, so a source part that was pretty-printed is not text content.
  if (hasCharacterData(root)) return null;
  return getChildElements(root);
}

/** A child node that is neither an element nor the whitespace between two. */
function hasCharacterData(el: XmlElement): boolean {
  return (el.elements ?? []).some(
    (node) => node.type !== 'element' && String(node.text ?? '').trim().length > 0
  );
}

/** Every child is an element the schema allows here. */
function childrenAllowed(el: XmlElement, allowed: Set<string>): boolean {
  if (hasCharacterData(el)) return false;
  for (const node of el.elements ?? []) {
    if (node.type !== 'element') continue;
    if (!allowed.has(node.name ?? '')) return false;
  }
  return true;
}

/**
 * The preserved `wps:bodyPr`, with `overrides` applied to its attributes.
 *
 * The model owns the handful of attributes it parses (the insets, the
 * anchor); the source owns every other attribute and every child it declared.
 * Letting the source win outright meant an edit to a modelled attribute was
 * silently dropped on export; rebuilding from the model alone lost the rest.
 *
 * @returns the element as XML, or undefined when there is nothing valid to replay.
 */
export function preservedBodyPrXml(
  xml: string | null | undefined,
  options: PreservedXmlOptions = {}
): string | undefined {
  const els = topLevelElements(xml, options);
  if (!els || els.length !== 1) return undefined;

  const el = els[0];
  if (el.name !== 'wps:bodyPr' || !childrenAllowed(el, BODY_PR_CHILDREN)) return undefined;

  const overrides = options.overrides ?? {};
  if (Object.keys(overrides).length === 0) return elementToXml(el);

  return elementToXml({ ...el, attributes: { ...(el.attributes ?? {}), ...overrides } });
}

/**
 * The preserved `wps:spPr` children, split so a model outline can replace the
 * line without taking the effects with it: the two are unrelated, and an
 * imported shape that had both lost its shadow as soon as anything set an
 * outline.
 *
 * An element that is neither drops the whole fragment rather than the one
 * sibling — from a `.docx` only these two are ever captured, so an unexpected
 * element means the string did not come from a parse.
 */
export function preservedSpPrExtra(
  xml: string | null | undefined,
  options: PreservedXmlOptions = {}
): { ln?: string; effectLst?: string } {
  const els = topLevelElements(xml, options);
  if (!els) return {};

  const out: { ln?: string; effectLst?: string } = {};
  for (const el of els) {
    if (el.name === 'a:ln' && !out.ln && childrenAllowed(el, LINE_CHILDREN)) {
      out.ln = elementToXml(el);
    } else if (
      el.name === 'a:effectLst' &&
      !out.effectLst &&
      childrenAllowed(el, EFFECT_CHILDREN)
    ) {
      out.effectLst = elementToXml(el);
    } else {
      return {};
    }
  }
  return out;
}

/** {@link preservedSpPrExtra} back as one string, in schema order. */
export function preservedSpPrExtraXml(
  xml: string | null | undefined,
  options: PreservedXmlOptions = {}
): string | undefined {
  const { ln, effectLst } = preservedSpPrExtra(xml, options);
  const joined = `${ln ?? ''}${effectLst ?? ''}`;
  return joined.length > 0 ? joined : undefined;
}

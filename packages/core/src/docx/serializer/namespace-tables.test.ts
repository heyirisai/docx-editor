/**
 * The namespace-table invariants, enforced instead of commented.
 *
 * `OOXML_NAMESPACE_URIS` is the one prefix→URI table. Every root the
 * serializers emit is a list of prefixes read out of it, and preserved markup
 * is re-declared from the same table — so a prefix a root emits with one URI
 * and a fragment binds with another is impossible by construction. These tests
 * exist so a future hand-written table gets caught at CI rather than by Word
 * refusing to open the file.
 */
import { describe, expect, test } from 'bun:test';
import { OOXML_NAMESPACE_URIS, elementToSelfContainedXml, parseXml } from '../xmlParser';
import {
  CAPTURING_ROOT_PREFIXES,
  FIXED_ROOT_PREFIXES,
  OOXML_NAMESPACES,
  buildRootNamespaces,
} from './rootNamespaces';
import { MC_IGNORABLE_PREFIXES, buildMcIgnorable } from './xmlUtils';

const allRootPrefixes = [...new Set([...CAPTURING_ROOT_PREFIXES, ...FIXED_ROOT_PREFIXES])];

describe('namespace tables', () => {
  test('every prefix a root emits resolves in the canonical table', () => {
    const unknown = allRootPrefixes.filter((p) => !(p in OOXML_NAMESPACE_URIS));
    expect(unknown).toEqual([]);
  });

  test('no root declares a prefix twice', () => {
    for (const list of [CAPTURING_ROOT_PREFIXES, FIXED_ROOT_PREFIXES]) {
      expect(new Set(list).size).toBe(list.length);
    }
  });

  test('a root never binds a prefix to a different URI than a fragment does', () => {
    // Both sides read `OOXML_NAMESPACE_URIS`, so this holds by construction —
    // the assertion is what makes reverting to a literal table fail.
    const { decl } = buildRootNamespaces(FIXED_ROOT_PREFIXES);
    for (const prefix of FIXED_ROOT_PREFIXES) {
      expect(decl).toContain(`xmlns:${prefix}="${OOXML_NAMESPACE_URIS[prefix]}"`);
    }
  });

  test('a preserved fragment re-declares from the same table the root uses', () => {
    // `a16:creationId` appears in nearly every Word drawing and is inherited
    // from the root rather than declared on the element.
    const el = parseXml('<w:drawing><a16:creationId id="1"/></w:drawing>').elements![0];
    expect(elementToSelfContainedXml(el)).toContain(`xmlns:a16="${OOXML_NAMESPACE_URIS.a16}"`);
  });

  test('mc:Ignorable only ever names prefixes the root declared', () => {
    const { prefixes } = buildRootNamespaces(CAPTURING_ROOT_PREFIXES);
    const ignorable = buildMcIgnorable(prefixes);
    const named = ignorable.match(/mc:Ignorable="([^"]*)"/)![1].split(' ');
    expect(named.filter((p) => !prefixes.includes(p))).toEqual([]);
    // ...and does not silently drop one it did declare.
    expect(named).toEqual(MC_IGNORABLE_PREFIXES.filter((p) => prefixes.includes(p)));
  });

  test('mc:Ignorable carries over what the source root listed', () => {
    const captured = { a16: 'urn:a16', w16du: 'urn:w16du' };
    const { prefixes } = buildRootNamespaces(CAPTURING_ROOT_PREFIXES, captured);
    const named = buildMcIgnorable(prefixes, ['a16', 'w16du', 'neverDeclared'])
      .match(/mc:Ignorable="([^"]*)"/)![1]
      .split(' ');
    expect(named).toContain('a16');
    expect(named).toContain('w16du');
    // Naming an undeclared prefix is itself invalid markup compatibility.
    expect(named).not.toContain('neverDeclared');
  });

  test('the fixed block is a superset of what a capturing root emits', () => {
    // `comments.xml`/`footnotes.xml` have no per-document capture, so anything
    // a header can emit has to already be declared there.
    const missing = CAPTURING_ROOT_PREFIXES.filter((p) => !FIXED_ROOT_PREFIXES.includes(p));
    expect(missing).toEqual([]);
    for (const prefix of FIXED_ROOT_PREFIXES) {
      expect(OOXML_NAMESPACES).toContain(`xmlns:${prefix}=`);
    }
  });
});

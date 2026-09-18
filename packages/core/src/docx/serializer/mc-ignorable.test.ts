/**
 * `mc:Ignorable` is the other half of the markup-compatibility contract.
 *
 * Declaring `xmlns:w16du` tells a consumer what the prefix means; listing
 * `w16du` in `mc:Ignorable` is what lets it SKIP an element in that namespace
 * instead of rejecting the part. This PR's preserved `mc:AlternateContent`
 * lands in headers, which emitted no `mc:Ignorable` at all.
 */
import { describe, expect, test } from 'bun:test';
import { serializeHeaderFooter } from './headerFooterSerializer';
import { serializeDocument } from './documentSerializer';
import { MC_IGNORABLE_PREFIXES } from './xmlUtils';
import type { Document, HeaderFooter } from '../../types/document';

const rootAttrs = (xml: string): string => {
  const start = xml.indexOf('<w:');
  return xml.slice(start, xml.indexOf('>', start) + 1);
};

const header = (over: Partial<HeaderFooter> = {}): HeaderFooter =>
  ({ type: 'header', rId: 'rId1', content: [], ...over }) as HeaderFooter;

const doc = (over: Record<string, unknown> = {}): Document =>
  ({
    package: { document: { content: [], ...over } },
  }) as unknown as Document;

describe('mc:Ignorable', () => {
  test('a header root carries one', () => {
    const xml = serializeHeaderFooter(header());
    expect(xml).toContain('mc:Ignorable="');
    for (const prefix of ['w14', 'w15', 'w16se', 'wp14']) {
      expect(xml).toMatch(new RegExp(`mc:Ignorable="[^"]*\\b${prefix}\\b`));
    }
  });

  test('a footer root carries one too', () => {
    expect(serializeHeaderFooter(header({ type: 'footer' }))).toContain('mc:Ignorable="');
  });

  test('the document root no longer drifts from the shared list', () => {
    const attrs = rootAttrs(serializeDocument(doc()));
    const named = attrs.match(/mc:Ignorable="([^"]*)"/)![1].split(' ');
    // Every prefix on the shared list that this root declares has to appear.
    for (const prefix of MC_IGNORABLE_PREFIXES) {
      if (attrs.includes(`xmlns:${prefix}=`)) expect(named).toContain(prefix);
    }
  });

  test('a prefix captured from the source root becomes ignorable with it', () => {
    const xml = serializeHeaderFooter(
      header({
        rootNamespaces: { a16: 'urn:a16', w16du: 'urn:w16du' },
        rootIgnorable: ['w14', 'a16', 'w16du'],
      })
    );
    const named = rootAttrs(xml)
      .match(/mc:Ignorable="([^"]*)"/)![1]
      .split(' ');
    expect(named).toContain('a16');
    expect(named).toContain('w16du');
    expect(xml).toContain('xmlns:a16="urn:a16"');
  });

  test('a captured prefix the root never declares is not named', () => {
    const named = rootAttrs(
      serializeHeaderFooter(header({ rootIgnorable: ['nobodyDeclaredThis'] }))
    )
      .match(/mc:Ignorable="([^"]*)"/)![1]
      .split(' ');
    expect(named).not.toContain('nobodyDeclaredThis');
  });
});

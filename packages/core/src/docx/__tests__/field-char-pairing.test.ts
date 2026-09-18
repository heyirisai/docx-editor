/**
 * Complex-field pairing across a paragraph.
 *
 * The walker keeps a depth counter and buffers runs between a `begin` and its
 * `end`. Anything that desyncs that counter routes the rest of the paragraph
 * into a buffer that is only flushed by a matching `end` — so the tail of the
 * document silently disappears. The unpaired-marker scan therefore reads the
 * SAME parsed items the walker does (object identity, not a shared index), the
 * depth never goes below zero, and a buffer left open at the end of the
 * paragraph is emitted rather than dropped.
 */
import { describe, expect, test } from 'bun:test';
import { parseXml } from '../xmlParser';
import { parseParagraphContents } from '../paragraphParser/content';
import type { ParagraphContent, Run } from '../../types/document';

function contents(inner: string): ParagraphContent[] {
  const el = parseXml(
    `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${inner}</w:p>`
  ).elements![0];
  return parseParagraphContents(el, null, null, null, null, null);
}

const fld = (t: string) => `<w:r><w:fldChar w:fldCharType="${t}"/></w:r>`;
const txt = (s: string) => `<w:r><w:t>${s}</w:t></w:r>`;
const instr = (s: string) => `<w:r><w:instrText xml:space="preserve">${s}</w:instrText></w:r>`;

/** Every `w:t` reachable from the parsed contents, in order. */
function allText(items: ParagraphContent[]): string {
  let out = '';
  const walkRun = (run: Run) => {
    for (const c of run.content) if (c.type === 'text') out += c.text;
  };
  for (const item of items) {
    if (item.type === 'run') walkRun(item);
    else if (item.type === 'complexField') {
      for (const r of item.fieldCode) walkRun(r);
      for (const r of item.fieldResult) walkRun(r);
    }
  }
  return out;
}

describe('complex field pairing', () => {
  test('models a self-contained field', () => {
    const items = contents(
      fld('begin') + instr(' PAGE ') + fld('separate') + txt('7') + fld('end') + txt('tail')
    );
    const field = items.find((i) => i.type === 'complexField');
    expect(field).toBeDefined();
    expect(field!.instruction).toBe('PAGE');
    expect(allText(items)).toContain('tail');
  });

  test('a nested field stays inside the outer one', () => {
    const items = contents(
      fld('begin') +
        instr(' TOC ') +
        fld('separate') +
        fld('begin') +
        instr(' PAGEREF _x ') +
        fld('separate') +
        txt('3') +
        fld('end') +
        fld('end') +
        txt('after')
    );
    expect(items.filter((i) => i.type === 'complexField')).toHaveLength(1);
    expect(allText(items)).toContain('after');
  });

  test('an unmatched begin passes through and keeps the rest of the paragraph', () => {
    // The TOC wrapper's `begin` lives in one paragraph and its `end` in another.
    const items = contents(fld('begin') + instr(' TOC \\o "1-3" ') + txt('entry one'));
    expect(items.every((i) => i.type !== 'complexField')).toBe(true);
    expect(allText(items)).toContain('entry one');
  });

  test('an unmatched end passes through and keeps the rest of the paragraph', () => {
    const items = contents(txt('before') + fld('end') + txt('after'));
    expect(allText(items)).toBe('beforeafter');
  });

  test('a stray end after a closed field does not swallow what follows', () => {
    // Depth is already 0 here — the second `end` must not drive it negative.
    const items = contents(
      fld('begin') +
        instr(' PAGE ') +
        fld('separate') +
        txt('7') +
        fld('end') +
        fld('end') +
        txt('tail')
    );
    expect(allText(items)).toContain('tail');
    expect(items.filter((i) => i.type === 'complexField')).toHaveLength(1);
  });

  test('a field split across runs keeps the text packed beside it', () => {
    // Word packs entry text and a whole field into one w:r.
    const items = contents(
      '<w:r><w:t>Chapter </w:t><w:fldChar w:fldCharType="begin"/>' +
        '<w:instrText xml:space="preserve"> PAGE </w:instrText>' +
        '<w:fldChar w:fldCharType="separate"/><w:t>12</w:t>' +
        '<w:fldChar w:fldCharType="end"/><w:t> end</w:t></w:r>'
    );
    const text = allText(items);
    expect(text).toContain('Chapter ');
    expect(text).toContain(' end');
  });
});

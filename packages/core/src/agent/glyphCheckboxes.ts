/**
 * Read-only detection of **glyph checkboxes** — the third way a Word document
 * says "tick this box", after `w:sdt` content controls and legacy form fields.
 *
 * Plenty of requirements matrices ship no interactive field at all: the author
 * typed a box character (`☐`), or inserted a Wingdings/Symbol glyph as a
 * `w:sym` run, and expects the responder to replace it by hand. There is no
 * structured state to edit, so this module only *finds* the candidates and
 * reports where they are — enough for staging to offer a column as answerable
 * and for a future converter to upgrade them to real controls.
 *
 * Nothing here mutates the document.
 */

import type {
  Document,
  DocumentBody,
  BlockContent,
  HeaderFooter,
  Paragraph,
  Run,
} from '../types/document';
import { getParagraphText } from './text-utils';
import type { ContentControlLocation } from './contentControls';

/**
 * Symbol fonts whose private-use code points Word renders as ballot boxes.
 * Matched case-insensitively against the `w:sym@w:font` value.
 */
const CHECKBOX_SYMBOL_FONTS = /^(wingdings( ?2)?|symbol|webdings)$/i;

/**
 * Private-use code points (as `w:sym@w:char` hex) that draw a box in those
 * fonts: Wingdings `F06F` empty / `F0FE` ticked / `F0FD` crossed, and the
 * Wingdings-2 `F0A3`/`F0A8` pair. Stored lowercase for lookup.
 */
const CHECKBOX_SYMBOL_CHARS: Record<string, boolean> = {
  f06f: false, // ❑ empty box
  f071: false, // ❑ empty box (Wingdings 2)
  f0a8: false, // ❑ empty box (Wingdings 2)
  f0fe: true, // ☑ ticked box
  f0fd: true, // ☒ crossed box
  f0a3: true, // ☑ ticked box (Wingdings 2)
};

/** Unicode ballot/box characters authored as ordinary `w:t` text. */
const CHECKBOX_TEXT_CHARS: Record<string, boolean> = {
  '☐': false, // ☐ BALLOT BOX
  '☑': true, // ☑ BALLOT BOX WITH CHECK
  '☒': true, // ☒ BALLOT BOX WITH X
  '□': false, // □ WHITE SQUARE
  '❑': false, // ❑ LOWER RIGHT SHADOWED WHITE SQUARE
};

/** A checkbox drawn as a character, with no structured state behind it. */
export interface GlyphCheckboxCandidate {
  /** Discriminator against `w:sdt` (`'sdt'`) and legacy (`'legacy'`) controls. */
  kind: 'glyph';
  /** How the glyph is authored: a `w:sym` run, or a box character in `w:t`. */
  encoding: 'sym' | 'text';
  /** Best-effort read of the drawn state (a ticked/crossed glyph = checked). */
  checked: boolean;
  /** The glyph itself — the character, or the `w:sym@w:char` hex for `sym`. */
  char: string;
  /** Symbol font (`w:sym@w:font`), when the glyph is a `w:sym` run. */
  font?: string;
  /** Block-index path of the enclosing paragraph (same scheme as `ContentControlInfo.path`). */
  path: number[];
  /** Index of the run within that paragraph's `content`. */
  runIndex: number;
  /** Plain text of the enclosing paragraph, as context for the caller. */
  paragraphText: string;
  /** Where the glyph lives (body vs a header/footer part). */
  location: ContentControlLocation;
}

/** Options for {@link findGlyphCheckboxes}. */
export interface FindGlyphCheckboxesOptions {
  /** Also scan header/footer parts (requires a full {@link Document}). */
  includeHeadersFooters?: boolean;
}

/** Narrow a {@link Document} or {@link DocumentBody} to its block list. */
function bodyOf(input: Document | DocumentBody): DocumentBody {
  return 'package' in input ? input.package.document : input;
}

/**
 * Classify a run as a checkbox glyph, or `null`. A text run only qualifies when
 * its *entire* text is one box character — a paragraph that merely mentions "☐"
 * mid-sentence is prose, not an answerable cell.
 */
function glyphOf(
  run: Run
): { encoding: 'sym' | 'text'; checked: boolean; char: string; font?: string } | null {
  const symbols = run.content.filter((c) => c.type === 'symbol');
  if (symbols.length === 1 && run.content.length === 1) {
    const sym = symbols[0];
    if (sym.type !== 'symbol') return null;
    const hex = sym.char.replace(/^0x/i, '').toLowerCase();
    if (CHECKBOX_SYMBOL_FONTS.test(sym.font) && hex in CHECKBOX_SYMBOL_CHARS) {
      return {
        encoding: 'sym',
        checked: CHECKBOX_SYMBOL_CHARS[hex],
        char: sym.char,
        font: sym.font,
      };
    }
    return null;
  }
  const text = run.content
    .map((c) => (c.type === 'text' ? c.text : ''))
    .join('')
    .trim();
  if (text.length === 1 && text in CHECKBOX_TEXT_CHARS) {
    return { encoding: 'text', checked: CHECKBOX_TEXT_CHARS[text], char: text };
  }
  return null;
}

/**
 * Find every checkbox-like glyph in the document, in document order.
 *
 * Walks body blocks, block SDTs, and table cells (row-major, nested tables
 * included); with `{ includeHeadersFooters: true }` and a full {@link Document}
 * it then walks header and footer parts, each sorted by relationship id, for
 * the same deterministic order `findContentControls` produces.
 *
 * Detection only — these have no `w:ffData`/`w:sdtPr` state to set, so there is
 * deliberately no matching setter.
 */
export function findGlyphCheckboxes(
  input: Document | DocumentBody,
  options: FindGlyphCheckboxesOptions = {}
): GlyphCheckboxCandidate[] {
  const out: GlyphCheckboxCandidate[] = [];

  const walkParagraph = (
    para: Paragraph,
    location: ContentControlLocation,
    path: number[]
  ): void => {
    para.content.forEach((node, runIndex) => {
      if (node.type !== 'run') return;
      const glyph = glyphOf(node);
      if (!glyph) return;
      out.push({
        kind: 'glyph',
        ...glyph,
        path,
        runIndex,
        paragraphText: getParagraphText(para),
        location,
      });
    });
  };

  const walkBlocks = (
    blocks: BlockContent[],
    location: ContentControlLocation,
    path: number[]
  ): void => {
    blocks.forEach((block, i) => {
      const blockPath = [...path, i];
      if (block.type === 'paragraph') walkParagraph(block, location, blockPath);
      else if (block.type === 'blockSdt') walkBlocks(block.content, location, blockPath);
      else if (block.type === 'table') {
        for (const row of block.rows) {
          for (const cell of row.cells) walkBlocks(cell.content, location, blockPath);
        }
      }
    });
  };

  walkBlocks(bodyOf(input).content, { part: 'body' }, []);

  if (options.includeHeadersFooters && 'package' in input) {
    const walkParts = (
      parts: Map<string, HeaderFooter> | undefined,
      part: 'header' | 'footer'
    ): void => {
      if (!parts) return;
      for (const rId of [...parts.keys()].sort()) {
        walkBlocks(parts.get(rId)!.content, { part, rId }, []);
      }
    };
    walkParts(input.package.headers, 'header');
    walkParts(input.package.footers, 'footer');
  }

  return out;
}

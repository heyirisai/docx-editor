/**
 * Document Serializer - Serialize complete document.xml
 *
 * Converts Document objects back to valid document.xml OOXML format.
 * Combines all content (paragraphs, tables) with section properties
 * and proper namespace declarations.
 *
 * OOXML Reference:
 * - Document root: w:document
 * - Document body: w:body
 * - Section properties: w:sectPr
 */

import type { Document, DocumentBody, BlockContent } from '../../types/document';

import { buildMcIgnorable } from './xmlUtils';
import { CAPTURING_ROOT_PREFIXES, buildRootNamespaces } from './rootNamespaces';
import { serializeParagraph } from './paragraphSerializer';
import { resetAutoIdCounter } from './runSerializer';
import { serializeTable } from './tableSerializer';
import { serializeBlockSdt } from './sdtSerializer';
import { serializeSectionProperties } from './sectionPropertiesSerializer';

// ============================================================================
// CONTENT SERIALIZATION
// ============================================================================

/**
 * Serialize a single block content item (paragraph, table, or block SDT).
 *
 * Exported so the footnote/endnote serializer can reuse the exact same
 * block-level emission the document body uses — preserving tracked-change
 * wrappers (`w:ins`/`w:del`), paragraph/run properties, fields, and tables
 * inside note bodies instead of reimplementing a minimal serializer.
 */
export function serializeBlockContent(block: BlockContent): string {
  if (block.type === 'paragraph') {
    return serializeParagraph(block);
  } else if (block.type === 'table') {
    return serializeTable(block);
  } else if (block.type === 'blockSdt') {
    return serializeBlockSdt(block, serializeBlockContent);
  }
  return '';
}

/**
 * Serialize document body content
 */
function serializeBodyContent(content: BlockContent[]): string {
  return content.map((block) => serializeBlockContent(block)).join('');
}

// ============================================================================
// MAIN DOCUMENT SERIALIZATION
// ============================================================================

/**
 * Serialize a DocumentBody to document.xml body content
 *
 * @param body - The document body to serialize
 * @returns XML string for the body element (without body tags)
 */
export function serializeDocumentBody(body: DocumentBody): string {
  const parts: string[] = [];

  // Serialize all content blocks
  parts.push(serializeBodyContent(body.content));

  // Final section properties (at the end of body)
  if (body.finalSectionProperties) {
    parts.push(serializeSectionProperties(body.finalSectionProperties));
  }

  return parts.join('');
}

/**
 * Serialize a complete Document to valid document.xml
 *
 * @param doc - The document to serialize
 * @returns Complete XML string for document.xml
 */
export function serializeDocument(doc: Document): string {
  // Reset auto-incrementing image/shape ID counter for this serialization pass
  resetAutoIdCounter();

  const parts: string[] = [];

  // XML declaration
  parts.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');

  // Document element with namespaces
  const ns = buildRootNamespaces(CAPTURING_ROOT_PREFIXES, doc.package.document.rootNamespaces);
  const ignorable = buildMcIgnorable(ns.prefixes, doc.package.document.rootIgnorable);
  parts.push(`<w:document ${ns.decl}${ignorable}>`);

  // Document body
  parts.push('<w:body>');
  parts.push(serializeDocumentBody(doc.package.document));
  parts.push('</w:body>');

  // Close document element
  parts.push('</w:document>');

  return parts.join('');
}

/**
 * Serialize just the document body (useful for partial updates)
 *
 * @param body - The document body to serialize
 * @returns XML string for the w:body element
 */
export function serializeDocumentBodyElement(body: DocumentBody): string {
  return `<w:body>${serializeDocumentBody(body)}</w:body>`;
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Check if document has any content
 */
export function hasDocumentContent(doc: Document): boolean {
  return doc.package.document.content.length > 0;
}

/**
 * Check if document has sections
 */
export function hasDocumentSections(doc: Document): boolean {
  return (doc.package.document.sections?.length ?? 0) > 0;
}

/**
 * Check if document has section properties
 */
export function hasSectionProperties(doc: Document): boolean {
  return doc.package.document.finalSectionProperties !== undefined;
}

/**
 * Get document content count (paragraphs + tables)
 */
export function getDocumentContentCount(doc: Document): number {
  return doc.package.document.content.length;
}

/**
 * Get paragraph count in document
 */
export function getDocumentParagraphCount(doc: Document): number {
  return doc.package.document.content.filter((b) => b.type === 'paragraph').length;
}

/**
 * Get table count in document
 */
export function getDocumentTableCount(doc: Document): number {
  return doc.package.document.content.filter((b) => b.type === 'table').length;
}

/**
 * Get plain text from document (for comparison/debugging)
 */
export function getDocumentPlainText(doc: Document): string {
  const texts: string[] = [];

  for (const block of doc.package.document.content) {
    if (block.type === 'paragraph') {
      for (const content of block.content) {
        if (content.type === 'run') {
          for (const item of content.content) {
            if (item.type === 'text') {
              texts.push(item.text);
            } else if (item.type === 'tab') {
              texts.push('\t');
            } else if (item.type === 'break') {
              texts.push('\n');
            }
          }
        }
      }
      texts.push('\n'); // Paragraph break
    }
  }

  return texts.join('');
}

/**
 * Create an empty document
 */
export function createEmptyDocument(): Document {
  return {
    package: {
      document: {
        content: [],
      },
    },
  };
}

/**
 * Create a simple document with text content
 */
export function createSimpleDocument(
  paragraphs: Array<{ text: string; styleId?: string }>
): Document {
  return {
    package: {
      document: {
        content: paragraphs.map((p) => ({
          type: 'paragraph' as const,
          formatting: p.styleId ? { styleId: p.styleId } : undefined,
          content: [
            {
              type: 'run' as const,
              content: [{ type: 'text' as const, text: p.text }],
            },
          ],
        })),
      },
    },
  };
}

export default serializeDocument;

import type { Node as ProseMirrorNode } from 'prosemirror-model';
import JSZip from 'jszip';
import { extractParagraphXml, findParagraphOffsets } from '../docx/selectiveXmlPatch';
import type { BlockContent } from '../types/content';
import { comparableJson } from '../utils/comparableJson';

export { comparableJson };

export type CollaborationExportFidelityFailure =
  | 'structural_change'
  | 'untracked_paragraph_change'
  | 'unsafe_paragraph_markup'
  | 'xml_patch_failed';

export class CollaborationExportFidelityError extends Error {
  readonly code = 'collaboration_export_fidelity_unsupported';
  readonly reason: CollaborationExportFidelityFailure;
  readonly paragraphId?: string;

  constructor(reason: CollaborationExportFidelityFailure, paragraphId?: string) {
    const paragraphSuffix = paragraphId ? ` in paragraph ${paragraphId}` : '';
    super(
      reason === 'structural_change'
        ? 'Collaboration export cannot safely merge structural changes into this complex template'
        : reason === 'untracked_paragraph_change'
          ? 'Collaboration export cannot safely identify a changed source paragraph'
          : reason === 'unsafe_paragraph_markup'
            ? `Collaboration export cannot safely replace unsupported source markup${paragraphSuffix}`
            : 'Collaboration export could not safely patch the source document XML'
    );
    this.name = 'CollaborationExportFidelityError';
    this.reason = reason;
    this.paragraphId = paragraphId;
  }
}

interface ParagraphChangeAnalysis {
  changedParagraphIds: Set<string>;
}

function valuesEqual(left: unknown, right: unknown): boolean {
  return JSON.stringify(comparableJson(left)) === JSON.stringify(comparableJson(right));
}

/**
 * The source XML patcher can replace paragraph bodies but cannot safely infer
 * how arbitrary block structure should be spliced into raw OOXML. Walk the
 * base/current PM trees in lockstep and accept content changes only inside an
 * existing, stable w14:paraId paragraph.
 */
export function analyzeCollaborationParagraphChanges(
  baseDocument: ProseMirrorNode,
  currentDocument: ProseMirrorNode
): ParagraphChangeAnalysis {
  const changedParagraphIds = new Set<string>();

  const visit = (baseNode: ProseMirrorNode, currentNode: ProseMirrorNode): void => {
    if (baseNode.type.name !== currentNode.type.name) {
      throw new CollaborationExportFidelityError('structural_change');
    }

    if (baseNode.type.name === 'paragraph') {
      if (valuesEqual(baseNode.toJSON(), currentNode.toJSON())) return;

      if (
        !valuesEqual(baseNode.attrs._sectionProperties, currentNode.attrs._sectionProperties) ||
        baseNode.attrs.sectionBreakType !== currentNode.attrs.sectionBreakType
      ) {
        throw new CollaborationExportFidelityError('structural_change');
      }

      const baseParagraphId =
        typeof baseNode.attrs.paraId === 'string' ? baseNode.attrs.paraId : '';
      const currentParagraphId =
        typeof currentNode.attrs.paraId === 'string' ? currentNode.attrs.paraId : '';
      if (!baseParagraphId || baseParagraphId !== currentParagraphId) {
        throw new CollaborationExportFidelityError('untracked_paragraph_change');
      }
      if (changedParagraphIds.has(baseParagraphId)) {
        throw new CollaborationExportFidelityError('untracked_paragraph_change', baseParagraphId);
      }
      changedParagraphIds.add(baseParagraphId);
      return;
    }

    if (
      !valuesEqual(baseNode.attrs, currentNode.attrs) ||
      !valuesEqual(baseNode.marks, currentNode.marks) ||
      baseNode.text !== currentNode.text ||
      baseNode.childCount !== currentNode.childCount
    ) {
      throw new CollaborationExportFidelityError('structural_change');
    }

    for (let index = 0; index < baseNode.childCount; index++) {
      visit(baseNode.child(index), currentNode.child(index));
    }
  };

  visit(baseDocument, currentDocument);
  return { changedParagraphIds };
}

interface CommentRangeSignatures {
  byParagraphId: Map<string, string>;
  untracked: string[];
}

function collectCommentRangeSignatures(blocks: readonly BlockContent[]): CommentRangeSignatures {
  const byParagraphId = new Map<string, string>();
  const untracked: string[] = [];

  const visit = (content: readonly BlockContent[]): void => {
    for (const block of content) {
      if (block.type === 'paragraph') {
        const signature = block.content
          .filter((item) => item.type === 'commentRangeStart' || item.type === 'commentRangeEnd')
          .map((item) => `${item.type}:${item.id}`)
          .join('|');
        if (block.paraId) {
          byParagraphId.set(block.paraId, signature);
        } else if (signature) {
          untracked.push(signature);
        }
      } else if (block.type === 'table') {
        for (const row of block.rows) {
          for (const cell of row.cells) visit(cell.content);
        }
      } else {
        visit(block.content);
      }
    }
  };

  visit(blocks);
  return { byParagraphId, untracked };
}

/**
 * Reply threads share their parent's visible PM mark, so adding or deleting a
 * reply can change OOXML range markers without changing the PM document. Add
 * those stable source paragraphs to the selective patch set.
 */
export function includeCollaborationCommentRangeChanges(input: {
  baseContent: readonly BlockContent[];
  currentContent: readonly BlockContent[];
  changedParagraphIds: Set<string>;
}): void {
  const base = collectCommentRangeSignatures(input.baseContent);
  const current = collectCommentRangeSignatures(input.currentContent);
  if (JSON.stringify(base.untracked) !== JSON.stringify(current.untracked)) {
    throw new CollaborationExportFidelityError('untracked_paragraph_change');
  }

  const paragraphIds = new Set([...base.byParagraphId.keys(), ...current.byParagraphId.keys()]);
  for (const paragraphId of paragraphIds) {
    if (base.byParagraphId.get(paragraphId) !== current.byParagraphId.get(paragraphId)) {
      input.changedParagraphIds.add(paragraphId);
    }
  }
}

const UNSAFE_PARAGRAPH_MARKUP =
  /<(?:w:(?:pict|object|control|altChunk|customXml|smartTag|subDoc)|mc:AlternateContent|v:[A-Za-z][\w.-]*|o:OLEObject|wps:[A-Za-z][\w.-]*|wpg:[A-Za-z][\w.-]*)\b/i;

function assertChangedParagraphsAreSafe(
  originalDocumentXml: string,
  changedParagraphIds: ReadonlySet<string>
): void {
  for (const paragraphId of changedParagraphIds) {
    const originalParagraphXml = extractParagraphXml(originalDocumentXml, paragraphId);
    if (!originalParagraphXml) {
      throw new CollaborationExportFidelityError('xml_patch_failed', paragraphId);
    }
    if (UNSAFE_PARAGRAPH_MARKUP.test(originalParagraphXml)) {
      throw new CollaborationExportFidelityError('unsafe_paragraph_markup', paragraphId);
    }
  }
}

function patchChangedParagraphsIntoSource(
  sourceDocumentXml: string,
  exportDocumentXml: string,
  changedParagraphIds: ReadonlySet<string>
): string | null {
  const replacements: Array<{ start: number; end: number; xml: string }> = [];
  for (const paragraphId of changedParagraphIds) {
    const sourceOffsets = findParagraphOffsets(sourceDocumentXml, paragraphId);
    const replacementXml = extractParagraphXml(exportDocumentXml, paragraphId);
    if (!sourceOffsets || !replacementXml) return null;
    replacements.push({ ...sourceOffsets, xml: replacementXml });
  }

  replacements.sort((left, right) => right.start - left.start);
  let result = sourceDocumentXml;
  for (const replacement of replacements) {
    result = result.slice(0, replacement.start) + replacement.xml + result.slice(replacement.end);
  }
  return result;
}

const HEADER_FOOTER_PART =
  /^word\/(?:(?:header|footer)\d*\.xml|_rels\/(?:header|footer)\d*\.xml\.rels)$/;
const COMMENT_PART = /^word\/comments(?:Extended|Ids|Extensible)?\.xml$/i;
const COMMENT_PART_NAME = /\/word\/comments(?:Extended|Ids|Extensible)?\.xml/i;
const COMMENT_RELATIONSHIP =
  /(?:Target="(?:\/?word\/)?comments(?:Extended|Ids|Extensible)?\.xml"|Type="[^"]*\/comments(?:Extended|Ids|Extensible)?")/i;

async function restoreUntouchedHeaderFooterParts(
  sourceZip: JSZip,
  exportZip: JSZip
): Promise<void> {
  for (const path of Object.keys(exportZip.files)) {
    if (HEADER_FOOTER_PART.test(path) && !sourceZip.file(path)) {
      exportZip.remove(path);
    }
  }

  for (const [path, file] of Object.entries(sourceZip.files)) {
    if (file.dir || !HEADER_FOOTER_PART.test(path)) continue;
    exportZip.file(path, await file.async('arraybuffer'), {
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
    });
  }
}

function restoreMatchingXmlElements(input: {
  sourceXml: string;
  exportXml: string;
  elementName: 'Override' | 'Relationship';
  matches: (element: string) => boolean;
  closingTag: '</Types>' | '</Relationships>';
}): string {
  const elementPattern = new RegExp(`<${input.elementName}\\b[^>]*/>`, 'gi');
  const sourceElements = input.sourceXml.match(elementPattern)?.filter(input.matches) ?? [];
  const withoutExportElements = input.exportXml.replace(elementPattern, (element) =>
    input.matches(element) ? '' : element
  );
  return withoutExportElements.replace(
    input.closingTag,
    `${sourceElements.join('')}${input.closingTag}`
  );
}

async function restoreUntouchedCommentParts(sourceZip: JSZip, exportZip: JSZip): Promise<void> {
  for (const path of Object.keys(exportZip.files)) {
    if (COMMENT_PART.test(path) && !sourceZip.file(path)) exportZip.remove(path);
  }
  for (const [path, file] of Object.entries(sourceZip.files)) {
    if (file.dir || !COMMENT_PART.test(path)) continue;
    exportZip.file(path, await file.async('arraybuffer'), {
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
    });
  }

  const contentTypesPath = '[Content_Types].xml';
  const sourceContentTypes = sourceZip.file(contentTypesPath);
  const exportContentTypes = exportZip.file(contentTypesPath);
  if (sourceContentTypes && exportContentTypes) {
    const [sourceXml, exportXml] = await Promise.all([
      sourceContentTypes.async('text'),
      exportContentTypes.async('text'),
    ]);
    exportZip.file(
      contentTypesPath,
      restoreMatchingXmlElements({
        sourceXml,
        exportXml,
        elementName: 'Override',
        matches: (element) => COMMENT_PART_NAME.test(element),
        closingTag: '</Types>',
      }),
      { compression: 'DEFLATE', compressionOptions: { level: 6 } }
    );
  }

  const relationshipsPath = 'word/_rels/document.xml.rels';
  const sourceRelationships = sourceZip.file(relationshipsPath);
  const exportRelationships = exportZip.file(relationshipsPath);
  if (sourceRelationships && exportRelationships) {
    const [sourceXml, exportXml] = await Promise.all([
      sourceRelationships.async('text'),
      exportRelationships.async('text'),
    ]);
    exportZip.file(
      relationshipsPath,
      restoreMatchingXmlElements({
        sourceXml,
        exportXml,
        elementName: 'Relationship',
        matches: (element) => COMMENT_RELATIONSHIP.test(element),
        closingTag: '</Relationships>',
      }),
      { compression: 'DEFLATE', compressionOptions: { level: 6 } }
    );
  }
}

/**
 * Use a fully repacked package for relationship/media/comment updates, then
 * replace its lossy document.xml with a paragraph-selective patch of the
 * immutable source XML. Source-only VML and other unsupported constructs in
 * untouched paragraphs remain byte-for-byte identical.
 */
export async function preserveSourceXmlAroundParagraphChanges(input: {
  sourceBuffer: ArrayBuffer;
  fullyRepackedBuffer: ArrayBuffer;
  changedParagraphIds: ReadonlySet<string>;
  preserveCommentParts?: boolean;
}): Promise<ArrayBuffer> {
  const [sourceZip, exportZip] = await Promise.all([
    JSZip.loadAsync(input.sourceBuffer),
    JSZip.loadAsync(input.fullyRepackedBuffer),
  ]);
  const sourceDocumentFile = sourceZip.file('word/document.xml');
  const exportDocumentFile = exportZip.file('word/document.xml');
  if (!sourceDocumentFile || !exportDocumentFile) {
    throw new CollaborationExportFidelityError('xml_patch_failed');
  }

  const [sourceDocumentXml, exportDocumentXml] = await Promise.all([
    sourceDocumentFile.async('text'),
    exportDocumentFile.async('text'),
  ]);
  assertChangedParagraphsAreSafe(sourceDocumentXml, input.changedParagraphIds);
  const patchedDocumentXml = patchChangedParagraphsIntoSource(
    sourceDocumentXml,
    exportDocumentXml,
    input.changedParagraphIds
  );
  if (!patchedDocumentXml) {
    throw new CollaborationExportFidelityError('xml_patch_failed');
  }

  exportZip.file('word/document.xml', patchedDocumentXml, {
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });
  await restoreUntouchedHeaderFooterParts(sourceZip, exportZip);
  if (input.preserveCommentParts) {
    await restoreUntouchedCommentParts(sourceZip, exportZip);
  }
  return exportZip.generateAsync({
    type: 'arraybuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });
}

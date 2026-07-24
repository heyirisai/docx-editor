import { IncompleteExternalMediaManifestError, parseDocx } from '../docx/parser';
import { repackDocx } from '../docx/rezip';
import {
  injectReplyRangeMarkers,
  injectTCReplyRangeMarkers,
} from '../docx/injectReplyRangeMarkers';
import { removeOrphanCommentRanges } from '../docx/commentRangeIntegrity';
import { schema } from '../prosemirror/schema';
import { toProseDoc } from '../prosemirror/conversion/toProseDoc';
import { fromProseDoc } from '../prosemirror/conversion/fromProseDoc';
import {
  collaborationProjectionIdentityMatches,
  rehydrateCollaborationDocument,
} from './fidelitySidecar';
import { projectProseMirrorDocument } from './projection';
import {
  analyzeCollaborationParagraphChanges,
  includeCollaborationCommentRangeChanges,
  preserveSourceXmlAroundParagraphChanges,
} from './sourcePreservingExport';
export { CollaborationExportFidelityError } from './sourcePreservingExport';
import type { ExternalMediaManifestEntry } from '../types/document';
import type { Comment } from '../types/content';
import type {
  CollaborationDiagnostic,
  CollaborationExportAsset,
  CollaborationExportInput,
  ParsedCollaborationPackage,
} from './types';

export class MissingCollaborationAssetError extends Error {
  readonly code = 'missing_collaboration_asset';
  readonly assetId: string;

  constructor(assetId: string) {
    super(`Missing original bytes for collaboration asset: ${assetId}`);
    this.name = 'MissingCollaborationAssetError';
    this.assetId = assetId;
  }
}

export class CollaborationIdentityMismatchError extends Error {
  readonly code = 'collaboration_identity_mismatch';

  constructor() {
    super('Collaboration projection does not match the immutable base revision');
    this.name = 'CollaborationIdentityMismatchError';
  }
}

export class IncompleteCollaborationManifestError extends Error {
  readonly code = 'incomplete_collaboration_manifest';
  readonly missingPaths: readonly string[];

  constructor(missingPaths: readonly string[] = []) {
    super('Collaboration package manifest is incomplete');
    this.name = 'IncompleteCollaborationManifestError';
    this.missingPaths = [...missingPaths];
  }
}

async function parseCollaborationSource(
  packageBuffer: ArrayBuffer,
  manifest: readonly ExternalMediaManifestEntry[]
) {
  try {
    return await parseDocx(packageBuffer, {
      externalMedia: { entries: manifest },
    });
  } catch (error) {
    if (error instanceof IncompleteExternalMediaManifestError) {
      throw new IncompleteCollaborationManifestError(error.missingPaths);
    }
    throw error;
  }
}

export async function parseCollaborationPackage(input: {
  packageBuffer: ArrayBuffer;
  manifest: readonly ExternalMediaManifestEntry[];
  baseRevision: string;
}): Promise<ParsedCollaborationPackage> {
  const document = await parseCollaborationSource(input.packageBuffer, input.manifest);
  const proseMirrorDocument = toProseDoc(document, {
    styles: document.package.styles,
  });
  const projection = projectProseMirrorDocument(proseMirrorDocument, {
    baseRevision: input.baseRevision,
  });
  if (projection.diagnostics.some((diagnostic) => diagnostic.severity === 'error')) {
    throw new IncompleteCollaborationManifestError();
  }
  return { document, projection };
}

function encodeBase64(bytes: Uint8Array): string {
  let binary = '';
  const chunkSize = 0x8000;
  for (let offset = 0; offset < bytes.byteLength; offset += chunkSize) {
    const chunk = bytes.subarray(offset, Math.min(offset + chunkSize, bytes.byteLength));
    binary += String.fromCharCode(...chunk);
  }
  return btoa(binary);
}

async function sha256Hex(buffer: ArrayBuffer): Promise<string> {
  const digest = await crypto.subtle.digest('SHA-256', buffer);
  return [...new Uint8Array(digest)].map((byte) => byte.toString(16).padStart(2, '0')).join('');
}

function commentsEqual(left: readonly Comment[], right: readonly Comment[]): boolean {
  const byId = (comments: readonly Comment[]) =>
    [...comments].sort((first, second) => first.id - second.id);
  return JSON.stringify(byId(left)) === JSON.stringify(byId(right));
}

function materializeNewAssets(
  value: unknown,
  assetsById: ReadonlyMap<string, CollaborationExportAsset>,
  visited: Set<object>
): void {
  if (!value || typeof value !== 'object' || visited.has(value)) return;
  if (value instanceof ArrayBuffer || ArrayBuffer.isView(value) || value instanceof Date) return;
  visited.add(value);

  if (value instanceof Map) {
    for (const child of value.values()) {
      materializeNewAssets(child, assetsById, visited);
    }
    return;
  }
  if (Array.isArray(value)) {
    for (const child of value) {
      materializeNewAssets(child, assetsById, visited);
    }
    return;
  }

  const record = value as Record<string, unknown>;
  if (record.type === 'image' && typeof record.assetId === 'string' && !record.rId) {
    const asset = assetsById.get(record.assetId);
    if (!asset) {
      throw new MissingCollaborationAssetError(record.assetId);
    }
    record.src = `data:${asset.mimeType};base64,${encodeBase64(asset.bytes)}`;
    record.mimeType = asset.mimeType;
  }

  for (const child of Object.values(record)) {
    materializeNewAssets(child, assetsById, visited);
  }
}

/**
 * Worker-safe full-fidelity export. The immutable source package is loaded
 * only inside the caller's document worker, then discarded with that worker.
 */
export async function exportCollaborationDocument(
  input: CollaborationExportInput
): Promise<ArrayBuffer> {
  if (
    input.sidecar.baseRevision !== input.baseRevision ||
    (await sha256Hex(input.sourceBuffer)) !== input.sourceSha256.toLowerCase()
  ) {
    throw new CollaborationIdentityMismatchError();
  }
  const baseDocument = await parseCollaborationSource(input.sourceBuffer, input.manifest);
  const baseProseMirrorDocument = toProseDoc(baseDocument, {
    styles: baseDocument.package.styles,
  });
  const baseProjection = projectProseMirrorDocument(baseProseMirrorDocument, {
    baseRevision: input.baseRevision,
  });
  if (
    !collaborationProjectionIdentityMatches({
      baseDocument: baseProjection.document,
      currentDocument: input.projectedDocument,
      canonicalSidecar: baseProjection.sidecar,
      suppliedSidecar: input.sidecar,
    })
  ) {
    throw new CollaborationIdentityMismatchError();
  }
  const hydratedBaseJson = rehydrateCollaborationDocument(
    baseProjection.document,
    baseProjection.sidecar
  );
  const hydratedBaseDocument = schema.nodeFromJSON(hydratedBaseJson);
  const hydrationDiagnostics: CollaborationDiagnostic[] = [];
  const hydratedJson = rehydrateCollaborationDocument(
    input.projectedDocument,
    baseProjection.sidecar,
    hydrationDiagnostics
  );
  if (hydrationDiagnostics.some((diagnostic) => diagnostic.severity === 'error')) {
    throw new CollaborationIdentityMismatchError();
  }
  const proseMirrorDocument = schema.nodeFromJSON(hydratedJson);
  const { changedParagraphIds } = analyzeCollaborationParagraphChanges(
    hydratedBaseDocument,
    proseMirrorDocument
  );
  const baseComments = baseDocument.package.document.comments ?? [];
  const commentsChanged =
    input.comments !== undefined && !commentsEqual(baseComments, input.comments);
  if (changedParagraphIds.size === 0 && !commentsChanged) {
    return input.sourceBuffer.slice(0);
  }

  const exportDocument = fromProseDoc(proseMirrorDocument, baseDocument);
  if (input.comments !== undefined) {
    const comments = input.comments.map((comment) => structuredClone(comment));
    exportDocument.package.document.comments = comments;
    removeOrphanCommentRanges(exportDocument);
    injectReplyRangeMarkers(exportDocument.package.document.content, comments);
    injectTCReplyRangeMarkers(exportDocument.package.document.content, comments);
    if (commentsChanged) {
      includeCollaborationCommentRangeChanges({
        baseContent: baseDocument.package.document.content,
        currentContent: exportDocument.package.document.content,
        changedParagraphIds,
      });
    }
  }
  const assetsById = new Map((input.assets ?? []).map((asset) => [asset.assetId, asset] as const));
  materializeNewAssets(exportDocument.package, assetsById, new Set());
  const fullyRepackedBuffer = await repackDocx(exportDocument, {
    preserveUnreferencedMedia: true,
  });
  return preserveSourceXmlAroundParagraphChanges({
    sourceBuffer: input.sourceBuffer,
    fullyRepackedBuffer,
    changedParagraphIds,
    preserveCommentParts: !commentsChanged,
  });
}

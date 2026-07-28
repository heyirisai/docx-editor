import type { Node as ProseMirrorNode } from 'prosemirror-model';

export interface CollaborationJsonMark {
  type: string;
  attrs?: Record<string, unknown>;
}

export interface CollaborationJsonNode {
  type: string;
  attrs?: Record<string, unknown>;
  content?: CollaborationJsonNode[];
  marks?: CollaborationJsonMark[];
  text?: string;
}

export interface CollaborationProjectionOptions {
  baseRevision: string;
  schemaVersion?: number;
}

export interface CollaborationDiagnostic {
  code: 'image_asset_id_missing' | 'sidecar_entry_mismatch';
  severity: 'error' | 'warning';
  nodeType: string;
}

export interface CollaborationProjectionStats {
  sourceJsonBytes: number;
  projectedJsonBytes: number;
  sidecarJsonBytes: number;
  mediaReferenceCount: number;
  strippedAttributeCount: number;
}

export interface CollaborationProjection {
  schemaVersion: number;
  baseRevision: string;
  document: CollaborationJsonNode;
  sidecar: FidelitySidecar;
  diagnostics: CollaborationDiagnostic[];
  stats: CollaborationProjectionStats;
}

export interface FidelitySidecarEntry {
  nodeType: string;
  attrs: Record<string, unknown>;
}

export interface FidelitySidecar {
  version: 1;
  baseRevision: string;
  entries: Record<string, FidelitySidecarEntry>;
}

export interface ParsedCollaborationPackage {
  document: import('../types/document').Document;
  projection: CollaborationProjection;
}

export interface CollaborationExportAsset {
  assetId: string;
  mimeType: string;
  bytes: Uint8Array;
}

export interface CollaborationExportInput {
  sourceBuffer: ArrayBuffer;
  sourceSha256: string;
  baseRevision: string;
  manifest: readonly import('../types/document').ExternalMediaManifestEntry[];
  projectedDocument: CollaborationJsonNode;
  sidecar: FidelitySidecar;
  /**
   * Current controlled comment threads. Omit only when the caller does not
   * manage comments; an explicit empty array removes all comments.
   */
  comments?: readonly import('../types/content').Comment[];
  assets?: readonly CollaborationExportAsset[];
}

export interface ProseMirrorProjectionInput {
  document: ProseMirrorNode;
  options: CollaborationProjectionOptions;
}

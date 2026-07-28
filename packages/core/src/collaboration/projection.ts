import type { Node as ProseMirrorNode } from 'prosemirror-model';
import type {
  CollaborationDiagnostic,
  CollaborationJsonNode,
  CollaborationProjection,
  CollaborationProjectionOptions,
  FidelitySidecar,
  FidelitySidecarEntry,
} from './types';
import { getCollaborationNodeHint } from './fidelitySidecar';
export { rehydrateCollaborationDocument } from './fidelitySidecar';

const DEFAULT_SCHEMA_VERSION = 2;
const PRIVATE_ATTRS_BY_NODE: Readonly<Record<string, readonly string[]>> = {
  paragraph: ['_originalFormatting', '_originalRunBoundaries'],
  table: ['_originalFormatting'],
  tableRow: ['_originalFormatting'],
  tableCell: ['_originalFormatting', '_originalResolvedFill'],
  tableHeader: ['_originalFormatting', '_originalResolvedFill'],
  blockSdt: ['rawPropertiesXml', 'rawEndPropertiesXml'],
  image: ['rId'],
};
const PROVENANCE_NODE_TYPES = new Set([
  'paragraph',
  'table',
  'tableRow',
  'tableCell',
  'tableHeader',
  'blockSdt',
  'image',
  'textBox',
]);
const FNV_128_OFFSET_BASIS = 0x6c62272e07bb014262b821756295c58dn;
const FNV_128_PRIME = 0x0000000001000000000000000000013bn;
const FNV_128_MASK = (1n << 128n) - 1n;

function byteLength(value: unknown): number {
  return new TextEncoder().encode(JSON.stringify(value)).byteLength;
}

function hashString(value: string): string {
  let hash = FNV_128_OFFSET_BASIS;
  for (const byte of new TextEncoder().encode(value)) {
    hash ^= BigInt(byte);
    hash = (hash * FNV_128_PRIME) & FNV_128_MASK;
  }
  return hash.toString(16).padStart(32, '0');
}

function cloneAttributeValue(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(cloneAttributeValue);
  if (value && typeof value === 'object') {
    const clone: Record<string, unknown> = {};
    for (const [key, child] of Object.entries(value)) {
      clone[key] = cloneAttributeValue(child);
    }
    return clone;
  }
  return value;
}

function cloneAttrs(attrs: Record<string, unknown> | undefined): Record<string, unknown> {
  const clone: Record<string, unknown> = {};
  if (!attrs) return clone;
  for (const [key, value] of Object.entries(attrs)) {
    clone[key] = cloneAttributeValue(value);
  }
  return clone;
}

export function projectProseMirrorDocument(
  document: ProseMirrorNode,
  options: CollaborationProjectionOptions
): CollaborationProjection {
  const source = document.toJSON() as CollaborationJsonNode;
  const sidecar: FidelitySidecar = {
    version: 1,
    baseRevision: options.baseRevision,
    entries: {},
  };
  const diagnostics: CollaborationDiagnostic[] = [];
  const collaborationIds = new Set<string>();
  let mediaReferenceCount = 0;
  let strippedAttributeCount = 0;

  const visit = (node: CollaborationJsonNode, path: readonly number[]): CollaborationJsonNode => {
    const attrs = cloneAttrs(node.attrs);
    const fidelityAttrs: Record<string, unknown> = {};
    const privateKeys = PRIVATE_ATTRS_BY_NODE[node.type] ?? [];

    for (const key of privateKeys) {
      if (attrs[key] === null || attrs[key] === undefined) continue;
      fidelityAttrs[key] = attrs[key];
      delete attrs[key];
      strippedAttributeCount++;
    }

    if (node.type === 'image') {
      mediaReferenceCount++;
      const assetId = typeof attrs.assetId === 'string' ? attrs.assetId : '';
      const sourceValue = typeof attrs.src === 'string' ? attrs.src : '';
      if (!assetId) {
        diagnostics.push({
          code: 'image_asset_id_missing',
          severity: 'error',
          nodeType: node.type,
        });
      }
      if (sourceValue) {
        attrs.src = '';
        strippedAttributeCount++;
      }
    }

    if (PROVENANCE_NODE_TYPES.has(node.type)) {
      const existingId =
        typeof attrs.collaborationId === 'string' ? attrs.collaborationId : undefined;
      const collaborationId =
        existingId ??
        `c-${hashString(
          `${options.baseRevision}|${node.type}|${getCollaborationNodeHint(node)}|${path.join('.')}`
        )}`;
      if (collaborationIds.has(collaborationId)) {
        throw new Error(`Duplicate collaboration ID: ${collaborationId}`);
      }
      collaborationIds.add(collaborationId);
      attrs.collaborationId = collaborationId;

      if (Object.keys(fidelityAttrs).length > 0) {
        const entry: FidelitySidecarEntry = {
          nodeType: node.type,
          attrs: fidelityAttrs,
        };
        sidecar.entries[collaborationId] = entry;
      }
    }

    return {
      ...node,
      attrs: Object.keys(attrs).length > 0 ? attrs : undefined,
      content: node.content?.map((child, index) => visit(child, [...path, index])),
      marks: node.marks?.map((mark) => ({
        ...mark,
        attrs: mark.attrs ? cloneAttrs(mark.attrs) : undefined,
      })),
    };
  };

  const projectedDocument = visit(source, []);
  return {
    schemaVersion: options.schemaVersion ?? DEFAULT_SCHEMA_VERSION,
    baseRevision: options.baseRevision,
    document: projectedDocument,
    sidecar,
    diagnostics,
    stats: {
      sourceJsonBytes: byteLength(source),
      projectedJsonBytes: byteLength(projectedDocument),
      sidecarJsonBytes: byteLength(sidecar),
      mediaReferenceCount,
      strippedAttributeCount,
    },
  };
}

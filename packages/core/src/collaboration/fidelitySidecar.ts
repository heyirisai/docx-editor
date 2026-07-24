import type { CollaborationDiagnostic, CollaborationJsonNode, FidelitySidecar } from './types';

interface CollaborationNodeIdentity {
  nodeType: string;
  path: string;
  stableHint: string;
}

function readCollaborationId(node: CollaborationJsonNode): {
  id?: string;
  valid: boolean;
} {
  const value = node.attrs?.collaborationId;
  if (value === null || value === undefined) return { valid: true };
  if (typeof value !== 'string' || value.length === 0) return { valid: false };
  return { id: value, valid: true };
}

export function getCollaborationNodeHint(node: CollaborationJsonNode): string {
  const attrs = node.attrs;
  for (const key of ['paraId', 'textId', 'textBoxId', 'assetId', 'id', 'tag']) {
    const value = attrs?.[key];
    if (typeof value === 'string' || typeof value === 'number') {
      return `${key}:${String(value)}`;
    }
  }
  return '';
}

function jsonValuesEqual(
  left: unknown,
  right: unknown,
  comparedObjects = new WeakMap<object, object>()
): boolean {
  if (Object.is(left, right)) return true;
  if (!left || !right || typeof left !== 'object' || typeof right !== 'object') return false;

  const previousRight = comparedObjects.get(left);
  if (previousRight) return previousRight === right;
  comparedObjects.set(left, right);

  if (Array.isArray(left) || Array.isArray(right)) {
    if (!Array.isArray(left) || !Array.isArray(right) || left.length !== right.length) {
      return false;
    }
    return left.every((value, index) => jsonValuesEqual(value, right[index], comparedObjects));
  }

  const leftRecord = left as Record<string, unknown>;
  const rightRecord = right as Record<string, unknown>;
  const leftKeys = Object.keys(leftRecord).sort();
  const rightKeys = Object.keys(rightRecord).sort();
  if (
    leftKeys.length !== rightKeys.length ||
    leftKeys.some((key, index) => key !== rightKeys[index])
  ) {
    return false;
  }
  return leftKeys.every((key) =>
    jsonValuesEqual(leftRecord[key], rightRecord[key], comparedObjects)
  );
}

function collectCollaborationIdentities(
  document: CollaborationJsonNode
): Map<string, CollaborationNodeIdentity> | null {
  const identities = new Map<string, CollaborationNodeIdentity>();
  let valid = true;

  const visit = (node: CollaborationJsonNode, path: readonly number[]): void => {
    const collaborationId = readCollaborationId(node);
    if (!collaborationId.valid) {
      valid = false;
      return;
    }
    if (collaborationId.id) {
      if (identities.has(collaborationId.id)) {
        valid = false;
        return;
      }
      identities.set(collaborationId.id, {
        nodeType: node.type,
        path: path.join('.'),
        stableHint: getCollaborationNodeHint(node),
      });
    }
    node.content?.forEach((child, index) => visit(child, [...path, index]));
  };

  visit(document, []);
  return valid ? identities : null;
}

function alignedIdentitiesMatch(
  baseNode: CollaborationJsonNode,
  currentNode: CollaborationJsonNode
): boolean {
  if (baseNode.type !== currentNode.type) {
    // Structural validation owns type changes. There is no aligned identity to
    // compare once the structures diverge.
    return true;
  }

  const baseIdentity = readCollaborationId(baseNode);
  const currentIdentity = readCollaborationId(currentNode);
  if (!baseIdentity.valid || !currentIdentity.valid) return false;
  if (baseIdentity.id && baseIdentity.id !== currentIdentity.id) return false;

  // Paragraph contents are the collaboration edit boundary, so identities
  // below a paragraph are matched by their stable hints instead of position.
  if (baseNode.type === 'paragraph') return true;
  const baseContent = baseNode.content ?? [];
  const currentContent = currentNode.content ?? [];
  if (baseContent.length !== currentContent.length) return true;

  return baseContent.every((child, index) => alignedIdentitiesMatch(child, currentContent[index]));
}

export function collaborationProjectionIdentityMatches(input: {
  baseDocument: CollaborationJsonNode;
  currentDocument: CollaborationJsonNode;
  canonicalSidecar: FidelitySidecar;
  suppliedSidecar: FidelitySidecar;
}): boolean {
  if (!jsonValuesEqual(input.canonicalSidecar, input.suppliedSidecar)) return false;

  const baseIdentities = collectCollaborationIdentities(input.baseDocument);
  const currentIdentities = collectCollaborationIdentities(input.currentDocument);
  if (!baseIdentities || !currentIdentities) return false;

  for (const [collaborationId, currentIdentity] of currentIdentities) {
    const baseIdentity = baseIdentities.get(collaborationId);
    if (!baseIdentity) continue;
    if (baseIdentity.nodeType !== currentIdentity.nodeType) return false;
    if (baseIdentity.stableHint || currentIdentity.stableHint) {
      if (baseIdentity.stableHint !== currentIdentity.stableHint) return false;
    } else if (baseIdentity.path !== currentIdentity.path) {
      return false;
    }
  }

  return alignedIdentitiesMatch(input.baseDocument, input.currentDocument);
}

export function rehydrateCollaborationDocument(
  document: CollaborationJsonNode,
  sidecar: FidelitySidecar,
  diagnostics: CollaborationDiagnostic[] = []
): CollaborationJsonNode {
  const visit = (node: CollaborationJsonNode): CollaborationJsonNode => {
    const attrs = node.attrs ? { ...node.attrs } : undefined;
    const collaborationId =
      typeof attrs?.collaborationId === 'string' ? attrs.collaborationId : undefined;
    const entry = collaborationId ? sidecar.entries[collaborationId] : undefined;

    if (entry && entry.nodeType !== node.type) {
      diagnostics.push({
        code: 'sidecar_entry_mismatch',
        severity: 'error',
        nodeType: node.type,
      });
    } else if (entry && attrs) {
      Object.assign(attrs, entry.attrs);
    }

    return {
      ...node,
      attrs,
      content: node.content?.map(visit),
      marks: node.marks?.map((mark) => ({
        ...mark,
        attrs: mark.attrs ? { ...mark.attrs } : undefined,
      })),
    };
  };

  return visit(document);
}

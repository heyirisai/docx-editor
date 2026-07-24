export { projectProseMirrorDocument, rehydrateCollaborationDocument } from './projection';
export {
  exportCollaborationDocument,
  parseCollaborationPackage,
  CollaborationIdentityMismatchError,
  CollaborationExportFidelityError,
  IncompleteCollaborationManifestError,
  MissingCollaborationAssetError,
} from './workerApi';
export type {
  CollaborationDiagnostic,
  CollaborationExportAsset,
  CollaborationExportInput,
  CollaborationJsonMark,
  CollaborationJsonNode,
  CollaborationProjection,
  CollaborationProjectionOptions,
  CollaborationProjectionStats,
  FidelitySidecar,
  FidelitySidecarEntry,
  ParsedCollaborationPackage,
  ProseMirrorProjectionInput,
} from './types';
export type { CollaborationExportFidelityFailure } from './sourcePreservingExport';

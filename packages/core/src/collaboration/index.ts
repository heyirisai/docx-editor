export { projectProseMirrorDocument, rehydrateCollaborationDocument } from './projection';
// A plain (non-collaboration) export needs the same source-preserving patch, so
// the host can keep untouched paragraphs byte-identical instead of re-serializing
// markup the model cannot represent (VML fallbacks, WordArt, data-bound SDTs).
export {
  analyzeCollaborationParagraphChanges,
  preserveSourceXmlAroundParagraphChanges,
} from './sourcePreservingExport';
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

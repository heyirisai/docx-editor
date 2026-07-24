import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { createEmptyDocument, parseCollaborationPackage } from '@eigenpal/docx-editor-core';
import type { ImageAssetResolver } from '@eigenpal/docx-editor-core/layout-painter';
import { schema } from '@eigenpal/docx-editor-core/prosemirror';
import type { Document } from '@eigenpal/docx-editor-core/types/document';
import { DocxEditor, type DocxEditorRef } from '@eigenpal/docx-editor-react';
import { prosemirrorToYXmlFragment } from 'y-prosemirror';
import { GitHubBadge } from '../../shared/GitHubBadge';
import { createCachedImageAssetResolver } from '../../shared/cachedImageAssetResolver';
import { loadExternalMediaFixture } from '../../shared/externalMediaFixture';
import { AvatarStack } from './AvatarStack';
import { useCollaboration, yHistoryOverride } from './useCollaboration';
import { getOrCreateRoomFromUrl, loadOrCreateUser } from './identity';

declare global {
  interface Window {
    __COLLAB_PERF__?: {
      getClientId: () => number;
      getDocSize: () => number | null;
      getPeerCount: () => number;
      getTextPrefix: () => string;
      isReady: () => boolean;
    };
  }
}

const styles: Record<string, React.CSSProperties> = {
  container: {
    display: 'flex',
    flexDirection: 'column',
    height: '100vh',
    overflow: 'hidden',
    background: '#f8fafc',
  },
  main: {
    flex: 1,
    display: 'flex',
    overflow: 'hidden',
  },
  status: {
    fontSize: 12,
    color: '#64748b',
    padding: '4px 8px',
    background: '#f1f5f9',
    borderRadius: 4,
    whiteSpace: 'nowrap',
  },
  shareButton: {
    padding: '6px 12px',
    background: '#fff',
    border: '1px solid #e2e8f0',
    borderRadius: 6,
    cursor: 'pointer',
    fontSize: 13,
    fontWeight: 500,
    color: '#334155',
    whiteSpace: 'nowrap',
  },
};

const statusDotStyle = (color: string): React.CSSProperties => ({
  width: 8,
  height: 8,
  borderRadius: '50%',
  background: color,
  display: 'inline-block',
  marginRight: 6,
});

function useResponsiveLayout() {
  const calcZoom = () => {
    const pageWidth = 816 + 48;
    const vw = window.innerWidth;
    return vw < pageWidth ? Math.max(0.35, Math.floor((vw / pageWidth) * 20) / 20) : 1.0;
  };

  const [zoom, setZoom] = useState(calcZoom);
  const [isMobile, setIsMobile] = useState(() => window.innerWidth <= 768);

  useEffect(() => {
    const onResize = () => {
      setZoom(calcZoom());
      setIsMobile(window.innerWidth <= 768);
    };
    window.addEventListener('resize', onResize);
    return () => window.removeEventListener('resize', onResize);
  }, []);

  return { zoom, isMobile };
}

export function App() {
  const [user] = useState(loadOrCreateUser);
  const [room] = useState(getOrCreateRoomFromUrl);
  const [shareCopied, setShareCopied] = useState(false);
  const editorRef = useRef<DocxEditorRef>(null);
  const query = useMemo(() => new URLSearchParams(window.location.search), []);
  const externalMediaPerf = query.get('externalMediaPerf') === '1';
  const shouldSeedExternalMedia = query.get('seed') === '1';
  const [externalDocument, setExternalDocument] = useState<Document | null>(null);
  const [externalLoadError, setExternalLoadError] = useState<string | null>(null);
  const [editorReady, setEditorReady] = useState(false);

  const { ydoc, plugins, users, status, comments, setComments } = useCollaboration(room, user, {
    localOnly: externalMediaPerf,
  });
  const { zoom: autoZoom, isMobile } = useResponsiveLayout();

  // Empty document acts purely as a schema seed. ySyncPlugin populates the real
  // content from the Y.Doc, which is why we set externalContent on the editor.
  const seedDocument = useMemo(() => createEmptyDocument(), []);
  const externalImageAssets = useMemo(() => {
    if (!externalMediaPerf) return undefined;
    return createCachedImageAssetResolver(
      (assetId) => `/e2e-external-assets/${encodeURIComponent(assetId)}`
    );
  }, [externalMediaPerf]);
  const externalImageAssetResolver: ImageAssetResolver | undefined = externalImageAssets?.resolver;
  useEffect(() => () => externalImageAssets?.dispose(), [externalImageAssets]);

  useEffect(() => {
    if (!externalMediaPerf) return;

    const controller = new AbortController();
    void loadExternalMediaFixture(controller.signal, parseCollaborationPackage)
      .then((parsed) => {
        if (controller.signal.aborted) return;

        const fragment = ydoc.getXmlFragment('prosemirror');
        if (shouldSeedExternalMedia && fragment.length === 0) {
          prosemirrorToYXmlFragment(schema.nodeFromJSON(parsed.projection.document), fragment);
        }
        setExternalDocument(parsed.document);
      })
      .catch((error: unknown) => {
        if (controller.signal.aborted) return;
        setExternalLoadError(
          error instanceof Error ? error.message : 'Failed to load collaboration fixture'
        );
      });

    return () => controller.abort();
  }, [externalMediaPerf, shouldSeedExternalMedia, ydoc]);

  useEffect(() => {
    if (!externalMediaPerf) return;
    window.__COLLAB_PERF__ = {
      getClientId: () => ydoc.clientID,
      getDocSize: () => editorRef.current?.getEditorRef()?.getState?.()?.doc.content.size ?? null,
      getPeerCount: () => users.length,
      getTextPrefix: () => {
        const doc = editorRef.current?.getEditorRef()?.getState?.()?.doc;
        if (!doc) return '';
        return doc.textBetween(0, Math.min(doc.content.size, 4_096), '\n');
      },
      isReady: () => editorReady,
    };
    return () => {
      delete window.__COLLAB_PERF__;
    };
  }, [editorReady, externalMediaPerf, users, ydoc]);

  const handleCopyShareLink = useCallback(async () => {
    try {
      await navigator.clipboard.writeText(window.location.href);
      setShareCopied(true);
      setTimeout(() => setShareCopied(false), 1500);
    } catch {
      setShareCopied(false);
    }
  }, []);

  const renderLogo = useCallback(() => <GitHubBadge />, []);

  const renderTitleBarRight = useCallback(
    () => (
      <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
        <AvatarStack users={users} />
        <span style={styles.status} title={`Room: ${room}`}>
          <span
            style={statusDotStyle(
              status === 'connected' ? '#22c55e' : status === 'connecting' ? '#eab308' : '#ef4444'
            )}
          />
          {status === 'connected'
            ? `Live · ${room}`
            : status === 'connecting'
              ? 'Connecting…'
              : 'Offline'}
        </span>
        <button style={styles.shareButton} onClick={handleCopyShareLink}>
          {shareCopied ? 'Link copied!' : 'Share link'}
        </button>
      </div>
    ),
    [users, room, status, handleCopyShareLink, shareCopied]
  );

  return (
    <div style={styles.container}>
      <main style={styles.main}>
        {externalLoadError ? (
          <span role="alert">Error: {externalLoadError}</span>
        ) : externalMediaPerf && !externalDocument ? (
          <span>Loading external-media collaboration fixture...</span>
        ) : (
          <DocxEditor
            ref={editorRef}
            document={externalDocument ?? seedDocument}
            externalContent
            externalPlugins={plugins}
            historyOverride={yHistoryOverride}
            comments={comments}
            onCommentsChange={setComments}
            author={user.name}
            showToolbar
            showRuler={!isMobile}
            showZoomControl
            initialZoom={autoZoom}
            imageAssetResolver={externalImageAssetResolver}
            forcePageVirtualization={externalMediaPerf}
            onEditorViewReady={() => setEditorReady(true)}
            renderLogo={renderLogo}
            documentName={`Shared document — ${room}`}
            renderTitleBarRight={renderTitleBarRight}
          />
        )}
      </main>
    </div>
  );
}

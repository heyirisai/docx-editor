/**
 * Media garbage collection on repack.
 *
 * Repacking preserves every part of the original package, and image
 * registration appends new media + rels — so any duplicate or replaced
 * image lives forever, and the stale rel entries keep it "referenced".
 * A live-collab session snapshotting every 60s compounded one proposal
 * to 228MB of media of which 9 images (~3MB) were actually displayed.
 *
 * This pass applies Word's own save-time semantics AFTER all parts have
 * been serialized into the new zip:
 *
 *  1. An IMAGE relationship is live only if its rId appears anywhere in
 *     the owning part's XML (`r:embed`, VML `r:id`/`o:relid`, …). The
 *     check is deliberately liberal — a bare substring match — so a rel
 *     is only dropped when its id provably never occurs.
 *  2. A `word/media/*` file is live only if some remaining relationship
 *     (of any type, in any rels file) resolves to it. Rels whose owning
 *     part can't be found (e.g. the package-level `_rels/.rels`) keep
 *     everything they reference.
 *
 * Only image-type rels are ever removed; every other relationship type
 * is left untouched.
 */

import type JSZip from 'jszip';
import { RELATIONSHIP_TYPES } from '../relsParser';

/** Resolve a rels Target against the rels file's base part directory. */
function resolveTarget(relsPath: string, target: string): string {
  // 'word/_rels/document.xml.rels' → base 'word/'; '_rels/.rels' → base ''.
  const baseDir = relsPath.replace(/_rels\/[^/]*$/, '');
  const raw = target.startsWith('/') ? target.slice(1) : baseDir + target;
  const segments: string[] = [];
  for (const seg of raw.split('/')) {
    if (seg === '..') segments.pop();
    else if (seg !== '.' && seg !== '') segments.push(seg);
  }
  return segments.join('/');
}

/** The part a rels file describes: 'word/_rels/header1.xml.rels' → 'word/header1.xml'. */
function ownerPartPath(relsPath: string): string | null {
  const m = relsPath.match(/^(.*)_rels\/([^/]+)\.rels$/);
  return m ? m[1] + m[2] : null;
}

interface RelElement {
  element: string;
  id: string | null;
  type: string | null;
  target: string | null;
  external: boolean;
  /** Only self-closing elements are ever pruned — removing the open tag of a
   * `<Relationship>…</Relationship>` pair would leave orphaned close tags.
   * (Liveness marking covers both forms.) */
  selfClosing: boolean;
}

function parseRelElements(relsXml: string): RelElement[] {
  const out: RelElement[] = [];
  for (const m of relsXml.matchAll(/<Relationship\b[^>]*?>/g)) {
    const element = m[0];
    out.push({
      element,
      id: /\bId="([^"]*)"/.exec(element)?.[1] ?? null,
      type: /\bType="([^"]*)"/.exec(element)?.[1] ?? null,
      target: /\bTarget="([^"]*)"/.exec(element)?.[1] ?? null,
      external: /\bTargetMode="External"/.test(element),
      selfClosing: element.endsWith('/>'),
    });
  }
  return out;
}

/**
 * Remove image rels whose rId is unused in their owning part, then remove
 * `word/media/*` files no remaining relationship resolves to.
 */
export async function pruneUnreferencedMedia(zip: JSZip, compressionLevel: number): Promise<void> {
  const relsPaths: string[] = [];
  const mediaPaths: string[] = [];
  zip.forEach((path, file) => {
    if (file.dir) return;
    if (path.endsWith('.rels')) relsPaths.push(path);
    else if (path.startsWith('word/media/')) mediaPaths.push(path);
  });
  if (mediaPaths.length === 0) return;

  const liveMedia = new Set<string>();

  for (const relsPath of relsPaths) {
    const relsFile = zip.file(relsPath);
    if (!relsFile) continue;
    const relsXml = await relsFile.async('text');
    const rels = parseRelElements(relsXml);
    if (rels.length === 0) continue;

    const ownerPath = ownerPartPath(relsPath);
    const ownerFile = ownerPath ? zip.file(ownerPath) : null;
    const ownerXml = ownerFile ? await ownerFile.async('text') : null;

    let updatedXml = relsXml;
    let changed = false;
    for (const rel of rels) {
      const prunable =
        rel.selfClosing &&
        rel.type === RELATIONSHIP_TYPES.image &&
        !rel.external &&
        rel.id !== null &&
        // No owning XML to check against → keep (package-level rels, etc.).
        ownerXml !== null &&
        !ownerXml.includes(rel.id);
      if (prunable) {
        updatedXml = updatedXml.replace(rel.element, '');
        changed = true;
        continue;
      }
      if (!rel.external && rel.target) {
        const resolved = resolveTarget(relsPath, rel.target);
        if (resolved.startsWith('word/media/')) liveMedia.add(resolved);
      }
    }
    if (changed) {
      zip.file(relsPath, updatedXml, {
        compression: 'DEFLATE',
        compressionOptions: { level: compressionLevel },
      });
    }
  }

  for (const path of mediaPaths) {
    if (!liveMedia.has(path)) zip.remove(path);
  }
}

interface ExternalMediaManifestEntry {
  assetId: string;
  path: string;
  mimeType: string;
}

interface ExternalMediaParserInput {
  packageBuffer: ArrayBuffer;
  manifest: readonly ExternalMediaManifestEntry[];
  baseRevision: string;
}

interface FixtureManifest {
  base_revision: string;
  assets: Array<{
    asset_id: string;
    media_path: string;
    content_type: string;
  }>;
}

/**
 * Fetch the E2E package and manifest, then hand normalized input to the core
 * parser supplied by the example.
 */
export async function loadExternalMediaFixture<T>(
  signal: AbortSignal,
  parse: (input: ExternalMediaParserInput) => Promise<T>
): Promise<T> {
  const [packageResponse, manifestResponse] = await Promise.all([
    fetch('/e2e-external-package.docx', { signal }),
    fetch('/e2e-external-manifest.json', { signal }),
  ]);
  if (!packageResponse.ok) {
    throw new Error('Failed to load external-media package');
  }
  if (!manifestResponse.ok) {
    throw new Error('Failed to load external-media manifest');
  }

  const [packageBuffer, fixtureManifest] = await Promise.all([
    packageResponse.arrayBuffer(),
    manifestResponse.json() as Promise<FixtureManifest>,
  ]);
  return parse({
    packageBuffer,
    manifest: fixtureManifest.assets.map((asset) => ({
      assetId: asset.asset_id,
      path: asset.media_path,
      mimeType: asset.content_type,
    })),
    baseRevision: fixtureManifest.base_revision,
  });
}

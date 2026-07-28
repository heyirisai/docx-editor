import { expect, test, type Page } from '@playwright/test';
import JSZip from 'jszip';
import { readFile } from 'node:fs/promises';

declare global {
  interface Window {
    __COLLAB_KEY_LATENCY__?: Promise<number>;
  }
}

interface FixtureManifestAsset {
  asset_id: string;
  media_path: string;
  content_type: string;
}

interface FixtureManifest {
  assets: FixtureManifestAsset[];
}

interface FixtureData {
  packageBytes: Buffer;
  manifestBytes: Buffer;
  manifest: FixtureManifest;
  sourceZip: JSZip;
}

interface FixtureRequestStats {
  requestedAssetIds: Set<string>;
  getAssetRequestCounts: () => Map<string, number>;
  getTotalAssetRequests: () => number;
  getPeakAssetRequests: () => number;
}

interface FixtureRequestSnapshot {
  total: number;
  countsByAssetId: Array<[string, number]>;
}

const packagePath = process.env.COLLAB_PERF_PACKAGE_PATH;
const manifestPath = process.env.COLLAB_PERF_MANIFEST_PATH;
const sourcePath = process.env.COLLAB_PERF_SOURCE_PATH;
const modifier = process.platform === 'darwin' ? 'Meta' : 'Control';

async function readFixtureData(): Promise<FixtureData> {
  const [packageBytes, manifestBytes, sourceBytes] = await Promise.all([
    readFile(packagePath!),
    readFile(manifestPath!),
    readFile(sourcePath!),
  ]);
  return {
    packageBytes,
    manifestBytes,
    manifest: JSON.parse(manifestBytes.toString('utf8')) as FixtureManifest,
    sourceZip: await JSZip.loadAsync(sourceBytes),
  };
}

async function installFixtureRoutes(
  page: Page,
  fixture: FixtureData
): Promise<FixtureRequestStats> {
  const assetsById = new Map(
    fixture.manifest.assets.map((asset) => [asset.asset_id, asset] as const)
  );
  const requestedAssetIds = new Set<string>();
  const assetRequestCounts = new Map<string, number>();
  let totalAssetRequests = 0;
  let activeAssetRequests = 0;
  let peakAssetRequests = 0;

  await page.route('**/e2e-external-package.docx', (route) =>
    route.fulfill({
      status: 200,
      contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      body: fixture.packageBytes,
    })
  );
  await page.route('**/e2e-external-manifest.json', (route) =>
    route.fulfill({
      status: 200,
      contentType: 'application/json',
      body: fixture.manifestBytes,
    })
  );
  await page.route('**/e2e-external-assets/*', async (route) => {
    const assetId = decodeURIComponent(new URL(route.request().url()).pathname.split('/').at(-1)!);
    const asset = assetsById.get(assetId);
    if (!asset) {
      await route.fulfill({ status: 404, body: 'Unknown asset' });
      return;
    }
    const zipEntry = fixture.sourceZip.file(asset.media_path);
    if (!zipEntry) {
      await route.fulfill({ status: 404, body: 'Missing source media' });
      return;
    }

    requestedAssetIds.add(assetId);
    assetRequestCounts.set(assetId, (assetRequestCounts.get(assetId) ?? 0) + 1);
    totalAssetRequests++;
    activeAssetRequests++;
    peakAssetRequests = Math.max(peakAssetRequests, activeAssetRequests);
    try {
      const bytes = await zipEntry.async('uint8array');
      await route.fulfill({
        status: 200,
        contentType: asset.content_type,
        body: Buffer.from(bytes),
      });
    } finally {
      activeAssetRequests--;
    }
  });

  return {
    requestedAssetIds,
    getAssetRequestCounts: () => new Map(assetRequestCounts),
    getTotalAssetRequests: () => totalAssetRequests,
    getPeakAssetRequests: () => peakAssetRequests,
  };
}

function snapshotAssetRequests(stats: FixtureRequestStats): FixtureRequestSnapshot {
  return {
    total: stats.getTotalAssetRequests(),
    countsByAssetId: [...stats.getAssetRequestCounts()].sort(([left], [right]) =>
      left.localeCompare(right)
    ),
  };
}

function expectAssetRequestsUnchanged(
  stats: FixtureRequestStats,
  snapshot: FixtureRequestSnapshot
): void {
  expect(stats.getTotalAssetRequests()).toBe(snapshot.total);
  expect(
    [...stats.getAssetRequestCounts()].sort(([left], [right]) => left.localeCompare(right))
  ).toEqual(snapshot.countsByAssetId);
}

function expectEachAssetFetchedOnce(stats: FixtureRequestStats): void {
  expect(stats.getTotalAssetRequests()).toBe(stats.requestedAssetIds.size);
  expect([...stats.getAssetRequestCounts().values()]).toEqual(
    Array.from({ length: stats.requestedAssetIds.size }, () => 1)
  );
}

function captureBrowserErrors(page: Page): string[] {
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  page.on('console', (message) => {
    if (message.type() === 'error') errors.push(message.text());
  });
  return errors;
}

async function nextPaint(page: Page): Promise<void> {
  await page.evaluate(
    () =>
      new Promise<void>((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)))
  );
}

async function waitForExternalImageRequestsToSettle(page: Page): Promise<void> {
  await expect
    .poll(() => page.locator('img[data-asset-id][data-asset-state="ready"]').count(), {
      timeout: 30_000,
    })
    .toBeGreaterThan(0);
  await expect
    .poll(
      () =>
        page
          .locator(
            'img[data-asset-id][data-asset-state="queued"], ' +
              'img[data-asset-id][data-asset-state="loading"], ' +
              'img[data-asset-id][data-asset-state="decoding"]'
          )
          .count(),
      { timeout: 30_000 }
    )
    .toBe(0);
  await nextPaint(page);
}

async function measureKeystrokes(page: Page, count: number, key = 'a'): Promise<number[]> {
  const latencies: number[] = [];
  for (let index = 0; index < count; index++) {
    await page.evaluate(() => {
      window.__COLLAB_KEY_LATENCY__ = new Promise<number>((resolve) => {
        window.addEventListener(
          'keydown',
          () => {
            const startedAt = performance.now();
            requestAnimationFrame(() =>
              requestAnimationFrame(() => resolve(performance.now() - startedAt))
            );
          },
          { capture: true, once: true }
        );
      });
    });
    await page.keyboard.press(key);
    const latency = await page.evaluate(async () => {
      const pending = window.__COLLAB_KEY_LATENCY__;
      delete window.__COLLAB_KEY_LATENCY__;
      if (!pending) throw new Error('Key latency measurement was not armed');
      return pending;
    });
    latencies.push(Math.round(latency));
  }
  return latencies;
}

function percentile(values: number[], percentileValue: number): number {
  const ordered = [...values].sort((left, right) => left - right);
  return ordered[Math.max(0, Math.ceil(ordered.length * percentileValue) - 1)]!;
}

async function focusDocumentStart(page: Page): Promise<void> {
  const firstParagraph = page.locator('[data-page-number="1"] .layout-paragraph').first();
  const firstParagraphBox = await firstParagraph.boundingBox();
  expect(firstParagraphBox).not.toBeNull();
  await page.mouse.click(firstParagraphBox!.x + 40, firstParagraphBox!.y + 8);
  await page.keyboard.press(`${modifier}+Home`);
}

async function waitForCollaborationEditor(page: Page): Promise<void> {
  await expect(page.locator('text=Loading external-media collaboration fixture...')).toHaveCount(
    0,
    { timeout: 90_000 }
  );
  await expect(page.locator('[data-testid="docx-editor"]')).toBeVisible({
    timeout: 90_000,
  });
  await expect
    .poll(() => page.evaluate(() => window.__COLLAB_PERF__?.isReady() ?? false), {
      timeout: 90_000,
    })
    .toBe(true);
  await expect
    .poll(() => page.evaluate(() => window.__COLLAB_PERF__?.getDocSize() ?? 0), {
      timeout: 90_000,
    })
    .toBeGreaterThan(1_000);
  await page.waitForSelector('[data-page-number]', { timeout: 90_000 });
  await nextPaint(page);
}

async function scrollToFraction(page: Page, fraction: number): Promise<void> {
  await page.evaluate((targetFraction) => {
    const container =
      document.querySelector<HTMLElement>('.docx-editor__scroll-container') ??
      document.querySelector<HTMLElement>('.paged-editor__scroll-container') ??
      document.querySelector<HTMLElement>('.paged-editor__pages');
    if (!container) throw new Error('Paged editor scroll container is missing');
    container.scrollTop = (container.scrollHeight - container.clientHeight) * targetFraction;
  }, fraction);
  await nextPaint(page);
  await nextPaint(page);
}

async function getRenderedPageNumbers(page: Page): Promise<number[]> {
  return page
    .locator('[data-page-number]:has(.layout-page-content)')
    .evaluateAll((pages) =>
      pages.map((pageElement) => Number((pageElement as HTMLElement).dataset.pageNumber))
    );
}

test.describe('external-media collaboration performance', () => {
  test.skip(
    !packagePath || !manifestPath || !sourcePath,
    'Set COLLAB_PERF_PACKAGE_PATH, COLLAB_PERF_MANIFEST_PATH, and COLLAB_PERF_SOURCE_PATH'
  );
  test.describe.configure({ timeout: 120_000, mode: 'serial' });

  test('collaboration off: bounds rendering and keeps the real proposal editable', async ({
    page,
  }) => {
    const fixture = await readFixtureData();
    const requestStats = await installFixtureRoutes(page, fixture);
    const browserErrors = captureBrowserErrors(page);

    const loadStartedAt = performance.now();
    await page.goto('/?e2e=1&externalMediaPerf=1', {
      waitUntil: 'domcontentloaded',
      timeout: 90_000,
    });
    await expect(page.locator('text=Loading external-media fixture...')).toHaveCount(0, {
      timeout: 90_000,
    });
    await expect(page.locator('[data-testid="docx-editor"]')).toBeVisible({
      timeout: 90_000,
    });
    await page.waitForSelector('[data-page-number]', { timeout: 90_000 });
    await nextPaint(page);
    const loadMs = Math.round(performance.now() - loadStartedAt);

    const totalPages = await page.locator('[data-page-number]').count();
    expect(totalPages).toBeGreaterThan(1);
    expect(requestStats.requestedAssetIds.size).toBeGreaterThan(0);
    expect(requestStats.requestedAssetIds.size).toBeLessThan(fixture.manifest.assets.length);
    expect(requestStats.getPeakAssetRequests()).toBeLessThanOrEqual(4);
    await waitForExternalImageRequestsToSettle(page);
    expectEachAssetFetchedOnce(requestStats);
    const requestsBeforeTyping = snapshotAssetRequests(requestStats);

    await focusDocumentStart(page);
    await page.keyboard.press('x');
    await nextPaint(page);
    await measureKeystrokes(page, 5, 'w');
    const latencies = await measureKeystrokes(page, 20);
    const p95Ms = percentile(latencies, 0.95);
    const p99Ms = percentile(latencies, 0.99);

    expect(p95Ms).toBeLessThan(100);
    expect(p99Ms).toBeLessThan(150);
    expectAssetRequestsUnchanged(requestStats, requestsBeforeTyping);

    const renderedPageNumbersByScroll: number[][] = [];
    for (const fraction of [0.25, 0.5, 0.75, 1]) {
      await scrollToFraction(page, fraction);
      renderedPageNumbersByScroll.push(await getRenderedPageNumbers(page));
    }
    const renderedPagesByScroll = renderedPageNumbersByScroll.map((pages) => pages.length);
    // Documents below the eight-page virtualization threshold render every
    // page eagerly. Longer documents retain at most the visible page plus the
    // configured buffer while scrolling.
    expect(Math.max(...renderedPagesByScroll)).toBeLessThanOrEqual(Math.min(totalPages, 7));
    expect(renderedPageNumbersByScroll.at(-1)).toContain(totalPages);
    expect(requestStats.requestedAssetIds.size).toBeLessThan(fixture.manifest.assets.length);
    expect(requestStats.getPeakAssetRequests()).toBeLessThanOrEqual(4);
    expect(browserErrors).toEqual([]);

    console.log(
      JSON.stringify({
        mode: 'collaboration-off',
        loadMs,
        totalPages,
        manifestAssets: fixture.manifest.assets.length,
        requestedAssets: requestStats.requestedAssetIds.size,
        assetRequests: requestStats.getTotalAssetRequests(),
        peakAssetRequests: requestStats.getPeakAssetRequests(),
        renderedPagesByScroll,
        renderedPageNumbersByScroll,
        typingP95Ms: p95Ms,
        typingP99Ms: p99Ms,
        typingLatenciesMs: latencies,
      })
    );
  });

  test('one player: live Yjs stays within the local typing budget', async ({ page }) => {
    test.skip(!process.env.COLLAB_PERF_MULTI, 'Set COLLAB_PERF_MULTI=1');
    const fixture = await readFixtureData();
    const requestStats = await installFixtureRoutes(page, fixture);
    const browserErrors = captureBrowserErrors(page);
    const room = `ideagen-one-${Date.now()}`;

    const loadStartedAt = performance.now();
    await page.goto(
      `http://localhost:5273/?externalMediaPerf=1&seed=1#${encodeURIComponent(room)}`,
      { waitUntil: 'domcontentloaded', timeout: 90_000 }
    );
    await waitForCollaborationEditor(page);
    const loadMs = Math.round(performance.now() - loadStartedAt);
    await waitForExternalImageRequestsToSettle(page);
    expectEachAssetFetchedOnce(requestStats);
    const requestsBeforeTyping = snapshotAssetRequests(requestStats);

    await focusDocumentStart(page);
    await measureKeystrokes(page, 5, 'w');
    const latencies = await measureKeystrokes(page, 20);
    const p95Ms = percentile(latencies, 0.95);
    const p99Ms = percentile(latencies, 0.99);

    expect(p95Ms).toBeLessThan(100);
    expect(p99Ms).toBeLessThan(150);
    expectAssetRequestsUnchanged(requestStats, requestsBeforeTyping);
    expect(requestStats.requestedAssetIds.size).toBeGreaterThan(0);
    expect(requestStats.requestedAssetIds.size).toBeLessThan(fixture.manifest.assets.length);
    expect(requestStats.getPeakAssetRequests()).toBeLessThanOrEqual(4);
    expect(browserErrors).toEqual([]);

    console.log(
      JSON.stringify({
        mode: 'one-live-player',
        loadMs,
        manifestAssets: fixture.manifest.assets.length,
        requestedAssets: requestStats.requestedAssetIds.size,
        assetRequests: requestStats.getTotalAssetRequests(),
        peakAssetRequests: requestStats.getPeakAssetRequests(),
        typingP95Ms: p95Ms,
        typingP99Ms: p99Ms,
        typingLatenciesMs: latencies,
      })
    );
  });

  test('two players: concurrent typing and 100ms-RTT remote paint stay responsive', async ({
    browser,
  }) => {
    test.skip(!process.env.COLLAB_PERF_MULTI, 'Set COLLAB_PERF_MULTI=1');
    const fixture = await readFixtureData();
    const context = await browser.newContext();
    await context.addInitScript(() => {
      const NativeBroadcastChannel = globalThis.BroadcastChannel;
      class DelayedBroadcastChannel extends NativeBroadcastChannel {
        override postMessage(message: unknown): void {
          window.setTimeout(() => super.postMessage(message), 50);
        }
      }
      Object.defineProperty(globalThis, 'BroadcastChannel', {
        configurable: true,
        value: DelayedBroadcastChannel,
      });
    });
    const first = await context.newPage();
    const second = await context.newPage();
    const firstStats = await installFixtureRoutes(first, fixture);
    const secondStats = await installFixtureRoutes(second, fixture);
    const firstErrors = captureBrowserErrors(first);
    const secondErrors = captureBrowserErrors(second);
    const room = `ideagen-two-${Date.now()}`;

    await first.goto(
      `http://localhost:5273/?externalMediaPerf=1&seed=1#${encodeURIComponent(room)}`,
      { waitUntil: 'domcontentloaded', timeout: 90_000 }
    );
    await waitForCollaborationEditor(first);
    await second.goto(`http://localhost:5273/?externalMediaPerf=1#${encodeURIComponent(room)}`, {
      waitUntil: 'domcontentloaded',
      timeout: 90_000,
    });
    await waitForCollaborationEditor(second);

    await expect
      .poll(
        async () => {
          const [firstPeers, secondPeers] = await Promise.all([
            first.evaluate(() => window.__COLLAB_PERF__?.getPeerCount() ?? 0),
            second.evaluate(() => window.__COLLAB_PERF__?.getPeerCount() ?? 0),
          ]);
          return Math.min(firstPeers, secondPeers);
        },
        { timeout: 10_000 }
      )
      .toBeGreaterThanOrEqual(2);

    await Promise.all([
      waitForExternalImageRequestsToSettle(first),
      waitForExternalImageRequestsToSettle(second),
    ]);
    expectEachAssetFetchedOnce(firstStats);
    expectEachAssetFetchedOnce(secondStats);
    const firstRequestsBeforeTyping = snapshotAssetRequests(firstStats);
    const secondRequestsBeforeTyping = snapshotAssetRequests(secondStats);

    await Promise.all([focusDocumentStart(first), focusDocumentStart(second)]);
    await Promise.all([measureKeystrokes(first, 5, 'w'), measureKeystrokes(second, 5, 'x')]);
    const [firstLocalLatencies, secondLocalLatencies] = await Promise.all([
      measureKeystrokes(first, 20, 'a'),
      measureKeystrokes(second, 20, 'b'),
    ]);
    const localLatencies = [...firstLocalLatencies, ...secondLocalLatencies];
    const localP95Ms = percentile(localLatencies, 0.95);
    const localP99Ms = percentile(localLatencies, 0.99);

    await expect
      .poll(
        async () => {
          const [firstSize, secondSize] = await Promise.all([
            first.evaluate(() => window.__COLLAB_PERF__?.getDocSize() ?? 0),
            second.evaluate(() => window.__COLLAB_PERF__?.getDocSize() ?? 0),
          ]);
          return firstSize === secondSize;
        },
        { timeout: 5_000 }
      )
      .toBe(true);

    const startingRemoteSize = await second.evaluate(
      () => window.__COLLAB_PERF__?.getDocSize() ?? 0
    );
    const remoteLatencies: number[] = [];
    for (let index = 0; index < 10; index++) {
      const startedAt = performance.now();
      await first.keyboard.press('c');
      await expect
        .poll(() => second.evaluate(() => window.__COLLAB_PERF__?.getDocSize() ?? 0), {
          intervals: [5, 10, 20],
          timeout: 2_000,
        })
        .toBeGreaterThanOrEqual(startingRemoteSize + index + 1);
      await nextPaint(second);
      remoteLatencies.push(Math.round(performance.now() - startedAt));
    }

    const remoteP95Ms = percentile(remoteLatencies, 0.95);
    expect(localP95Ms).toBeLessThan(100);
    expect(localP99Ms).toBeLessThan(150);
    expect(remoteP95Ms).toBeLessThan(350);
    expectAssetRequestsUnchanged(firstStats, firstRequestsBeforeTyping);
    expectAssetRequestsUnchanged(secondStats, secondRequestsBeforeTyping);
    expect(firstStats.getPeakAssetRequests()).toBeLessThanOrEqual(4);
    expect(secondStats.getPeakAssetRequests()).toBeLessThanOrEqual(4);
    expect(firstErrors).toEqual([]);
    expect(secondErrors).toEqual([]);

    console.log(
      JSON.stringify({
        mode: 'two-live-players',
        simulatedRoundTripMs: 100,
        localTypingP95Ms: localP95Ms,
        localTypingP99Ms: localP99Ms,
        firstPlayerLatenciesMs: firstLocalLatencies,
        secondPlayerLatenciesMs: secondLocalLatencies,
        remotePaintP95Ms: remoteP95Ms,
        remotePaintLatenciesMs: remoteLatencies,
        firstPlayerRequestedAssets: firstStats.requestedAssetIds.size,
        secondPlayerRequestedAssets: secondStats.requestedAssetIds.size,
        firstPlayerAssetRequests: firstStats.getTotalAssetRequests(),
        secondPlayerAssetRequests: secondStats.getTotalAssetRequests(),
        firstPlayerPeakAssetRequests: firstStats.getPeakAssetRequests(),
        secondPlayerPeakAssetRequests: secondStats.getPeakAssetRequests(),
      })
    );

    await context.close();
  });
});

import logger from '../logger.js';
import fs from 'fs';
import path from 'path';
import { MockRegistry } from './registry.js';
import { loadDefaultMocks } from './defaults.js';
import { applyOverridesFromFile } from './loader.js';
import { MockResponse } from './MockResponse.js';
import { stripQuery } from './pathMatcher.js';

let registry: MockRegistry | null = null;
let initialized = false;

export function isDryRunEnabled(): boolean {
  const v = process.env.MS365_MCP_DRYRUN;
  return v === '1' || (v ?? '').toLowerCase() === 'true';
}

type DryrunMode = 'off' | 'mock' | 'partial';
let selectedMode: DryrunMode | null = null;

function fileExists(filePath: string): boolean {
  try {
    const abs = path.isAbsolute(filePath) ? filePath : path.join(process.cwd(), filePath);
    return fs.existsSync(abs) && fs.statSync(abs).isFile();
  } catch {
    return false;
  }
}

export function getDryrunMode(): DryrunMode {
  if (selectedMode) return selectedMode;
  const file = process.env.MS365_MCP_DRYRUN_FILE;
  if (file && fileExists(file)) {
    selectedMode = 'mock';
  } else if (isDryRunEnabled()) {
    selectedMode = 'partial';
  } else {
    selectedMode = 'off';
  }
  logger.info('[DRYRUN] mode selected', {
    mode: selectedMode,
    hasDryrunEnv: isDryRunEnabled(),
    dryrunFile: file || null,
    dryrunFileExists: file ? fileExists(file) : false,
  });
  return selectedMode;
}

export function isMockMode(): boolean {
  return getDryrunMode() === 'mock';
}

export function ensureMocksInitialized(): void {
  if (initialized) return;
  registry = new MockRegistry();
  const seedStr = process.env.MS365_MCP_DRYRUN_SEED;
  const seed = seedStr ? Number(seedStr) : undefined;
  loadDefaultMocks(registry, seed);
  applyOverridesFromFile(registry);
  initialized = true;
  logger.info('[DRYRUN:MOCK] mock registry initialized');
}

export async function mockFetch(
  method: string,
  endpoint: string,
  headers: Record<string, string> = {},
  body?: string
): Promise<MockResponse> {
  if (!registry) {
    ensureMocksInitialized();
  }
  const pathOnly = stripQuery(endpoint);
  let entry = registry!.find(method, pathOnly);
  const query = new URLSearchParams(endpoint.includes('?') ? endpoint.split('?')[1] : '');

  if (entry) {
    try {
      const parsedBody = parseBody(body);
      const res = await entry.handler({
        method,
        path: pathOnly,
        params: entry.params,
        headers,
        body: parsedBody,
        query,
      });
      logger.info(`[DRYRUN:MOCK] ${method} ${endpoint} → mocked ${res.status}`);
      return res;
    } catch (e) {
      logger.error(`[DRYRUN:MOCK] mock handler failed: ${(e as Error).message}`);
      return new MockResponse(
        { error: 'Mock handler error', message: (e as Error).message },
        { status: 500, statusText: 'Mock Error' }
      );
    }
  }

  // Fallback: if /users/{id}/... has no mock, try corresponding /me/... mock
  if (pathOnly.startsWith('/users/')) {
    const altPath = pathOnly.replace(/^\/users\/[^/]+/, '/me');
    const altEntry = registry!.find(method, altPath);
    if (altEntry) {
      try {
        const parsedBody = parseBody(body);
        const res = await altEntry.handler({
          method,
          path: altPath,
          params: altEntry.params,
          headers,
          body: parsedBody,
          query,
        });
        logger.info(`[DRYRUN:MOCK] ${method} ${endpoint} → fallback to ${altPath} mocked ${res.status}`);
        return res;
      } catch (e) {
        logger.error(`[DRYRUN:MOCK] fallback mock handler failed: ${(e as Error).message}`);
        return new MockResponse(
          { error: 'Mock handler error', message: (e as Error).message },
          { status: 500, statusText: 'Mock Error' }
        );
      }
    }
  }

  // Generic fallbacks
  if (method.toUpperCase() === 'GET') {
    return new MockResponse({ value: [] }, { status: 200 });
  }
  if (method.toUpperCase() === 'DELETE') {
    return new MockResponse('', { status: 204, statusText: 'No Content' });
  }
  if (method.toUpperCase() === 'POST') {
    return new MockResponse('', { status: 202, statusText: 'Accepted' });
  }
  return new MockResponse({ success: true }, { status: 200 });
}

function parseBody(body?: string): unknown {
  if (!body) return undefined;
  try {
    return JSON.parse(body);
  } catch {
    return body;
  }
}



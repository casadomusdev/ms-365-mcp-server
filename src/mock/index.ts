import logger from '../logger.js';
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

export function ensureMocksInitialized(): void {
  if (initialized) return;
  registry = new MockRegistry();
  const seedStr = process.env.MS365_MCP_DRYRUN_SEED;
  const seed = seedStr ? Number(seedStr) : undefined;
  loadDefaultMocks(registry, seed);
  applyOverridesFromFile(registry);
  initialized = true;
  logger.info('[dryrun] mock registry initialized');
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
      logger.info(`[dryrun] ${method} ${endpoint} → mocked ${res.status}`);
      return res;
    } catch (e) {
      logger.error(`[dryrun] mock handler failed: ${(e as Error).message}`);
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
        logger.info(`[dryrun] ${method} ${endpoint} → fallback to ${altPath} mocked ${res.status}`);
        return res;
      } catch (e) {
        logger.error(`[dryrun] fallback mock handler failed: ${(e as Error).message}`);
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



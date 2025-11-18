import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import GraphClient from '../src/graph-client.js';
import path from 'path';

vi.mock('../src/logger.js', () => {
  return {
    default: {
      info: vi.fn(),
      error: vi.fn(),
      debug: vi.fn(),
    },
  };
});

describe('Dry-run mock mode via DRYRUN_FILE', () => {
  const authManager = { getToken: async () => 'test-token' };

  beforeEach(() => {
    vi.clearAllMocks();
    delete process.env.MS365_MCP_DRYRUN;
    process.env.MS365_MCP_DRYRUN_FILE = path.join(process.cwd(), 'test/fixtures/dryrun-mocks.json');
    delete process.env.MS365_MCP_CLIENT_SECRET; // avoid /me rewrite branch
  });

  afterEach(() => {
    vi.resetAllMocks();
    delete process.env.MS365_MCP_DRYRUN_FILE;
  });

  it('routes GET through mocks when overrides file exists', async () => {
    const fetchSpy = vi.fn();
    // @ts-ignore
    global.fetch = fetchSpy;

    const client = new GraphClient(authManager as any);
    const result = await client.makeRequest('/me/messages', { method: 'GET' });

    expect(result).toEqual({
      value: [{ id: 'm1', subject: 'Hello from fixture' }],
    });
    expect(fetchSpy).not.toHaveBeenCalled();
  });
});



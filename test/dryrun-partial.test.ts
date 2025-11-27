import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import GraphClient from '../src/graph-client.js';

vi.mock('../src/logger.js', () => {
  return {
    default: {
      info: vi.fn(),
      error: vi.fn(),
      debug: vi.fn(),
    },
  };
});

describe('Dry-run partial mode', () => {
  const authManager = { getToken: async () => 'test-token' };

  beforeEach(() => {
    vi.clearAllMocks();
    // Partial mode: MS365_MCP_DRYRUN=true and no valid DRYRUN_FILE
    process.env.MS365_MCP_DRYRUN = 'true';
    delete process.env.MS365_MCP_DRYRUN_FILE;
    delete process.env.MS365_MCP_CLIENT_SECRET; // avoid /me rewrite branch
  });

  afterEach(() => {
    vi.resetAllMocks();
    delete process.env.MS365_MCP_DRYRUN;
    delete process.env.MS365_MCP_DRYRUN_FILE;
  });

  it('passes through GET requests to real fetch', async () => {
    const fetchSpy = vi.fn(async () => ({
      ok: true,
      status: 200,
      statusText: 'OK',
      text: async () => JSON.stringify({ value: [{ id: '1' }] }),
      headers: { get: () => null },
    }));
    (global as any).fetch = fetchSpy;

    const client = new GraphClient(authManager as any);
    const result = await client.makeRequest('/me/messages', { method: 'GET' });

    expect(result).toEqual({ value: [{ id: '1' }] });
    expect(fetchSpy).toHaveBeenCalledTimes(1);
    expect(fetchSpy.mock.calls[0][0]).toMatch('https://graph.microsoft.com/v1.0/me/messages');
  });

  it('suppresses POST requests and returns accepted without calling fetch', async () => {
    const fetchSpy = vi.fn();
    (global as any).fetch = fetchSpy;

    const client = new GraphClient(authManager as any);
    const result = await client.makeRequest('/me/sendMail', {
      method: 'POST',
      body: JSON.stringify({ message: { subject: 'Hello' } }),
    });

    // In GraphClient, empty body maps to { message: 'OK!' }
    expect(result).toEqual({ message: 'OK!' });
    expect(fetchSpy).not.toHaveBeenCalled();
  });
});

import logger from '../logger.js';
import { matchPattern } from './pathMatcher.js';
import { MockResponse } from './MockResponse.js';

export type MockContext = {
  method: string;
  path: string;
  params: Record<string, string>;
  headers: Record<string, string>;
  body?: unknown;
  query?: URLSearchParams;
};

export type MockHandler = (ctx: MockContext) => MockResponse | Promise<MockResponse>;

type MockEntry = {
  method: string;
  pattern: string;
  handler: MockHandler;
};

export class MockRegistry {
  private entries: MockEntry[] = [];

  registerMock(method: string, pattern: string, handler: MockHandler): void {
    const m = method.toUpperCase();
    this.entries.push({ method: m, pattern, handler });
    logger.info(`[dryrun] registered mock: ${m} ${pattern}`);
  }

  find(method: string, path: string): { handler: MockHandler; params: Record<string, string> } | null {
    const m = method.toUpperCase();
    // Prefer last-registered entries (overrides take precedence)
    for (let i = this.entries.length - 1; i >= 0; i--) {
      const entry = this.entries[i];
      if (entry.method !== m) continue;
      const res = matchPattern(entry.pattern, path);
      if (res) {
        return { handler: entry.handler, params: res.params };
      }
    }
    return null;
  }
}



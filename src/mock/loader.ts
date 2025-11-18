import fs from 'fs';
import path from 'path';
import logger from '../logger.js';
import { MockRegistry } from './registry.js';
import { MockResponse } from './MockResponse.js';
import { stripQuery } from './pathMatcher.js';

type OverrideEntry =
  | unknown
  | {
      status?: number;
      headers?: Record<string, string>;
      body?: unknown;
    };

export function applyOverridesFromFile(reg: MockRegistry): void {
  const file = process.env.MS365_MCP_DRYRUN_FILE;
  if (!file) return;
  try {
    const abs = path.isAbsolute(file) ? file : path.join(process.cwd(), file);
    if (!fs.existsSync(abs)) {
      logger.warn(`[DRYRUN:MOCK] overrides file not found: ${abs}`);
      return;
    }
    const raw = fs.readFileSync(abs, 'utf8');
    const json = JSON.parse(raw) as Record<string, OverrideEntry>;
    for (const key of Object.keys(json)) {
      // key syntax: "METHOD /path/pattern"
      const spaceIdx = key.indexOf(' ');
      if (spaceIdx <= 0) {
        logger.warn(`[DRYRUN:MOCK] invalid override key: ${key}`);
        continue;
      }
      const method = key.slice(0, spaceIdx).toUpperCase();
      const pattern = stripQuery(key.slice(spaceIdx + 1).trim());
      const val = json[key];
      reg.registerMock(method, pattern, () => {
        if (
          val &&
          typeof val === 'object' &&
          !Array.isArray(val) &&
          ('status' in (val as any) || 'headers' in (val as any) || 'body' in (val as any))
        ) {
          const v = val as { status?: number; headers?: Record<string, string>; body?: unknown };
          return new MockResponse(v.body ?? {}, {
            status: v.status ?? 200,
            headers: v.headers,
          });
        }
        return new MockResponse(val);
      });
    }
    logger.info(`[DRYRUN:MOCK] overrides loaded from ${abs}`);
  } catch (e) {
    logger.error(`[DRYRUN:MOCK] failed to load overrides: ${(e as Error).message}`);
  }
}



import logger from '../logger.js';

export type MailboxInfo = {
  id: string;
  email: string;
  displayName?: string;
  type: 'personal' | 'shared' | 'delegated';
  permissions: Array<'read' | 'write' | 'send'>;
};

type UserMailboxCache = {
  userEmail: string;
  discoveredAt: number;
  expiresAt: number;
  allowedMailboxes: MailboxInfo[];
};

export class MailboxDiscoveryCache {
  private cache = new Map<string, UserMailboxCache>();
  private cacheTTLms: number;

  constructor() {
    const ttlSec = Number(process.env.MS365_MCP_IMPERSONATE_CACHE_TTL || '3600');
    this.cacheTTLms = Math.max(60, ttlSec) * 1000;
  }

  async getMailboxes(userEmail: string): Promise<MailboxInfo[]> {
    const key = userEmail.toLowerCase();
    const existing = this.cache.get(key);
    const now = Date.now();
    if (existing && existing.expiresAt > now) {
      return existing.allowedMailboxes;
    }

    const allowed = await this.discoverMailboxes(userEmail);
    this.cache.set(key, {
      userEmail,
      discoveredAt: now,
      expiresAt: now + this.cacheTTLms,
      allowedMailboxes: allowed,
    });
    return allowed;
  }

  // Phase 0: strict + optional allowlist via env
  private async discoverMailboxes(userEmail: string): Promise<MailboxInfo[]> {
    const result: MailboxInfo[] = [];
    // Always include the impersonated user's own mailbox
    result.push({
      id: userEmail,
      email: userEmail,
      displayName: userEmail,
      type: 'personal',
      permissions: ['read', 'write', 'send'],
    });

    const allowListEnv = process.env.MS365_MCP_IMPERSONATE_ALLOWED_MAILBOXES || '';
    const extras = allowListEnv
      .split(',')
      .map((s) => s.trim())
      .filter((s) => s && s.toLowerCase() !== userEmail.toLowerCase());

    for (const email of extras) {
      result.push({
        id: email,
        email,
        displayName: email,
        type: 'delegated',
        permissions: ['read', 'write', 'send'],
      });
    }

    logger.info(
      `Impersonation mailbox set for ${userEmail}: ${result.map((m) => m.email).join(', ')}`
    );
    return result;
  }
}

export default MailboxDiscoveryCache;



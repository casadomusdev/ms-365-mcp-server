import logger from '../logger.js';
import type GraphClient from '../graph-client.js';
import type AuthManager from '../auth.js';
import { MailboxDiscoveryService } from './MailboxDiscoveryService.js';
import PowerShellService from '../lib/PowerShellService.js';

/**
 * Mailbox information with type and permissions.
 */
export type MailboxInfo = {
  id: string;
  email: string;
  displayName?: string;
  type: 'personal' | 'shared' | 'delegated';
  permissions: Array<'read' | 'write' | 'send'>;
};

/**
 * Cache entry for discovered mailboxes.
 */
type UserMailboxCache = {
  userEmail: string;
  discoveredAt: number;
  expiresAt: number;
  allowedMailboxes: MailboxInfo[];
};

/**
 * Caching layer for mailbox discovery results.
 * 
 * This class provides TTL-based caching for mailbox discovery operations.
 * It delegates actual discovery to MailboxDiscoveryService and caches the results
 * to minimize expensive Graph API calls.
 * 
 * Environment variables:
 * - MS365_MCP_IMPERSONATE_CACHE_TTL: Cache TTL in seconds (default: 3600)
 * 
 * Cache behavior:
 * - Results cached per user email (case-insensitive)
 * - Automatic expiration based on TTL
 * - Cache hit returns immediately without API calls
 * - Cache miss triggers full discovery via service
 * 
 * @example
 * ```typescript
 * const cache = new MailboxDiscoveryCache(graphClient);
 * const mailboxes = await cache.getMailboxes('user@company.com');
 * // First call: Full discovery via Graph API (~2-5 seconds)
 * // Subsequent calls within TTL: Instant cache hit
 * ```
 */
export class MailboxDiscoveryCache {
  private readonly cache = new Map<string, UserMailboxCache>();
  private readonly cacheTTLms: number;
  private readonly service: MailboxDiscoveryService;

  /**
   * Creates a new MailboxDiscoveryCache instance.
   * 
   * @param graphClient - Authenticated GraphClient instance for making Microsoft Graph API calls
   * @param authManager - AuthManager instance for PowerShell authentication
   */
  constructor(graphClient: GraphClient, authManager: AuthManager) {
    // Create PowerShellService for mailbox discovery
    const powerShellService = new PowerShellService(authManager);
    
    // Create discovery service with optional PowerShell support
    this.service = new MailboxDiscoveryService(graphClient, powerShellService);
    
    const ttlSec = Number(process.env.MS365_MCP_IMPERSONATE_CACHE_TTL || '3600');
    this.cacheTTLms = Math.max(60, ttlSec) * 1000;
    
    logger.info(`MailboxDiscoveryCache initialized with TTL: ${ttlSec} seconds`);
  }

  /**
   * Gets mailboxes for a user, using cache when available.
   * 
   * This method checks the cache first and returns cached results if still valid.
   * On cache miss or expiration, it delegates to MailboxDiscoveryService for
   * full discovery and caches the results.
   * 
   * @param userEmail - Email address of the user to get mailboxes for
   * @returns Array of accessible mailboxes with metadata
   * @throws Error if user cannot be found or if Graph API access fails
   */
  async getMailboxes(userEmail: string): Promise<MailboxInfo[]> {
    const key = userEmail.toLowerCase();
    const now = Date.now();
    
    // Check cache first
    const existing = this.cache.get(key);
    if (existing && existing.expiresAt > now) {
      const age = Math.round((now - existing.discoveredAt) / 1000);
      logger.debug(`Cache hit for ${userEmail} (age: ${age}s, ${existing.allowedMailboxes.length} mailboxes)`);
      return existing.allowedMailboxes;
    }

    // Cache miss or expired - perform discovery
    if (existing) {
      logger.debug(`Cache expired for ${userEmail}, re-discovering...`);
    } else {
      logger.debug(`Cache miss for ${userEmail}, discovering mailboxes...`);
    }
    
    const allowed = await this.service.discoverMailboxes(userEmail);
    
    // Store in cache
    this.cache.set(key, {
      userEmail,
      discoveredAt: now,
      expiresAt: now + this.cacheTTLms,
      allowedMailboxes: allowed,
    });
    
    logger.info(
      `Discovered ${allowed.length} mailbox(es) for ${userEmail}: ${allowed.map((m) => m.email).join(', ')}`
    );
    
    return allowed;
  }

  /**
   * Clears all cached discovery results.
   * 
   * Useful for testing or when you need to force rediscovery.
   */
  clearCache(): void {
    this.cache.clear();
    logger.info('Mailbox discovery cache cleared');
  }

  /**
   * Gets current cache statistics.
   * 
   * @returns Object with cache size and entry details
   */
  getCacheStats(): { 
    size: number; 
    entries: Array<{ 
      email: string; 
      mailboxCount: number; 
      age: number;
      expiresIn: number;
    }> 
  } {
    const now = Date.now();
    const entries = Array.from(this.cache.values()).map(entry => ({
      email: entry.userEmail,
      mailboxCount: entry.allowedMailboxes.length,
      age: Math.round((now - entry.discoveredAt) / 1000), // age in seconds
      expiresIn: Math.max(0, Math.round((entry.expiresAt - now) / 1000)), // time until expiry in seconds
    }));

    return {
      size: this.cache.size,
      entries,
    };
  }
}

export default MailboxDiscoveryCache;

import logger from '../logger.js';
import type GraphClient from '../graph-client.js';
import type { MailboxInfo } from './MailboxDiscoveryCache.js';
import type PowerShellService from '../lib/PowerShellService.js';

/**
 * Service for discovering mailboxes accessible to a user.
 * 
 * Detection Strategy:
 * - Personal mailbox (via Graph API - always included)
 * - Shared mailboxes (via Exchange Online PowerShell when enabled)
 * 
 * PowerShell Integration:
 * - When MS365_POWERSHELL_ENABLED=true, uses Exchange Online PowerShell to query
 *   actual mailbox delegation permissions (Full Access, SendAs)
 * - This provides accurate detection of shared mailbox access
 * - Falls back to personal mailbox only if PowerShell is disabled
 * 
 * How It Works:
 * 1. Gets the user's personal mailbox via Graph API (fast, always works)
 * 2. If PowerShell enabled: Queries Exchange Online for shared mailbox permissions
 * 3. Combines personal + shared mailboxes into final result
 * 
 * Performance characteristics:
 * - Personal mailbox: Fast (~100ms via Graph API)
 * - Shared mailboxes: Slow (~2-5s via PowerShell, but cached with 1 hour TTL)
 * - Cache is critical for acceptable performance
 * 
 * @example
 * ```typescript
 * const service = new MailboxDiscoveryService(graphClient, powerShellService);
 * const mailboxes = await service.discoverMailboxes('user@company.com');
 * // Returns: [{ id: '...', email: 'user@company.com', type: 'personal', ... }, ...]
 * ```
 */
export class MailboxDiscoveryService {
  private readonly graphClient: GraphClient;
  private readonly powerShellService: PowerShellService | null;
  private readonly debugMode: boolean;

  /**
   * Creates a new MailboxDiscoveryService instance.
   * 
   * @param graphClient - Authenticated GraphClient instance for making Microsoft Graph API calls
   * @param powerShellService - Optional PowerShellService for querying Exchange mailbox permissions
   */
  constructor(graphClient: GraphClient, powerShellService?: PowerShellService) {
    this.graphClient = graphClient;
    this.powerShellService = powerShellService || null;
    this.debugMode = process.env.MS365_MCP_IMPERSONATE_DEBUG === 'true' || 
                     process.env.MS365_MCP_IMPERSONATE_DEBUG === '1';
    
    if (this.debugMode) {
      logger.info('[MailboxDiscovery] Debug mode enabled - verbose logging active');
    }
    
    if (this.powerShellService?.isEnabled()) {
      logger.info('[MailboxDiscovery] PowerShell integration enabled for shared mailbox discovery');
    } else {
      logger.info('[MailboxDiscovery] PowerShell integration disabled - only personal mailbox will be discovered');
    }
  }

  /**
   * Discovers all mailboxes accessible to the specified user.
   * 
   * This method discovers mailboxes using a hybrid approach:
   * 1. Personal mailbox via Graph API (always included, fast)
   * 2. Shared mailboxes via Exchange Online PowerShell (if enabled)
   * 
   * When PowerShell is enabled, this provides accurate detection of shared mailbox
   * permissions (Full Access, SendAs) that cannot be queried via Graph API.
   * 
   * @param userEmail - Email address of the user to discover mailboxes for
   * @returns Array of discovered mailboxes with metadata
   * @throws Error if user cannot be found or if Graph API access fails
   */
  async discoverMailboxes(userEmail: string): Promise<MailboxInfo[]> {
    logger.info(`[MailboxDiscovery] Starting discovery for ${userEmail}`);
    const mailboxes: MailboxInfo[] = [];

    // 1. Get the user's personal mailbox (via Graph API - always works)
    try {
      const userData = await this.graphClient.makeRequest(
        `/users/${encodeURIComponent(userEmail)}`,
        {}
      ) as any;

      mailboxes.push({
        id: userData.id,
        type: 'personal',
        displayName: userData.displayName,
        email: userData.userPrincipalName || userData.mail,
        permissions: ['read', 'write', 'send'],
      });
      logger.info(`[MailboxDiscovery] ✓ Found personal mailbox: ${userData.displayName}`);
    } catch (error) {
      logger.error(`[MailboxDiscovery] ✗ Error fetching user ${userEmail}: ${(error as Error).message}`);
      throw new Error(`Could not find user ${userEmail}: ${(error as Error).message}`);
    }

    // 2. Discover shared mailboxes (via PowerShell if enabled)
    if (this.powerShellService?.isEnabled()) {
      try {
        logger.info('[MailboxDiscovery] Querying shared mailboxes via Exchange Online PowerShell...');
        const sharedMailboxes = await this.powerShellService.checkPermissions(userEmail);
        
        // Convert PowerShell results to MailboxInfo format
        for (const mailbox of sharedMailboxes) {
          mailboxes.push({
            id: mailbox.id,
            type: 'shared',
            displayName: mailbox.displayName,
            email: mailbox.email,
            permissions: ['read', 'write', 'send'], // All PowerShell-detected mailboxes have full access
          });
        }
        
        logger.info(`[MailboxDiscovery] ✓ Found ${sharedMailboxes.length} shared mailbox(es) via PowerShell`);
      } catch (error) {
        logger.error(`[MailboxDiscovery] ✗ Error querying shared mailboxes via PowerShell: ${(error as Error).message}`);
        logger.warn('[MailboxDiscovery] Continuing with personal mailbox only');
      }
    } else {
      logger.info('[MailboxDiscovery] PowerShell integration disabled - skipping shared mailbox discovery');
      logger.info('[MailboxDiscovery] Set MS365_POWERSHELL_ENABLED=true to enable shared mailbox detection');
    }

    // Log summary
    const sharedCount = mailboxes.filter(m => m.type === 'shared').length;
    logger.info(`[MailboxDiscovery] Discovery complete for ${userEmail}:`);
    logger.info(`[MailboxDiscovery]   Total accessible: ${mailboxes.length} (1 personal + ${sharedCount} shared)`);
    
    if (sharedCount === 0 && this.powerShellService?.isEnabled()) {
      logger.info(`[MailboxDiscovery] No shared mailboxes found for this user`);
    }

    return mailboxes;
  }
}

export default MailboxDiscoveryService;

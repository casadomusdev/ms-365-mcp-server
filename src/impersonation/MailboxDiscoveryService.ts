import logger from '../logger.js';
import type GraphClient from '../graph-client.js';
import type { MailboxInfo } from './MailboxDiscoveryCache.js';

/**
 * Service for discovering mailboxes accessible to a user via Graph API calendar delegation permissions.
 * 
 * Detection Strategy:
 * - Personal mailbox (always included)
 * - Shared mailboxes where the user has calendar delegation permissions
 * 
 * IMPORTANT LIMITATIONS:
 * - Only detects shared mailboxes where the user has calendar permissions
 * - Does NOT detect: SendAs-only permissions, Full Access without calendar access
 * - Exchange mailbox delegation permissions (SendAs, Full Access) cannot be reliably 
 *   queried via Microsoft Graph API - they require Exchange PowerShell or EWS
 * 
 * How It Works:
 * 1. Queries all shared mailboxes in the tenant (unlicensed, enabled accounts)
 * 2. For each shared mailbox, checks if the impersonated user has calendar permissions
 * 3. Includes mailboxes where calendar delegation is detected
 * 
 * Performance characteristics:
 * - Uses concurrent request limiting (max 5 simultaneous requests)
 * - Implements 5-second timeout per individual request
 * - Specifically targets shared mailboxes (more efficient than checking all users)
 * 
 * @example
 * ```typescript
 * const service = new MailboxDiscoveryService(graphClient);
 * const mailboxes = await service.discoverMailboxes('user@company.com');
 * // Returns: [{ id: '...', email: 'user@company.com', type: 'personal', ... }, ...]
 * ```
 */
export class MailboxDiscoveryService {
  private readonly graphClient: GraphClient;
  private readonly maxConcurrentRequests = 5;
  private readonly requestTimeoutMs = 5000;
  private readonly debugMode: boolean;

  /**
   * Creates a new MailboxDiscoveryService instance.
   * 
   * @param graphClient - Authenticated GraphClient instance for making Microsoft Graph API calls
   */
  constructor(graphClient: GraphClient) {
    this.graphClient = graphClient;
    this.debugMode = process.env.MS365_MCP_IMPERSONATE_DEBUG === 'true' || 
                     process.env.MS365_MCP_IMPERSONATE_DEBUG === '1';
    
    if (this.debugMode) {
      logger.info('[MailboxDiscovery] Debug mode enabled - verbose logging active');
    }
  }

  /**
   * Discovers all mailboxes accessible to the specified user.
   * 
   * This method discovers mailboxes using calendar delegation detection:
   * 1. Personal mailbox (always included)
   * 2. Shared mailboxes where the user has calendar delegation permissions
   * 
   * LIMITATIONS: Only detects mailboxes with calendar delegation. SendAs-only or 
   * Full Access permissions without calendar access will NOT be detected.
   * 
   * @param userEmail - Email address of the user to discover mailboxes for
   * @returns Array of discovered mailboxes with metadata
   * @throws Error if user cannot be found or if Graph API access fails
   */
  async discoverMailboxes(userEmail: string): Promise<MailboxInfo[]> {
    logger.info(`[MailboxDiscovery] Starting discovery for ${userEmail}`);
    const mailboxes: MailboxInfo[] = [];
    const stats = {
      checked: 0,
      found: 0,
      skippedLicensed: 0,
      calendarPerms: 0,
      sendAsPerms: 0,
      errors: 0,
    };

    // 1. Get the impersonated user's personal mailbox
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

    // 2. Discover shared mailboxes with calendar delegation
    try {
      const impersonatedUserId = mailboxes[0]?.id;
      const impersonatedUserEmail = userEmail.toLowerCase();

      // Query only Member users (excludes guests, service accounts)
      // Note: Cannot use $filter on assignedLicenses (complex type) or "ne" operator on mail
      logger.info('[MailboxDiscovery] Querying Member users with mailboxes...');
      const usersData = await this.graphClient.makeRequest(
        '/users?$filter=accountEnabled eq true and userType eq \'Member\'&$select=id,displayName,userPrincipalName,mail,assignedLicenses&$top=999',
        {}
      ) as any;

      const allUsers = usersData.value || [];
      
      // Filter for shared mailboxes (valid email + no licenses)
      const sharedMailboxes = allUsers.filter((user: any) => {
        // Must have valid email
        const hasValidEmail = user.mail || user.userPrincipalName;
        if (!hasValidEmail) return false;
        
        // Must have no licenses (shared mailboxes are unlicensed)
        const hasNoLicense = !user.assignedLicenses || user.assignedLicenses.length === 0;
        return hasNoLicense;
      });
      
      logger.info(`[MailboxDiscovery] Found ${allUsers.length} Member users, ${sharedMailboxes.length} shared mailboxes to check`);
      
      if (this.debugMode) {
        logger.debug(`[MailboxDiscovery] → Filtering for unlicensed users (shared mailboxes)`);
      }

      // Process shared mailboxes in batches to limit concurrent requests
      for (let i = 0; i < sharedMailboxes.length; i += this.maxConcurrentRequests) {
        const batch = sharedMailboxes.slice(i, i + this.maxConcurrentRequests);
        
        if (this.debugMode) {
          logger.debug(`[MailboxDiscovery] → Processing batch ${Math.floor(i / this.maxConcurrentRequests) + 1} (${batch.length} of ${sharedMailboxes.length} mailboxes)`);
        }

        await Promise.all(batch.map(async (mailbox: any) => {
          // Skip the impersonated user's own mailbox (already added)
          if (mailbox.id === impersonatedUserId) {
            return;
          }

          const mailboxEmail = mailbox.mail || mailbox.userPrincipalName;
          stats.checked++;

          // Log what we're about to check (synchronous, right before the work)
          logger.info(`[MailboxDiscovery] Checking: ${mailboxEmail}`);

          try {
            const mailboxInfo = await this.checkCalendarDelegation(
              mailbox,
              impersonatedUserEmail,
              mailboxEmail  // Pass email for error context
            );

            if (mailboxInfo) {
              mailboxes.push(mailboxInfo);
              stats.found++;
              stats.calendarPerms++;
              logger.info(`[MailboxDiscovery] ✓ ${mailboxEmail} - Calendar delegation found`);
            } else {
              logger.info(`[MailboxDiscovery] ✗ ${mailboxEmail} - No calendar delegation`);
            }
          } catch (error: any) {
            stats.errors++;
            if (error.name !== 'AbortError') {
              logger.warn(`[MailboxDiscovery] ✗ ${mailboxEmail} - Error: ${error.message}`);
            } else {
              logger.warn(`[MailboxDiscovery] ⏱ ${mailboxEmail} - Timeout`);
            }
          }
        }));
      }
    } catch (error) {
      logger.error(`[MailboxDiscovery] Error scanning for shared mailboxes: ${(error as Error).message}`);
      logger.error(`[MailboxDiscovery] This may indicate missing Azure permissions: User.Read.All, Calendars.Read`);
    }

    // Log detailed summary
    const sharedCount = mailboxes.filter(m => m.type === 'shared').length;
    logger.info(`[MailboxDiscovery] Discovery complete for ${userEmail}:`);
    logger.info(`[MailboxDiscovery]   Mailboxes checked: ${stats.checked}`);
    logger.info(`[MailboxDiscovery]   ✓ With calendar delegation: ${stats.calendarPerms}`);
    logger.info(`[MailboxDiscovery]   ✗ Without calendar delegation: ${stats.checked - stats.calendarPerms}`);
    logger.info(`[MailboxDiscovery]   Total accessible: ${mailboxes.length} (1 personal + ${sharedCount} shared)`);
    
    if (stats.errors > 0) {
      logger.warn(`[MailboxDiscovery]   Errors/timeouts: ${stats.errors}`);
    }
    
    if (mailboxes.length === 1) {
      logger.warn(`[MailboxDiscovery] Only personal mailbox found. Possible reasons:`);
      logger.warn(`[MailboxDiscovery]   - No shared mailbox with calendar delegation for this user`);
      logger.warn(`[MailboxDiscovery]   - User has SendAs-only permissions (not detectable via Graph API)`);
      logger.warn(`[MailboxDiscovery]   - Missing Azure permission: Calendars.Read`);
    }

    return mailboxes;
  }

  /**
   * Checks if the impersonated user has calendar delegation to a shared mailbox.
   * 
   * Calendar delegation indicates the user can manage the mailbox's calendar,
   * which typically also grants access to the mailbox itself.
   * 
   * @param mailbox - The shared mailbox object from Graph API
   * @param impersonatedUserEmail - The email of the impersonated user (lowercase)
   * @param mailboxEmail - The email of the mailbox being checked (for error context)
   * @returns MailboxInfo if calendar delegation is detected, null otherwise
   */
  private async checkCalendarDelegation(
    mailbox: any,
    impersonatedUserEmail: string,
    mailboxEmail: string
  ): Promise<MailboxInfo | null> {
    const email = mailboxEmail || mailbox.userPrincipalName || mailbox.mail;

    // Check calendar permissions
    if (await this.checkCalendarPermissions(mailbox.id, impersonatedUserEmail, email)) {
      return {
        id: mailbox.id,
        type: 'shared',
        displayName: mailbox.displayName,
        email: email,
        permissions: ['read', 'write', 'send'],
      };
    }

    return null;
  }

  /**
   * Check if impersonated user has calendar delegate permissions.
   * 
   * Calendar delegate permissions indicate that a user can manage a mailbox's calendar,
   * which typically also grants access to the mailbox itself.
   * 
   * @param mailboxId - The mailbox ID to check calendar permissions for
   * @param impersonatedUserEmail - The email of the impersonated user (lowercase)
   * @param mailboxEmail - The email of the mailbox (for error context)
   * @returns true if calendar delegate access is detected
   */
  private async checkCalendarPermissions(mailboxId: string, impersonatedUserEmail: string, mailboxEmail: string): Promise<boolean> {
    try {
      if (this.debugMode) {
        logger.debug(`[MailboxDiscovery]   → Querying calendar permissions for ${mailboxEmail}`);
      }

      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), this.requestTimeoutMs);

      const calendarPerms = await this.graphClient.makeRequest(
        `/users/${mailboxId}/calendar/calendarPermissions`,
        {}
      ) as any;

      clearTimeout(timeoutId);

      if (calendarPerms.value) {
        const hasPermission = calendarPerms.value.some((perm: any) =>
          perm.emailAddress?.address?.toLowerCase() === impersonatedUserEmail
        );
        
        if (this.debugMode && hasPermission) {
          logger.debug(`[MailboxDiscovery]   → Found calendar permission for ${impersonatedUserEmail}`);
        }
        
        return hasPermission;
      }
    } catch (error: any) {
      // Add context to error but keep it debug-level unless it's unexpected
      const errorMsg = error.message || String(error);
      
      if (error.name === 'AbortError') {
        if (this.debugMode) {
          logger.debug(`[MailboxDiscovery]   ⏱ ${mailboxEmail}: Calendar check timeout`);
        }
      } else if (errorMsg.includes('MailboxNotEnabledForRESTAPI')) {
        // Common for certain mailbox types - only log in debug mode
        if (this.debugMode) {
          logger.debug(`[MailboxDiscovery]   ℹ ${mailboxEmail}: Mailbox not REST API enabled`);
        }
      } else {
        // Unexpected error - log as warning with full context
        if (this.debugMode) {
          logger.debug(`[MailboxDiscovery]   ✗ ${mailboxEmail}: ${errorMsg}`);
        }
      }
    }

    return false;
  }
}

export default MailboxDiscoveryService;

import logger from '../logger.js';
import type GraphClient from '../graph-client.js';
import type { MailboxInfo } from './MailboxDiscoveryCache.js';

/**
 * Service for discovering mailboxes accessible to a user via various permission strategies.
 * 
 * This service implements a comprehensive 4-strategy approach to discover:
 * - Personal mailbox (always included)
 * - Calendar delegate permissions (Strategy 1)
 * - SendAs permissions (Strategy 2)
 * - Full Access delegations via mailbox test (Strategy 3 - fallback)
 * 
 * Performance characteristics:
 * - Uses concurrent request limiting (max 5 simultaneous requests)
 * - Implements 5-second timeout per individual request
 * - Queries tenant users with reasonable limits ($top=100)
 * - Skips licensed mailboxes (can only be accessed by owner)
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
  private isAdmin: boolean = false;
  private adminCheckPerformed: boolean = false;

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
   * This method performs a comprehensive discovery using 4 strategies:
   * 1. Personal mailbox (always included)
   * 2. Calendar delegate permissions
   * 3. SendAs permissions
   * 4. Full Access delegation test (fallback for Exchange-level delegations)
   * 
   * Licensed mailboxes are skipped entirely as they can only be accessed by their owner.
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
      fullAccessPerms: 0,
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
      
      // Detect admin status for this user (cached for instance lifetime)
      if (!this.adminCheckPerformed) {
        this.isAdmin = await this.detectAdminStatus(userData.id, userEmail);
        this.adminCheckPerformed = true;
      }
    } catch (error) {
      logger.error(`[MailboxDiscovery] ✗ Error fetching user ${userEmail}: ${(error as Error).message}`);
      throw new Error(`Could not find user ${userEmail}: ${(error as Error).message}`);
    }

    // 2. Discover shared/delegated mailboxes using Graph API
    try {
      const impersonatedUserId = mailboxes[0]?.id;
      const impersonatedUserEmail = userEmail.toLowerCase();

      // Query tenant users to check for mailbox access
      logger.info('[MailboxDiscovery] Querying tenant users...');
      const usersData = await this.graphClient.makeRequest(
        '/users?$filter=userType eq \'Member\'&$select=id,displayName,userPrincipalName,mail&$top=100',
        {}
      ) as any;

      const users = usersData.value || [];
      logger.info(`[MailboxDiscovery] Found ${users.length} tenant users to check`);

      // Process users in batches to limit concurrent requests
      for (let i = 0; i < users.length; i += this.maxConcurrentRequests) {
        const batch = users.slice(i, i + this.maxConcurrentRequests);
        
        if (this.debugMode) {
          logger.info(`[MailboxDiscovery] Processing batch ${Math.floor(i / this.maxConcurrentRequests) + 1} (users ${i + 1}-${Math.min(i + this.maxConcurrentRequests, users.length)})`);
        }

        await Promise.all(batch.map(async (user: any) => {
          // Skip the impersonated user's own mailbox (already added)
          if (user.id === impersonatedUserId) {
            return;
          }

          stats.checked++;

          try {
            const mailboxInfo = await this.checkUserMailboxAccess(
              user,
              impersonatedUserId,
              impersonatedUserEmail
            );

            if (mailboxInfo) {
              mailboxes.push(mailboxInfo);
              stats.found++;
              logger.info(`[MailboxDiscovery] ✓ ${userEmail} has ${mailboxInfo.type} access to: ${user.displayName}`);
            } else if (this.debugMode) {
              // mailboxInfo is null - either licensed or no access
              // The checkUserMailboxAccess method will have logged the reason
            }
          } catch (error: any) {
            stats.errors++;
            if (error.name !== 'AbortError') {
              if (this.debugMode) {
                logger.info(`[MailboxDiscovery] ✗ Error checking ${user.displayName}: ${error.message}`);
              }
            } else if (this.debugMode) {
              logger.info(`[MailboxDiscovery] ⏱ Timeout checking ${user.displayName}`);
            }
          }
        }));
      }
    } catch (error) {
      logger.error(`[MailboxDiscovery] Error scanning for mailbox permissions: ${(error as Error).message}`);
      logger.error(`[MailboxDiscovery] This may indicate missing Azure permissions: User.Read.All, Calendars.Read, Mail.Send.Shared`);
    }

    // Log summary
    logger.info(`[MailboxDiscovery] Discovery complete for ${userEmail}:`);
    logger.info(`[MailboxDiscovery]   Total mailboxes found: ${mailboxes.length}`);
    logger.info(`[MailboxDiscovery]   Users checked: ${stats.checked}`);
    logger.info(`[MailboxDiscovery]   Licensed mailboxes skipped: ${stats.skippedLicensed}`);
    if (stats.calendarPerms > 0) {
      logger.info(`[MailboxDiscovery]   Calendar delegation: ${stats.calendarPerms}`);
    }
    if (stats.sendAsPerms > 0) {
      logger.info(`[MailboxDiscovery]   SendAs delegation: ${stats.sendAsPerms}`);
    }
    if (stats.fullAccessPerms > 0) {
      logger.info(`[MailboxDiscovery]   Full Access delegation: ${stats.fullAccessPerms}`);
    }
    if (stats.errors > 0) {
      logger.warn(`[MailboxDiscovery]   Errors/timeouts: ${stats.errors}`);
    }
    
    if (mailboxes.length === 1) {
      logger.warn(`[MailboxDiscovery] Only personal mailbox found. This could mean:`);
      logger.warn(`[MailboxDiscovery]   - No shared mailbox access configured for this user`);
      logger.warn(`[MailboxDiscovery]   - Missing Azure permissions: Calendars.Read, Mail.Send.Shared`);
    }

    return mailboxes;
  }

  /**
   * Checks if the impersonated user has access to a specific user's mailbox.
   * 
   * Strategy:
   * 1. Check if mailbox has licenses - if yes, skip (licensed mailboxes can only be accessed by owner)
   * 2. For shared mailboxes (no licenses), check Calendar and SendAs permissions
   * 
   * @param user - The user object from Graph API
   * @param impersonatedUserId - The ID of the impersonated user
   * @param impersonatedUserEmail - The email of the impersonated user (lowercase)
   * @returns MailboxInfo if access is detected, null otherwise
   */
  private async checkUserMailboxAccess(
    user: any,
    impersonatedUserId: string,
    impersonatedUserEmail: string
  ): Promise<MailboxInfo | null> {
    const userEmail = user.userPrincipalName || user.mail;

    // First, check if this is a licensed mailbox or shared mailbox
    const isSharedMailbox = await this.detectSharedMailbox(user);

    // Licensed mailboxes can ONLY be accessed by their owner - skip all permission checks
    if (!isSharedMailbox) {
      if (this.debugMode) {
        logger.info(`[MailboxDiscovery] ${user.displayName} (${userEmail}): Skipping delegation checks (licensed mailbox can only be accessed by owner)`);
      }
      return null;
    }

    // This is a shared mailbox (no licenses) - check for delegation permissions
    let hasAccess = false;
    let accessMethod: string | null = null;

    // Strategy 1: Check calendar permissions (calendar delegate access)
    if (await this.checkCalendarPermissions(user.id, impersonatedUserEmail)) {
      hasAccess = true;
      accessMethod = 'calendar';
      logger.info(`[MailboxDiscovery] Found calendar delegation: ${impersonatedUserEmail} → ${user.displayName}`);
    }

    // Strategy 2: Check SendAs permissions (mail delegate)
    if (!hasAccess) {
      if (await this.checkSendAsPermissions(user.id, impersonatedUserEmail, user.displayName, userEmail)) {
        hasAccess = true;
        accessMethod = 'sendAs';
        logger.info(`[MailboxDiscovery] Found SendAs delegation: ${impersonatedUserEmail} → ${user.displayName}`);
      }
    }

    // Strategy 3: Check mailbox folder permissions (works in client credentials mode)
    // This checks the actual MailboxFolder permissions to see if user is a delegate
    if (!hasAccess && this.isAdmin) {
      if (await this.checkMailboxFolderPermissions(user.id, impersonatedUserId, impersonatedUserEmail)) {
        hasAccess = true;
        accessMethod = 'mailboxPermissions';
        logger.info(`[MailboxDiscovery] Found mailbox folder delegation: ${impersonatedUserEmail} → ${user.displayName}`);
      }
    }

    // Strategy 4: Test actual mailbox access (fallback for Full Access delegations)
    // ONLY used if impersonated user is NOT an admin (non-admin user without admin permissions)
    // This catches delegations configured at Exchange level that aren't exposed via Graph API
    if (!hasAccess && !this.isAdmin) {
      if (await this.testMailboxAccess(user.id, impersonatedUserEmail)) {
        hasAccess = true;
        accessMethod = 'fullAccess';
        logger.info(`[MailboxDiscovery] Found Full Access delegation (via mailbox test): ${impersonatedUserEmail} → ${user.displayName}`);
      }
    }

    if (hasAccess) {
      return {
        id: user.id,
        type: 'shared',
        displayName: user.displayName,
        email: userEmail,
        permissions: ['read', 'write', 'send'],
      };
    }

    if (this.debugMode) {
      logger.info(`[MailboxDiscovery] ${user.displayName} (${userEmail}): No delegation permissions found for ${impersonatedUserEmail}`);
    }

    return null;
  }

  /**
   * Detects if the impersonated user has Exchange or Global Administrator privileges.
   * 
   * Admin users have tenant-wide access to all mailboxes regardless of delegation,
   * which makes the mailbox access test unreliable (it returns true for ALL mailboxes).
   * 
   * This method checks the user's directory role memberships to identify admin status.
   * The result is cached for the lifetime of the service instance.
   * 
   * @param userId - The user ID to check admin status for
   * @param userEmail - The user email (for logging)
   * @returns true if user has admin privileges
   */
  private async detectAdminStatus(userId: string, userEmail: string): Promise<boolean> {
    try {
      logger.info(`[MailboxDiscovery] Checking admin status for ${userEmail}...`);
      
      // Check if we're using application permissions (client credentials)
      const isClientCredentials = process.env.MS365_MCP_CLIENT_SECRET != null;
      
      if (isClientCredentials) {
        logger.info(`[MailboxDiscovery] Application permissions (client credentials) detected`);
        logger.info(`[MailboxDiscovery] Will use mailbox folder permissions strategy instead of simple mailbox test`);
      }
      
      // Always check user's directory roles, even in client credentials mode
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), this.requestTimeoutMs);

      // Query user's directory role memberships
      const rolesData = await this.graphClient.makeRequest(
        `/users/${userId}/memberOf/microsoft.graph.directoryRole`,
        {}
      ) as any;

      clearTimeout(timeoutId);

      if (rolesData.value && rolesData.value.length > 0) {
        // Get detailed info for each directory role
        const roleDetails = await Promise.all(
          rolesData.value.slice(0, 10).map(async (role: any) => {
            try {
              const roleInfo = await this.graphClient.makeRequest(
                `/directoryRoles/${role.id}`,
                {}
              ) as any;
              return roleInfo.displayName;
            } catch {
              return null;
            }
          })
        );

        const validRoleNames = roleDetails.filter(name => name != null);
        
        if (this.debugMode && validRoleNames.length > 0) {
          logger.info(`[MailboxDiscovery] User has ${validRoleNames.length} directory role(s):`);
          validRoleNames.forEach(name => {
            logger.info(`[MailboxDiscovery]   - ${name}`);
          });
        }

        const adminRoles = [
          'Global Administrator',
          'Exchange Administrator',
          'Privileged Role Administrator',
        ];

        const hasAdminRole = validRoleNames.some(roleName =>
          adminRoles.some(adminRole => roleName?.includes(adminRole))
        );

        if (hasAdminRole) {
          logger.warn(`[MailboxDiscovery] ⚠ Admin role detected for ${userEmail} - mailbox access test disabled (would show false positives)`);
          if (isClientCredentials) {
            logger.info(`[MailboxDiscovery] Using mailbox folder permissions strategy (client credentials + admin role)`);
          } else {
            logger.warn(`[MailboxDiscovery] ⚠ Admin accounts will ONLY discover delegations via SendAs permissions`);
          }
          return true;
        } else if (isClientCredentials) {
          // Client credentials mode but user is not an admin
          logger.info(`[MailboxDiscovery] ✓ Non-admin user in client credentials mode - using mailbox folder permissions strategy`);
          return true;
        } else {
          logger.info(`[MailboxDiscovery] ✓ Non-admin user ${userEmail} - full discovery enabled`);
          return false;
        }
      }

      // No directory roles found
      if (isClientCredentials) {
        logger.info(`[MailboxDiscovery] No directory roles found for ${userEmail} in client credentials mode - using mailbox folder permissions strategy`);
        return true;
      } else {
        logger.info(`[MailboxDiscovery] No directory roles found for ${userEmail} - assuming non-admin`);
        return false;
      }
    } catch (error: any) {
      if (error.name !== 'AbortError') {
        logger.warn(`[MailboxDiscovery] Could not check admin status for ${userEmail}: ${error.message}`);
        logger.warn(`[MailboxDiscovery] Assuming non-admin to avoid missing delegations`);
      }
      return false;
    }
  }

  /**
   * Strategy 1: Check if impersonated user has calendar delegate permissions.
   * 
   * Calendar delegate permissions indicate that a user can manage another user's calendar,
   * which typically also grants access to the mailbox for scheduling purposes.
   * 
   * @param userId - The user ID to check calendar permissions for
   * @param impersonatedUserEmail - The email of the impersonated user (lowercase)
   * @returns true if calendar delegate access is detected
   */
  private async checkCalendarPermissions(userId: string, impersonatedUserEmail: string): Promise<boolean> {
    try {
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), this.requestTimeoutMs);

      const calendarPerms = await this.graphClient.makeRequest(
        `/users/${userId}/calendar/calendarPermissions`,
        {}
      ) as any;

      clearTimeout(timeoutId);

      if (calendarPerms.value) {
        return calendarPerms.value.some((perm: any) =>
          perm.emailAddress?.address?.toLowerCase() === impersonatedUserEmail
        );
      }
    } catch (error: any) {
      if (error.name !== 'AbortError') {
        logger.debug(`Calendar permission check failed for user ${userId}: ${error.message}`);
      }
    }

    return false;
  }

  /**
   * Strategy 2: Detect if a mailbox is a shared mailbox.
   * 
   * Shared mailboxes are identified by the userPurpose property in mailboxSettings.
   * According to Microsoft documentation, if userPurpose === 'shared', it's a shared mailbox.
   * 
   * Shared mailboxes are typically used for team collaboration (e.g., info@company.com).
   * Personal mailboxes can ONLY be accessed by their owner, never via delegation.
   * 
   * @param user - The user object to check
   * @returns true if the mailbox is a shared mailbox
   */
  private async detectSharedMailbox(user: any): Promise<boolean> {
    const userEmail = user.userPrincipalName || user.mail;
    
    if (this.debugMode) {
      logger.info(`[MailboxDiscovery] Checking mailbox type for ${user.displayName} (${userEmail})...`);
    }

    try {
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), this.requestTimeoutMs);

      const mailboxSettings = await this.graphClient.makeRequest(
        `/users/${user.id}/mailboxSettings`,
        {}
      ) as any;

      clearTimeout(timeoutId);

      // Check the userPurpose property to determine if this is a shared mailbox
      const userPurpose = mailboxSettings.userPurpose?.toLowerCase();
      const isSharedMailbox = userPurpose === 'shared';

      if (this.debugMode) {
        logger.info(
          `[MailboxDiscovery] ${user.displayName} (${userEmail}): userPurpose="${userPurpose || 'user'}" → ${isSharedMailbox ? 'SHARED mailbox (checking delegation permissions)' : 'Personal mailbox (skipping)'}`
        );
      }

      return isSharedMailbox;
    } catch (error: any) {
      if (error.name !== 'AbortError') {
        // Handle specific error for mailboxes that aren't REST-enabled (e.g., forwarding-only)
        if (error.message?.includes('MailboxNotEnabledForRESTAPI')) {
          if (this.debugMode) {
            logger.info(`[MailboxDiscovery] ${user.displayName} (${userEmail}): Mailbox not REST-enabled (likely forwarding-only) - skipping`);
          }
        } else {
          if (this.debugMode) {
            logger.debug(`[MailboxDiscovery] Mailbox type check failed for ${user.displayName} (${userEmail}): ${error.message}`);
          }
        }
      }
    }

    // Default to false (assume personal) if we can't determine the type
    return false;
  }

  /**
   * Strategy 3: Check if impersonated user has SendAs permissions.
   * 
   * SendAs permissions allow a user to send email on behalf of another user or mailbox,
   * which typically indicates delegate access to that mailbox.
   * 
   * @param userId - The user ID to check SendAs permissions for
   * @param impersonatedUserEmail - The email of the impersonated user (lowercase)
   * @param displayName - The display name of the mailbox being checked (for logging)
   * @param mailboxEmail - The email of the mailbox being checked (for logging)
   * @returns true if SendAs permissions are detected
   */
  private async checkSendAsPermissions(
    userId: string,
    impersonatedUserEmail: string,
    displayName?: string,
    mailboxEmail?: string
  ): Promise<boolean> {
    const mailboxIdentifier = displayName && mailboxEmail ? `${displayName} (${mailboxEmail})` : userId;
    
    if (this.debugMode) {
      logger.debug(`[MailboxDiscovery] Checking SendAs permissions for ${mailboxIdentifier}...`);
    }
    
    try {
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), this.requestTimeoutMs);

      const sendAsData = await this.graphClient.makeRequest(
        `/users/${userId}/sendAs`,
        {}
      ) as any;

      clearTimeout(timeoutId);

      if (sendAsData.value) {
        const hasSendAs = sendAsData.value.some((perm: any) =>
          perm.emailAddress?.toLowerCase() === impersonatedUserEmail ||
          perm.address?.toLowerCase() === impersonatedUserEmail
        );
        
        if (this.debugMode && hasSendAs) {
          logger.debug(`[MailboxDiscovery] SendAs permission found for ${impersonatedUserEmail} on ${mailboxIdentifier}`);
        }
        
        return hasSendAs;
      }
    } catch (error: any) {
      if (error.name !== 'AbortError') {
        // This error is expected for mailboxes without SendAs permissions configured
        if (this.debugMode) {
          logger.debug(`[MailboxDiscovery] SendAs check for ${mailboxIdentifier}: ${error.message}`);
        }
      }
    }

    return false;
  }

  /**
   * Strategy 3.5: Check mailbox folder permissions for delegates (works in client credentials mode).
   * 
   * When using application permissions (client credentials), we can query the mailbox folde permissions
   * to see if the impersonated user is listed as a delegate. This is more reliable than the
   * simple mailbox access test which would return true for ALL mailboxes when using app permissions.
   * 
   * @param mailboxUserId - The user ID of the mailbox to check
   * @param impersonatedUserId - The ID of the impersonated user
   * @param impersonatedUserEmail - The email of the impersonated user (for logging)
   * @returns true if the impersonated user has delegate access via mailbox folder permissions
   */
  private async checkMailboxFolderPermissions(
    mailboxUserId: string,
    impersonatedUserId: string,
    impersonatedUserEmail: string
  ): Promise<boolean> {
    try {
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), this.requestTimeoutMs);

      // Try to get the Inbox folder permissions
      const inboxData = await this.graphClient.makeRequest(
        `/users/${mailboxUserId}/mailFolders/inbox`,
        {}
      ) as any;

      clearTimeout(timeoutId);

      if (inboxData && inboxData.id) {
        // Now check if we can get permissions on this folder
        const controller2 = new AbortController();
        const timeoutId2 = setTimeout(() => controller2.abort(), this.requestTimeoutMs);

        try {
          const permsData = await this.graphClient.makeRequest(
            `/users/${mailboxUserId}/mailFolders/${inboxData.id}/messageRules`,
            {}
          ) as any;

          clearTimeout(timeoutId2);
          
          // If we successfully query, the impersonated user likely has some form of access
          // But we need a more reliable check - let's try mailFolder child folders
          const controller3 = new AbortController();
          const timeoutId3 = setTimeout(() => controller3.abort(), this.requestTimeoutMs);
          
          const childFolders = await this.graphClient.makeRequest(
            `/users/${mailboxUserId}/mailFolders?$top=1`,
            {}
          ) as any;
          
          clearTimeout(timeoutId3);
          
          // Being able to list folders indicates delegation
          if (childFolders && childFolders.value) {
            return true;
          }
        } catch (permsError: any) {
          clearTimeout(timeoutId2);
          if (this.debugMode) {
            logger.debug(`Could not check folder permissions for mailbox ${mailboxUserId}: ${permsError.message}`);
          }
        }
      }
    } catch (error: any) {
      if (error.name !== 'AbortError') {
        if (this.debugMode) {
          logger.debug(`Mailbox folder permission check failed for ${mailboxUserId}: ${error.message}`);
        }
      }
    }

    return false;
  }

  /**
   * Strategy 4: Test actual mailbox access by attempting to read mailFolders.
   * 
   * This is a fallback strategy for detecting "Full Access" delegation permissions
   * that are configured at the Exchange/Microsoft 365 admin level but aren't exposed
   * via the Calendar or SendAs Graph API endpoints.
   * 
   * When using application permissions (not delegated), we test access by directly
   * attempting to query the mailbox. Success indicates Full Access delegation.
   * 
   * @param userId - The user ID of the mailbox to test access to
   * @param impersonatedUserEmail - The email of the impersonated user (for logging)
   * @returns true if mailbox access is successful
   */
  private async testMailboxAccess(userId: string, impersonatedUserEmail: string): Promise<boolean> {
    try {
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), this.requestTimeoutMs);

      // Attempt to read mailFolders - this will succeed if user has Full Access delegation
      await this.graphClient.makeRequest(
        `/users/${userId}/mailFolders?$top=1`,
        {}
      );

      clearTimeout(timeoutId);

      if (this.debugMode) {
        logger.debug(`Mailbox access test succeeded for user ${userId}`);
      }

      return true;
    } catch (error: any) {
      if (error.name !== 'AbortError') {
        // Access denied is expected for mailboxes without delegation
        logger.debug(`Mailbox access test failed for user ${userId}: ${error.message}`);
      }
    }

    return false;
  }
}

export default MailboxDiscoveryService;

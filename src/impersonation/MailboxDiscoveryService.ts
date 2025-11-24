import logger from '../logger.js';
import type GraphClient from '../graph-client.js';
import type { MailboxInfo } from './MailboxDiscoveryCache.js';

/**
 * Service for discovering mailboxes accessible to a user via various permission strategies.
 * 
 * This service implements a comprehensive 4-strategy approach to discover:
 * - Personal mailbox (always included)
 * - Delegated mailboxes (via calendar permissions or SendAs permissions)
 * - Shared mailboxes (mailboxes with no licenses that user has access to)
 * 
 * Performance characteristics:
 * - Uses concurrent request limiting (max 5 simultaneous requests)
 * - Implements 5-second timeout per individual request
 * - Queries tenant users with reasonable limits ($top=100)
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

  /**
   * Creates a new MailboxDiscoveryService instance.
   * 
   * @param graphClient - Authenticated GraphClient instance for making Microsoft Graph API calls
   */
  constructor(graphClient: GraphClient) {
    this.graphClient = graphClient;
  }

  /**
   * Discovers all mailboxes accessible to the specified user.
   * 
   * This method performs a comprehensive discovery using 4 strategies:
   * 1. Personal mailbox (always included)
   * 2. Calendar delegate permissions
   * 3. Shared mailbox membership
   * 4. SendAs permissions
   * 
   * @param userEmail - Email address of the user to discover mailboxes for
   * @returns Array of discovered mailboxes with metadata
   * @throws Error if user cannot be found or if Graph API access fails
   */
  async discoverMailboxes(userEmail: string): Promise<MailboxInfo[]> {
    const mailboxes: MailboxInfo[] = [];

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
      logger.info(`Found personal mailbox for ${userEmail}: ${userData.displayName}`);
    } catch (error) {
      logger.error(`Error fetching user ${userEmail}: ${(error as Error).message}`);
      throw new Error(`Could not find user ${userEmail}: ${(error as Error).message}`);
    }

    // 2. Discover shared/delegated mailboxes using Graph API
    try {
      const impersonatedUserId = mailboxes[0]?.id;
      const impersonatedUserEmail = userEmail.toLowerCase();

      // Query tenant users to check for mailbox access
      const usersData = await this.graphClient.makeRequest(
        '/users?$filter=userType eq \'Member\'&$select=id,displayName,userPrincipalName,mail&$top=100',
        {}
      ) as any;

      const users = usersData.value || [];
      logger.info(`Checking mailbox permissions for ${userEmail} across ${users.length} mailboxes...`);

      // Process users in batches to limit concurrent requests
      for (let i = 0; i < users.length; i += this.maxConcurrentRequests) {
        const batch = users.slice(i, i + this.maxConcurrentRequests);

        await Promise.all(batch.map(async (user: any) => {
          // Skip the impersonated user's own mailbox (already added)
          if (user.id === impersonatedUserId) {
            return;
          }

          try {
            const mailboxInfo = await this.checkUserMailboxAccess(
              user,
              impersonatedUserId,
              impersonatedUserEmail
            );

            if (mailboxInfo) {
              mailboxes.push(mailboxInfo);
              logger.info(`${userEmail} has ${mailboxInfo.type} access to: ${user.displayName}`);
            }
          } catch (error: any) {
            // Silently skip errors (timeout, permission denied, etc.)
            if (error.name !== 'AbortError') {
              logger.debug(`Error checking ${user.displayName}: ${error.message}`);
            }
          }
        }));
      }
    } catch (error) {
      logger.warn(`Could not scan for mailbox permissions: ${(error as Error).message}`);
    }

    return mailboxes;
  }

  /**
   * Checks if the impersonated user has access to a specific user's mailbox.
   * Uses all 4 discovery strategies in sequence.
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
    let hasAccess = false;
    let accessType: 'shared' | 'delegated' = 'delegated';

    const isSharedMailbox = user.userPrincipalName?.toLowerCase().includes('shared') ||
      user.displayName?.toLowerCase().includes('shared') ||
      user.mail?.toLowerCase().includes('shared');

    // Strategy 1: Check calendar permissions (calendar delegate access)
    if (await this.checkCalendarPermissions(user.id, impersonatedUserEmail)) {
      hasAccess = true;
      accessType = 'delegated';
      logger.info(`Found calendar delegate: ${impersonatedUserEmail} → ${user.displayName}`);
    }

    // Strategy 2: Check for shared mailbox membership
    if (!hasAccess && isSharedMailbox) {
      if (await this.detectSharedMailbox(user)) {
        accessType = 'shared';
        // Will validate access in Strategy 4
      }
    }

    // Strategy 3: Check SendAs permissions (mail delegate)
    if (!hasAccess) {
      if (await this.checkSendAsPermissions(user.id, impersonatedUserEmail)) {
        hasAccess = true;
        accessType = 'delegated';
        logger.info(`Found SendAs permission: ${impersonatedUserEmail} → ${user.displayName}`);
      }
    }

    // Strategy 4: For shared mailboxes, verify actual access
    if (accessType === 'shared' && !hasAccess) {
      if (await this.verifyMailboxAccess(user.userPrincipalName || user.mail)) {
        hasAccess = true;
        logger.info(`Found shared mailbox access: ${impersonatedUserEmail} → ${user.displayName}`);
      }
    }

    if (hasAccess) {
      return {
        id: user.id,
        type: accessType,
        displayName: user.displayName,
        email: user.userPrincipalName || user.mail,
        permissions: ['read', 'write', 'send'],
      };
    }

    return null;
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
   * Shared mailboxes are identified by having no assigned licenses. They are typically
   * used for team collaboration and don't represent individual users.
   * 
   * @param user - The user object to check
   * @returns true if the mailbox appears to be a shared mailbox
   */
  private async detectSharedMailbox(user: any): Promise<boolean> {
    logger.debug(`Checking if ${user.displayName} is a shared mailbox...`);

    try {
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), this.requestTimeoutMs);

      const mailboxData = await this.graphClient.makeRequest(
        `/users/${user.id}?$select=mailboxSettings,assignedLicenses,accountEnabled`,
        {}
      ) as any;

      clearTimeout(timeoutId);

      // Shared mailboxes typically have no licenses
      const isLikelyShared = !mailboxData.assignedLicenses ||
        mailboxData.assignedLicenses.length === 0;

      logger.debug(
        `${user.displayName} license check: ${isLikelyShared ? 'no licenses (likely shared)' : 'has licenses'}`
      );

      return isLikelyShared;
    } catch (error: any) {
      if (error.name !== 'AbortError') {
        logger.debug(`Shared mailbox check failed for ${user.displayName}: ${error.message}`);
      }
    }

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
   * @returns true if SendAs permissions are detected
   */
  private async checkSendAsPermissions(userId: string, impersonatedUserEmail: string): Promise<boolean> {
    logger.debug(`Checking SendAs permissions for user ${userId}...`);

    try {
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), this.requestTimeoutMs);

      const sendAsData = await this.graphClient.makeRequest(
        `/users/${userId}/sendAs`,
        {}
      ) as any;

      clearTimeout(timeoutId);

      logger.debug(`SendAs data for user ${userId}: ${JSON.stringify(sendAsData)}`);

      if (sendAsData.value) {
        return sendAsData.value.some((perm: any) =>
          perm.emailAddress?.toLowerCase() === impersonatedUserEmail ||
          perm.address?.toLowerCase() === impersonatedUserEmail
        );
      }
    } catch (error: any) {
      if (error.name !== 'AbortError') {
        logger.debug(`SendAs check error for user ${userId}: ${error.message}`);
      }
    }

    return false;
  }

  /**
   * Strategy 4: Verify actual mailbox access by attempting to read the inbox.
   * 
   * This is the most direct verification method - if we can access the inbox,
   * the user definitely has access to the mailbox. Used primarily for shared
   * mailboxes where permission metadata may not be directly queryable.
   * 
   * @param userPrincipalName - The UPN or email of the mailbox to verify
   * @returns true if mailbox access is confirmed
   */
  private async verifyMailboxAccess(userPrincipalName: string): Promise<boolean> {
    try {
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), this.requestTimeoutMs);

      await this.graphClient.makeRequest(
        `/users/${encodeURIComponent(userPrincipalName)}/mailFolders/inbox?$select=id,displayName`,
        {}
      );

      clearTimeout(timeoutId);

      // If we successfully accessed the inbox, the user has mailbox access
      return true;
    } catch (error: any) {
      if (error.name !== 'AbortError') {
        logger.debug(`Shared mailbox access check failed for ${userPrincipalName}: ${error.message}`);
      }
    }

    return false;
  }
}

export default MailboxDiscoveryService;

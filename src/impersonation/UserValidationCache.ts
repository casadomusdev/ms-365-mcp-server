import logger from '../logger.js';
import type GraphClient from '../graph-client.js';

/**
 * Cache entry for user validation results.
 */
interface ValidationCacheEntry {
  email: string;
  exists: boolean;
  validatedAt: number;
  expiresAt: number;
}

/**
 * Optional user validation service with TTL-based caching.
 * 
 * This service validates that impersonated users exist in the tenant before
 * attempting mailbox discovery. It provides two levels of validation:
 * 
 * 1. **Format validation** (always enabled, instant)
 *    - Validates email format using regex
 *    - No API calls required
 * 
 * 2. **Existence validation** (optional, cached)
 *    - Queries Microsoft Graph API to verify user exists
 *    - Results cached for TTL duration to minimize API calls
 * 
 * Trade-offs:
 * - **Disabled**: Maximum performance, validation happens at Graph API call time
 * - **Enabled**: Fail-fast with clear errors, better UX, slight performance cost on first request
 * 
 * Environment variables:
 * - MS365_MCP_VALIDATE_IMPERSONATION_USER: Enable/disable validation (default: false)
 * - MS365_MCP_USER_VALIDATION_TTL: Cache TTL in seconds (default: 3600)
 * 
 * @example
 * ```typescript
 * const cache = new UserValidationCache(graphClient, 3600000); // 1 hour TTL
 * 
 * try {
 *   await cache.validateUser('user@company.com');
 *   // User exists, proceed with operation
 * } catch (error) {
 *   // User doesn't exist or invalid format
 *   console.error(error.message);
 * }
 * ```
 */
export class UserValidationCache {
  private readonly graphClient: GraphClient;
  private readonly cache = new Map<string, ValidationCacheEntry>();
  private readonly ttlMs: number;

  /**
   * Creates a new UserValidationCache instance.
   * 
   * @param graphClient - Authenticated GraphClient instance for making Microsoft Graph API calls
   * @param ttlMs - Cache TTL in milliseconds (default: 3600000 = 1 hour)
   */
  constructor(graphClient: GraphClient, ttlMs: number = 3600000) {
    this.graphClient = graphClient;
    this.ttlMs = ttlMs;
  }

  /**
   * Validates that a user email exists in the tenant.
   * 
   * This method performs two-stage validation:
   * 1. Format validation (always, instant)
   * 2. Existence check via Graph API (with caching)
   * 
   * @param email - Email address to validate
   * @throws Error if email format is invalid
   * @throws Error if user does not exist in tenant
   * @returns Promise that resolves if user is valid
   */
  async validateUser(email: string): Promise<void> {
    // Stage 1: Format validation (instant, no API call)
    if (!this.isValidEmailFormat(email)) {
      throw new Error(`Invalid email format: ${email}`);
    }

    // Stage 2: Existence check (with caching)
    const exists = await this.checkUserExistsWithCache(email);
    
    if (!exists) {
      throw new Error(`User not found in tenant: ${email}`);
    }
  }

  /**
   * Checks if user exists in tenant, using cache when available.
   * 
   * @param email - Email address to check
   * @returns true if user exists, false otherwise
   */
  private async checkUserExistsWithCache(email: string): Promise<boolean> {
    const key = email.toLowerCase();
    const now = Date.now();

    // Check cache first
    const cached = this.cache.get(key);
    if (cached && cached.expiresAt > now) {
      logger.debug(`User validation cache hit for: ${email}`);
      return cached.exists;
    }

    // Cache miss or expired - check with Graph API
    logger.debug(`User validation cache miss for: ${email}, checking Graph API...`);
    const exists = await this.checkUserExists(email);

    // Store result in cache
    this.cache.set(key, {
      email,
      exists,
      validatedAt: now,
      expiresAt: now + this.ttlMs,
    });

    logger.info(`User validation result for ${email}: ${exists ? 'exists' : 'not found'}`);
    return exists;
  }

  /**
   * Validates email format using regex.
   * 
   * Checks for basic email structure: localpart@domain.tld
   * This is a permissive check that allows most valid email formats.
   * 
   * @param email - Email address to validate
   * @returns true if format is valid
   */
  private isValidEmailFormat(email: string): boolean {
    // Basic email regex - permissive to allow most valid formats
    // Matches: user@domain.com, user.name@subdomain.domain.com, etc.
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email);
  }

  /**
   * Checks if a user exists in the tenant via Microsoft Graph API.
   * 
   * This method queries the /users endpoint to verify user existence.
   * It does not check permissions or mailbox access - only existence.
   * 
   * @param email - Email address to check
   * @returns true if user exists in tenant, false otherwise
   */
  private async checkUserExists(email: string): Promise<boolean> {
    try {
      // Try to fetch user by userPrincipalName
      await this.graphClient.makeRequest(
        `/users/${encodeURIComponent(email)}?$select=id,userPrincipalName`,
        {}
      );
      
      // If we get here without error, user exists
      return true;
    } catch (error: any) {
      // Check if it's a "not found" error vs other errors
      if (error.message?.includes('404') || 
          error.message?.includes('not found') ||
          error.message?.includes('does not exist')) {
        logger.debug(`User not found in tenant: ${email}`);
        return false;
      }

      // For other errors (permissions, network, etc.), log and re-throw
      logger.error(`Error checking user existence for ${email}: ${error.message}`);
      throw new Error(`Failed to validate user ${email}: ${error.message}`);
    }
  }

  /**
   * Clears all cached validation results.
   * 
   * Useful for testing or when you need to force revalidation.
   */
  clearCache(): void {
    this.cache.clear();
    logger.info('User validation cache cleared');
  }

  /**
   * Gets current cache statistics.
   * 
   * @returns Object with cache size and entry details
   */
  getCacheStats(): { size: number; entries: Array<{ email: string; exists: boolean; age: number }> } {
    const now = Date.now();
    const entries = Array.from(this.cache.values()).map(entry => ({
      email: entry.email,
      exists: entry.exists,
      age: Math.round((now - entry.validatedAt) / 1000), // age in seconds
    }));

    return {
      size: this.cache.size,
      entries,
    };
  }
}

export default UserValidationCache;

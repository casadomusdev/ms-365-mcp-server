import logger from '../logger.js';
import { UserValidationCache } from './UserValidationCache.js';

/**
 * Result of impersonation resolution.
 */
export interface ImpersonationResult {
  /** The resolved email address */
  email: string;
  /** The source from which the email was resolved */
  source: 'meta-header' | 'http-context' | 'env-var';
}

/**
 * Service for resolving impersonated user from multiple sources with clear precedence.
 * 
 * This resolver implements a 3-tier precedence system for determining which user
 * to impersonate when making Microsoft Graph API calls:
 * 
 * **Precedence (highest to lowest):**
 * 1. HTTP Header (X-Impersonate-User from _meta.headers)
 * 2. AsyncLocalStorage Context (set by HTTP middleware)
 * 3. Environment Variable (MS365_MCP_IMPERSONATE_USER)
 * 
 * **Validation:**
 * - Always validates email format using regex
 * - Optionally validates user existence via Graph API (if UserValidationCache provided)
 * - Empty/whitespace-only values are skipped and next source is tried
 * 
 * **Error Handling:**
 * - Clear error messages indicating which source caused the error
 * - Throws on invalid format or non-existent user (if validation enabled)
 * - Throws if no valid source is found
 * 
 * @example
 * ```typescript
 * // Without validation
 * const resolver = new ImpersonationResolver();
 * const result = await resolver.resolveImpersonatedUser(
 *   { 'x-impersonate-user': 'user@company.com' },
 *   undefined,
 *   undefined
 * );
 * // result.email = 'user@company.com', result.source = 'meta-header'
 * 
 * // With validation
 * const validationCache = new UserValidationCache(graphClient);
 * const resolver = new ImpersonationResolver(validationCache);
 * await resolver.resolveImpersonatedUser(...); // throws if user doesn't exist
 * ```
 */
export class ImpersonationResolver {
  private readonly validationCache?: UserValidationCache;
  private readonly headerName: string;

  /**
   * Creates a new ImpersonationResolver instance.
   * 
   * @param validationCache - Optional UserValidationCache for user existence validation
   */
  constructor(validationCache?: UserValidationCache) {
    this.validationCache = validationCache;
    this.headerName = (process.env.MS365_MCP_IMPERSONATE_HEADER || 'X-Impersonate-User').toLowerCase();
  }

  /**
   * Resolves the impersonated user from multiple sources with defined precedence.
   * 
   * Sources are checked in order:
   * 1. metaHeaders (HTTP header)
   * 2. storageContext (AsyncLocalStorage)
   * 3. envVar (environment variable)
   * 
   * Empty or whitespace-only values are skipped. The first non-empty value
   * with valid format (and valid user, if validation enabled) is returned.
   * 
   * @param metaHeaders - Headers from _meta (e.g., { 'x-impersonate-user': 'user@company.com' })
   * @param storageContext - User from AsyncLocalStorage context
   * @param envVar - Value from MS365_MCP_IMPERSONATE_USER env var
   * @returns Resolved email and source
   * @throws Error if no valid source found or validation fails
   */
  async resolveImpersonatedUser(
    metaHeaders?: Record<string, string>,
    storageContext?: string,
    envVar?: string
  ): Promise<ImpersonationResult> {
    // Try sources in order of precedence
    
    // 1. HTTP Header (meta-header)
    const fromMetaHeaders = metaHeaders?.[this.headerName]?.trim();
    if (fromMetaHeaders) {
      try {
        await this.validateEmail(fromMetaHeaders, 'meta-header');
        logger.info(`Impersonation resolved from HTTP header: ${fromMetaHeaders}`);
        return {
          email: fromMetaHeaders,
          source: 'meta-header',
        };
      } catch (error) {
        // Add source context to error message
        const message = (error as Error).message;
        throw new Error(`Impersonation failed (source: HTTP header '${this.headerName}'): ${message}`);
      }
    }

    // 2. AsyncLocalStorage Context (http-context)
    const fromContext = storageContext?.trim();
    if (fromContext) {
      try {
        await this.validateEmail(fromContext, 'http-context');
        logger.info(`Impersonation resolved from AsyncLocalStorage context: ${fromContext}`);
        return {
          email: fromContext,
          source: 'http-context',
        };
      } catch (error) {
        // Add source context to error message
        const message = (error as Error).message;
        throw new Error(`Impersonation failed (source: AsyncLocalStorage context): ${message}`);
      }
    }

    // 3. Environment Variable (env-var)
    const fromEnv = envVar?.trim();
    if (fromEnv) {
      try {
        await this.validateEmail(fromEnv, 'env-var');
        logger.info(`Impersonation resolved from environment variable: ${fromEnv}`);
        return {
          email: fromEnv,
          source: 'env-var',
        };
      } catch (error) {
        // Add source context to error message
        const message = (error as Error).message;
        throw new Error(`Impersonation failed (source: environment variable 'MS365_MCP_IMPERSONATE_USER'): ${message}`);
      }
    }

    // No valid source found
    throw new Error(
      'Impersonation not configured: No valid user found in HTTP header, AsyncLocalStorage context, or MS365_MCP_IMPERSONATE_USER environment variable'
    );
  }

  /**
   * Validates an email address (format and optionally existence).
   * 
   * @param email - Email address to validate
   * @param source - Source of the email (for error messages)
   * @throws Error if email format is invalid or user doesn't exist (if validation enabled)
   */
  private async validateEmail(email: string, source: string): Promise<void> {
    // Always validate format
    if (!this.isValidEmailFormat(email)) {
      throw new Error(`Invalid email format: '${email}'`);
    }

    // Optionally validate existence
    if (this.validationCache) {
      await this.validationCache.validateUser(email);
    }
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
   * Gets the configured header name for impersonation.
   * 
   * @returns The header name (lowercase)
   */
  getHeaderName(): string {
    return this.headerName;
  }
}

export default ImpersonationResolver;

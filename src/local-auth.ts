import logger from './logger.js';
import { refreshAccessToken } from './lib/microsoft-auth.js';

interface LoginTestResult {
  success: boolean;
  message?: string;
  userData?: {
    displayName: string;
    userPrincipalName: string;
  };
}

/**
 * LocalAuthManager implements a refresh-token-based auth flow controlled by USE_LOCAL_AUTH.
 *
 * It exchanges a long-lived refresh token for short-lived access tokens on demand
 * and refreshes them automatically when expired.
 */
class LocalAuthManager {
  private tenantId: string;
  private clientId: string;
  private clientSecret: string;
  private refreshToken: string;
  private accessToken: string | null = null;
  private tokenExpiry: number | null = null;

  // Scopes parameter kept for API compatibility with existing AuthManager
  // but not required for refresh-token grant (scopes are bound to the refresh token).
  constructor(_scopes: string[] = []) {
    const tenantId =
      process.env.MAILER_GRAPH_TENANT_ID || process.env.MS365_MCP_TENANT_ID || 'common';
    const clientId =
      process.env.MAILER_GRAPH_CLIENT_ID || process.env.MS365_MCP_CLIENT_ID || '';
    const clientSecret =
      process.env.MAILER_GRAPH_CLIENT_SECRET || process.env.MS365_MCP_CLIENT_SECRET || '';
    const refreshToken =
      process.env.MAILER_GRAPH_REFRESH_TOKEN ||
      process.env.MAILER_OAUTH_REFRESH_TOKEN ||
      process.env.MS365_MCP_REFRESH_TOKEN ||
      '';

    if (!clientId) {
      throw new Error('Local auth: client ID not configured');
    }
    if (!clientSecret) {
      throw new Error('Local auth: client secret not configured');
    }
    if (!refreshToken) {
      throw new Error('Local auth: refresh token not configured');
    }

    this.tenantId = tenantId;
    this.clientId = clientId;
    this.clientSecret = clientSecret;
    this.refreshToken = refreshToken;
  }

  async loadTokenCache(): Promise<void> {
    // No-op for local auth; tokens are derived from the stored refresh token.
    return;
  }

  async setOAuthToken(token: string): Promise<void> {
    // Allow setting an access token directly (used by OAuth provider verify hook if wired).
    this.accessToken = token;
    // Unknown expiry; force refresh after 5 minutes as a safety margin.
    this.tokenExpiry = Date.now() + 5 * 60 * 1000;
  }

  async getToken(forceRefresh = false): Promise<string | null> {
    if (!forceRefresh && this.accessToken && this.tokenExpiry && this.tokenExpiry > Date.now()) {
      return this.accessToken;
    }

    await this.refreshAccessToken();
    return this.accessToken;
  }

  private async refreshAccessToken(): Promise<void> {
    logger.info('Local auth: refreshing Microsoft access token using refresh token');
    const response = await refreshAccessToken(
      this.refreshToken,
      this.clientId,
      this.clientSecret,
      this.tenantId
    );

    this.accessToken = response.access_token;
    // If a new refresh token is returned, store it
    if (response.refresh_token) {
      this.refreshToken = response.refresh_token;
    }

    // Set expiry with a small safety buffer of 60 seconds
    const expiresInSeconds = response.expires_in || 3600;
    this.tokenExpiry = Date.now() + (expiresInSeconds - 60) * 1000;
  }

  async testLogin(): Promise<LoginTestResult> {
    try {
      const token = await this.getToken();
      if (!token) {
        return { success: false, message: 'No access token available' };
      }

      const resp = await fetch('https://graph.microsoft.com/v1.0/me', {
        headers: { Authorization: `Bearer ${token}` },
      });

      if (!resp.ok) {
        const text = await resp.text();
        return { success: false, message: `Graph /me failed: ${resp.status} ${text}` };
      }

      const user = (await resp.json()) as { displayName: string; userPrincipalName: string };
      return {
        success: true,
        userData: {
          displayName: user.displayName,
          userPrincipalName: user.userPrincipalName,
        },
      };
    } catch (error) {
      return { success: false, message: (error as Error).message };
    }
  }
}

export default LocalAuthManager;



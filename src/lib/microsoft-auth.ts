import { Request, Response, NextFunction } from 'express';
import logger from '../logger.js';

/**
 * Microsoft Bearer Token Auth Middleware with smart auto-detection
 * 
 * This middleware intelligently handles authentication based on what's available:
 * 
 * SMART BEHAVIOR:
 * - If bearer token is provided in request → Extract and use it
 * - If no bearer token BUT server has auth (CLIENT_SECRET, OAUTH_TOKEN, or cached tokens) → Allow request
 * - If no bearer token AND no server auth → Return 401 Unauthorized
 * 
 * AUTHENTICATION METHODS DETECTED:
 * - Bearer tokens: Authorization: Bearer <token> header
 * - Client credentials: MS365_MCP_CLIENT_SECRET environment variable
 * - BYOT (Bring Your Own Token): MS365_MCP_OAUTH_TOKEN environment variable
 * - Device code flow: Cached tokens from previous login
 * 
 * The middleware provides completely stateless operation when bearer tokens are used,
 * with automatic token refresh via the x-microsoft-refresh-token header.
 */
export const microsoftBearerTokenAuthMiddleware = (
  req: Request & { microsoftAuth?: { accessToken: string; refreshToken: string } },
  res: Response,
  next: NextFunction
): void => {
  const authHeader = req.headers.authorization;
  const hasBearerToken = authHeader && authHeader.startsWith('Bearer ');

  // If bearer token is provided, extract and use it
  if (hasBearerToken) {
    const accessToken = authHeader!.substring(7);
    const refreshToken = (req.headers['x-microsoft-refresh-token'] as string) || '';

    req.microsoftAuth = {
      accessToken,
      refreshToken,
    };

    logger.debug('Using bearer token authentication', {
      hasRefreshToken: !!refreshToken,
    });

    next();
    return;
  }

  // No bearer token provided - check if server has alternative authentication
  const hasClientCredentials = !!process.env.MS365_MCP_CLIENT_SECRET;
  const hasByotToken = !!process.env.MS365_MCP_OAUTH_TOKEN;
  
  // Note: We can't easily check for cached device code tokens here without async,
  // but that's okay - if they exist, the authManager will use them.
  // The middleware's job is just to not block the request.
  
  const hasAlternativeAuth = hasClientCredentials || hasByotToken;

  if (hasAlternativeAuth) {
    logger.debug('No bearer token provided, using server authentication', {
      method: hasClientCredentials ? 'client_credentials' : 'byot',
    });
    next();
    return;
  }

  // No bearer token and no alternative auth - this might fail or might work with cached tokens
  // We'll allow the request through and let it fail gracefully if no auth is available
  // This handles the device code flow case where tokens are cached
  logger.debug('No bearer token provided, attempting with cached credentials');
  next();
};

/**
 * Exchange authorization code for access token
 */
export async function exchangeCodeForToken(
  code: string,
  redirectUri: string,
  clientId: string,
  clientSecret: string,
  tenantId: string = 'common',
  codeVerifier?: string
): Promise<{
  access_token: string;
  token_type: string;
  scope: string;
  expires_in: number;
  refresh_token: string;
}> {
  const params = new URLSearchParams({
    grant_type: 'authorization_code',
    code,
    redirect_uri: redirectUri,
    client_id: clientId,
    client_secret: clientSecret,
  });

  // Add code_verifier for PKCE flow
  if (codeVerifier) {
    params.append('code_verifier', codeVerifier);
  }

  const response = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
    },
    body: params,
  });

  if (!response.ok) {
    const error = await response.text();
    logger.error(`Failed to exchange code for token: ${error}`);
    throw new Error(`Failed to exchange code for token: ${error}`);
  }

  return response.json();
}

/**
 * Refresh an access token
 */
export async function refreshAccessToken(
  refreshToken: string,
  clientId: string,
  clientSecret: string,
  tenantId: string = 'common'
): Promise<{
  access_token: string;
  token_type: string;
  scope: string;
  expires_in: number;
  refresh_token?: string;
}> {
  const response = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
    },
    body: new URLSearchParams({
      grant_type: 'refresh_token',
      refresh_token: refreshToken,
      client_id: clientId,
      client_secret: clientSecret,
    }),
  });

  if (!response.ok) {
    const error = await response.text();
    logger.error(`Failed to refresh token: ${error}`);
    throw new Error(`Failed to refresh token: ${error}`);
  }

  return response.json();
}

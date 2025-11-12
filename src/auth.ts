import type { AccountInfo, Configuration } from '@azure/msal-node';
import { PublicClientApplication, ConfidentialClientApplication } from '@azure/msal-node';
import keytar from 'keytar';
import logger from './logger.js';
import fs, { existsSync, readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import path from 'path';

interface EndpointConfig {
  pathPattern: string;
  method: string;
  toolName: string;
  scopes?: string[];
  workScopes?: string[];
}

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const endpointsData = JSON.parse(
  readFileSync(path.join(__dirname, 'endpoints.json'), 'utf8')
) as EndpointConfig[];

const endpoints = {
  default: endpointsData,
};

const SERVICE_NAME = 'ms-365-mcp-server';
const TOKEN_CACHE_ACCOUNT = 'msal-token-cache';
const SELECTED_ACCOUNT_KEY = 'selected-account';
// Use TOKEN_CACHE_DIR env var if set (e.g., /app/data in Docker), otherwise use project root for backward compatibility
const FALLBACK_DIR = process.env.MS365_MCP_TOKEN_CACHE_DIR || path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const FALLBACK_PATH = path.join(FALLBACK_DIR, '.token-cache.json');
const SELECTED_ACCOUNT_PATH = path.join(FALLBACK_DIR, '.selected-account.json');

const DEFAULT_CONFIG: Configuration = {
  auth: {
    clientId: process.env.MS365_MCP_CLIENT_ID || '084a3e9f-a9f4-43f7-89f9-d229cf97853e',
    authority: `https://login.microsoftonline.com/${process.env.MS365_MCP_TENANT_ID || 'common'}`,
  },
};

interface ScopeHierarchy {
  [key: string]: string[];
}

const SCOPE_HIERARCHY: ScopeHierarchy = {
  'Mail.ReadWrite': ['Mail.Read'],
  'Calendars.ReadWrite': ['Calendars.Read'],
  'Files.ReadWrite': ['Files.Read'],
  'Tasks.ReadWrite': ['Tasks.Read'],
  'Contacts.ReadWrite': ['Contacts.Read'],
};

function buildScopesFromEndpoints(includeWorkAccountScopes: boolean = false): string[] {
  const scopesSet = new Set<string>();

  endpoints.default.forEach((endpoint) => {
    // Skip endpoints that only have workScopes if not in work mode
    if (!includeWorkAccountScopes && !endpoint.scopes && endpoint.workScopes) {
      return;
    }

    // Add regular scopes
    if (endpoint.scopes && Array.isArray(endpoint.scopes)) {
      endpoint.scopes.forEach((scope) => scopesSet.add(scope));
    }

    // Add workScopes if in work mode
    if (includeWorkAccountScopes && endpoint.workScopes && Array.isArray(endpoint.workScopes)) {
      endpoint.workScopes.forEach((scope) => scopesSet.add(scope));
    }
  });

  Object.entries(SCOPE_HIERARCHY).forEach(([higherScope, lowerScopes]) => {
    if (lowerScopes.every((scope) => scopesSet.has(scope))) {
      lowerScopes.forEach((scope) => scopesSet.delete(scope));
      scopesSet.add(higherScope);
    }
  });

  return Array.from(scopesSet);
}

interface LoginTestResult {
  success: boolean;
  message: string;
  userData?: {
    displayName: string;
    userPrincipalName: string;
  };
}

class AuthManager {
  private config: Configuration;
  private scopes: string[];
  private msalApp: PublicClientApplication | ConfidentialClientApplication;
  private accessToken: string | null;
  private tokenExpiry: number | null;
  private oauthToken: string | null;
  private isOAuthMode: boolean;
  private isClientCredentialsMode: boolean;
  private selectedAccountId: string | null;

  constructor(
    config: Configuration = DEFAULT_CONFIG,
    scopes: string[] = buildScopesFromEndpoints()
  ) {
    this.config = config;
    this.scopes = scopes;
    this.accessToken = null;
    this.tokenExpiry = null;
    this.selectedAccountId = null;

    // Check for OAuth token mode
    const oauthTokenFromEnv = process.env.MS365_MCP_OAUTH_TOKEN;
    this.oauthToken = oauthTokenFromEnv ?? null;
    this.isOAuthMode = oauthTokenFromEnv != null;

    // Check for client credentials mode (app permissions)
    const clientSecret = process.env.MS365_MCP_CLIENT_SECRET;
    this.isClientCredentialsMode = !this.isOAuthMode && !!clientSecret;

    // Initialize appropriate MSAL application based on authentication mode
    if (this.isClientCredentialsMode) {
      // Client credentials flow (app permissions) - requires client secret
      const confidentialConfig: Configuration = {
        auth: {
          clientId: this.config.auth.clientId,
          authority: this.config.auth.authority,
          clientSecret: clientSecret,
        },
      };
      this.msalApp = new ConfidentialClientApplication(confidentialConfig);
      // Convert scopes to application permission format
      this.scopes = ['https://graph.microsoft.com/.default'];
      logger.info('Initialized in CLIENT CREDENTIALS mode (application permissions)');
      logger.info(`Using scopes: ${this.scopes.join(', ')}`);
    } else {
      // Device code flow (delegated permissions) - for user context
      this.msalApp = new PublicClientApplication(this.config);
      logger.info('Initialized in DEVICE CODE mode (delegated permissions)');
      logger.info(`Using scopes: ${scopes.join(', ')}`);
    }
  }

  async loadTokenCache(): Promise<void> {
    try {
      let cacheData: string | undefined;

      // Force file-based cache if MS365_MCP_FORCE_FILE_CACHE env var is set
      const forceFileCache = process.env.MS365_MCP_FORCE_FILE_CACHE === 'true' || process.env.MS365_MCP_FORCE_FILE_CACHE === '1';

      if (!forceFileCache) {
        try {
          const cachedData = await keytar.getPassword(SERVICE_NAME, TOKEN_CACHE_ACCOUNT);
          if (cachedData) {
            cacheData = cachedData;
          }
        } catch (keytarError) {
          logger.warn(
            `Keychain access failed, falling back to file storage: ${(keytarError as Error).message}`
          );
        }
      }

      if (!cacheData && existsSync(FALLBACK_PATH)) {
        cacheData = readFileSync(FALLBACK_PATH, 'utf8');
      }

      if (cacheData) {
        this.msalApp.getTokenCache().deserialize(cacheData);
      }

      // Load selected account
      await this.loadSelectedAccount();
    } catch (error) {
      logger.error(`Error loading token cache: ${(error as Error).message}`);
    }
  }

  private async loadSelectedAccount(): Promise<void> {
    try {
      let selectedAccountData: string | undefined;

      try {
        const cachedData = await keytar.getPassword(SERVICE_NAME, SELECTED_ACCOUNT_KEY);
        if (cachedData) {
          selectedAccountData = cachedData;
        }
      } catch (keytarError) {
        logger.warn(
          `Keychain access failed for selected account, falling back to file storage: ${(keytarError as Error).message}`
        );
      }

      if (!selectedAccountData && existsSync(SELECTED_ACCOUNT_PATH)) {
        selectedAccountData = readFileSync(SELECTED_ACCOUNT_PATH, 'utf8');
      }

      if (selectedAccountData) {
        const parsed = JSON.parse(selectedAccountData);
        this.selectedAccountId = parsed.accountId;
        logger.info(`Loaded selected account: ${this.selectedAccountId}`);
      }
    } catch (error) {
      logger.error(`Error loading selected account: ${(error as Error).message}`);
    }
  }

  async saveTokenCache(): Promise<void> {
    try {
      const cacheData = this.msalApp.getTokenCache().serialize();

      // Force file-based cache if MS365_MCP_FORCE_FILE_CACHE env var is set
      const forceFileCache = process.env.MS365_MCP_FORCE_FILE_CACHE === 'true' || process.env.MS365_MCP_FORCE_FILE_CACHE === '1';

      if (forceFileCache) {
        // Skip keytar, write directly to file
        fs.writeFileSync(FALLBACK_PATH, cacheData);
        logger.info('Token cache saved to file (forced)');
      } else {
        try {
          await keytar.setPassword(SERVICE_NAME, TOKEN_CACHE_ACCOUNT, cacheData);
        } catch (keytarError) {
          logger.warn(
            `Keychain save failed, falling back to file storage: ${(keytarError as Error).message}`
          );

          fs.writeFileSync(FALLBACK_PATH, cacheData);
        }
      }
    } catch (error) {
      logger.error(`Error saving token cache: ${(error as Error).message}`);
    }
  }

  private async saveSelectedAccount(): Promise<void> {
    try {
      const selectedAccountData = JSON.stringify({ accountId: this.selectedAccountId });

      // Force file-based cache if FORCE_FILE_CACHE env var is set
      const forceFileCache = process.env.FORCE_FILE_CACHE === 'true' || process.env.FORCE_FILE_CACHE === '1';

      if (forceFileCache) {
        // Skip keytar, write directly to file
        fs.writeFileSync(SELECTED_ACCOUNT_PATH, selectedAccountData);
        logger.info('Selected account saved to file (forced)');
      } else {
        try {
          await keytar.setPassword(SERVICE_NAME, SELECTED_ACCOUNT_KEY, selectedAccountData);
        } catch (keytarError) {
          logger.warn(
            `Keychain save failed for selected account, falling back to file storage: ${(keytarError as Error).message}`
          );

          fs.writeFileSync(SELECTED_ACCOUNT_PATH, selectedAccountData);
        }
      }
    } catch (error) {
      logger.error(`Error saving selected account: ${(error as Error).message}`);
    }
  }

  async setOAuthToken(token: string): Promise<void> {
    this.oauthToken = token;
    this.isOAuthMode = true;
  }

  async getToken(forceRefresh = false): Promise<string | null> {
    if (this.isOAuthMode && this.oauthToken) {
      return this.oauthToken;
    }

    if (this.accessToken && this.tokenExpiry && this.tokenExpiry > Date.now() && !forceRefresh) {
      return this.accessToken;
    }

    // Client credentials flow - no user account needed
    if (this.isClientCredentialsMode) {
      try {
        const response = await (this.msalApp as ConfidentialClientApplication).acquireTokenByClientCredential({
          scopes: this.scopes,
        });
        
        if (response) {
          this.accessToken = response.accessToken;
          this.tokenExpiry = response.expiresOn ? new Date(response.expiresOn).getTime() : null;
          await this.saveTokenCache();
          return this.accessToken;
        }
        
        throw new Error('Client credentials token acquisition returned no response');
      } catch (error) {
        logger.error(`Client credentials token acquisition failed: ${(error as Error).message}`);
        throw new Error(`Client credentials token acquisition failed: ${(error as Error).message}`);
      }
    }

    // Device code flow - requires user account
    const currentAccount = await this.getCurrentAccount();

    if (currentAccount) {
      const silentRequest = {
        account: currentAccount,
        scopes: this.scopes,
      };

      try {
        const response = await this.msalApp.acquireTokenSilent(silentRequest);
        this.accessToken = response.accessToken;
        this.tokenExpiry = response.expiresOn ? new Date(response.expiresOn).getTime() : null;
        return this.accessToken;
      } catch {
        logger.error('Silent token acquisition failed');
        throw new Error('Silent token acquisition failed');
      }
    }

    throw new Error('No valid token found');
  }

  async getCurrentAccount(): Promise<AccountInfo | null> {
    const accounts = await this.msalApp.getTokenCache().getAllAccounts();

    if (accounts.length === 0) {
      return null;
    }

    // If a specific account is selected, find it
    if (this.selectedAccountId) {
      const selectedAccount = accounts.find(
        (account: AccountInfo) => account.homeAccountId === this.selectedAccountId
      );
      if (selectedAccount) {
        return selectedAccount;
      }
      logger.warn(
        `Selected account ${this.selectedAccountId} not found, falling back to first account`
      );
    }

    // Fall back to first account (backward compatibility)
    return accounts[0];
  }

  async acquireTokenByDeviceCode(hack?: (message: string) => void): Promise<string | null> {
    // Device code flow is only available with PublicClientApplication
    if (this.isClientCredentialsMode) {
      logger.error('Device code flow is not supported in client credentials mode');
      throw new Error('Device code flow is not supported in client credentials mode. Use client secret authentication instead.');
    }

    const deviceCodeRequest = {
      scopes: this.scopes,
      deviceCodeCallback: (response: { message: string }) => {
        const text = ['\n', response.message, '\n'].join('');
        if (hack) {
          hack(text + 'After login run the "verify login" command');
        } else {
          console.log(text);
        }
        logger.info('Device code login initiated');
      },
    };

    try {
      logger.info('Requesting device code...');
      logger.info(`Requesting scopes: ${this.scopes.join(', ')}`);
      const response = await (this.msalApp as PublicClientApplication).acquireTokenByDeviceCode(deviceCodeRequest);
      logger.info(`Granted scopes: ${response?.scopes?.join(', ') || 'none'}`);
      logger.info('Device code login successful');
      this.accessToken = response?.accessToken || null;
      this.tokenExpiry = response?.expiresOn ? new Date(response.expiresOn).getTime() : null;

      // Set the newly authenticated account as selected if no account is currently selected
      if (!this.selectedAccountId && response?.account) {
        this.selectedAccountId = response.account.homeAccountId;
        await this.saveSelectedAccount();
        logger.info(`Auto-selected new account: ${response.account.username}`);
      }

      await this.saveTokenCache();
      return this.accessToken;
    } catch (error) {
      logger.error(`Error in device code flow: ${(error as Error).message}`);
      throw error;
    }
  }

  async testLogin(): Promise<LoginTestResult> {
    try {
      logger.info('Testing login...');
      const token = await this.getToken();

      if (!token) {
        logger.error('Login test failed - no token received');
        return {
          success: false,
          message: 'Login failed - no token received',
        };
      }

      logger.info('Token retrieved successfully, testing Graph API access...');

      try {
        const response = await fetch('https://graph.microsoft.com/v1.0/me', {
          headers: {
            Authorization: `Bearer ${token}`,
          },
        });

        if (response.ok) {
          const userData = await response.json();
          logger.info('Graph API user data fetch successful');
          return {
            success: true,
            message: 'Login successful',
            userData: {
              displayName: userData.displayName,
              userPrincipalName: userData.userPrincipalName,
            },
          };
        } else {
          const errorText = await response.text();
          logger.error(`Graph API user data fetch failed: ${response.status} - ${errorText}`);
          return {
            success: false,
            message: `Login successful but Graph API access failed: ${response.status}`,
          };
        }
      } catch (graphError) {
        logger.error(`Error fetching user data: ${(graphError as Error).message}`);
        return {
          success: false,
          message: `Login successful but Graph API access failed: ${(graphError as Error).message}`,
        };
      }
    } catch (error) {
      logger.error(`Login test failed: ${(error as Error).message}`);
      return {
        success: false,
        message: `Login failed: ${(error as Error).message}`,
      };
    }
  }

  async logout(): Promise<boolean> {
    try {
      const accounts = await this.msalApp.getTokenCache().getAllAccounts();
      for (const account of accounts) {
        await this.msalApp.getTokenCache().removeAccount(account);
      }
      this.accessToken = null;
      this.tokenExpiry = null;
      this.selectedAccountId = null;

      try {
        await keytar.deletePassword(SERVICE_NAME, TOKEN_CACHE_ACCOUNT);
        await keytar.deletePassword(SERVICE_NAME, SELECTED_ACCOUNT_KEY);
      } catch (keytarError) {
        logger.warn(`Keychain deletion failed: ${(keytarError as Error).message}`);
      }

      if (fs.existsSync(FALLBACK_PATH)) {
        fs.unlinkSync(FALLBACK_PATH);
      }

      if (fs.existsSync(SELECTED_ACCOUNT_PATH)) {
        fs.unlinkSync(SELECTED_ACCOUNT_PATH);
      }

      return true;
    } catch (error) {
      logger.error(`Error during logout: ${(error as Error).message}`);
      throw error;
    }
  }

  // Multi-account support methods
  async listAccounts(): Promise<AccountInfo[]> {
    return await this.msalApp.getTokenCache().getAllAccounts();
  }

  async selectAccount(accountId: string): Promise<boolean> {
    const accounts = await this.listAccounts();
    const account = accounts.find((acc: AccountInfo) => acc.homeAccountId === accountId);

    if (!account) {
      logger.error(`Account with ID ${accountId} not found`);
      return false;
    }

    this.selectedAccountId = accountId;
    await this.saveSelectedAccount();

    // Clear cached tokens to force refresh with new account
    this.accessToken = null;
    this.tokenExpiry = null;

    logger.info(`Selected account: ${account.username} (${accountId})`);
    return true;
  }

  async removeAccount(accountId: string): Promise<boolean> {
    const accounts = await this.listAccounts();
    const account = accounts.find((acc: AccountInfo) => acc.homeAccountId === accountId);

    if (!account) {
      logger.error(`Account with ID ${accountId} not found`);
      return false;
    }

    try {
      await this.msalApp.getTokenCache().removeAccount(account);

      // If this was the selected account, clear the selection
      if (this.selectedAccountId === accountId) {
        this.selectedAccountId = null;
        await this.saveSelectedAccount();
        this.accessToken = null;
        this.tokenExpiry = null;
      }

      logger.info(`Removed account: ${account.username} (${accountId})`);
      return true;
    } catch (error) {
      logger.error(`Failed to remove account ${accountId}: ${(error as Error).message}`);
      return false;
    }
  }

  getSelectedAccountId(): string | null {
    return this.selectedAccountId;
  }

  async listMailboxes(): Promise<any> {
    try {
      const token = await this.getToken();
      if (!token) {
        throw new Error('No valid token found');
      }

      const mailboxes: any[] = [];

      // 1. Get the current user's personal mailbox
      try {
        const meResponse = await fetch('https://graph.microsoft.com/v1.0/me', {
          headers: {
            Authorization: `Bearer ${token}`,
          },
        });

        if (meResponse.ok) {
          const userData = await meResponse.json();
          mailboxes.push({
            id: userData.id,
            type: 'personal',
            displayName: userData.displayName,
            email: userData.userPrincipalName || userData.mail,
            isPrimary: true,
          });
          logger.info(`Found personal mailbox: ${userData.displayName}`);
        }
      } catch (error) {
        logger.error(`Error fetching personal mailbox: ${(error as Error).message}`);
      }

      // 2. Try to discover shared/delegated mailboxes using a more efficient approach
      // Instead of querying all users, we'll use the MailboxSettings endpoint
      try {
        // Try to get delegated mailboxes through mailbox settings
        const settingsResponse = await fetch(
          'https://graph.microsoft.com/v1.0/me/mailboxSettings',
          {
            headers: {
              Authorization: `Bearer ${token}`,
            },
          }
        );

        if (settingsResponse.ok) {
          const settings = await settingsResponse.json();
          if (settings.delegateMeetingMessageDeliveryOptions) {
            logger.info('User has delegate settings configured');
          }
        }
      } catch (error) {
        logger.debug(`Could not query mailbox settings: ${(error as Error).message}`);
      }

      // 3. Try to find shared mailboxes by querying for mailbox folders we have access to
      // This is more efficient than testing all users
      try {
        // Query for shared/delegated folders - limit to reasonable number
        const usersResponse = await fetch(
          'https://graph.microsoft.com/v1.0/users?$filter=userType eq \'Member\'&$select=id,displayName,userPrincipalName,mail&$top=50',
          {
            headers: {
              Authorization: `Bearer ${token}`,
            },
          }
        );

        if (usersResponse.ok) {
          const usersData = await usersResponse.json();
          const users = usersData.value || [];
          
          logger.info(`Checking access to ${users.length} potential mailboxes...`);

          // Limit concurrent requests to avoid overwhelming the API
          const maxConcurrent = 5;
          for (let i = 0; i < users.length; i += maxConcurrent) {
            const batch = users.slice(i, i + maxConcurrent);
            
            await Promise.all(batch.map(async (user: any) => {
              // Skip the current user (already added as personal mailbox)
              if (mailboxes.length > 0 && user.id === mailboxes[0].id) {
                return;
              }

              try {
                // Try to read mail folders to test access with a timeout
                const controller = new AbortController();
                const timeoutId = setTimeout(() => controller.abort(), 5000); // 5 second timeout
                
                const mailFoldersResponse = await fetch(
                  `https://graph.microsoft.com/v1.0/users/${user.id}/mailFolders?$top=1`,
                  {
                    headers: {
                      Authorization: `Bearer ${token}`,
                    },
                    signal: controller.signal,
                  }
                );

                clearTimeout(timeoutId);

                if (mailFoldersResponse.ok) {
                  // User has access to this mailbox
                  const isShared = user.userPrincipalName?.toLowerCase().includes('shared') || 
                                  user.displayName?.toLowerCase().includes('shared') ||
                                  user.mail?.toLowerCase().includes('shared');
                  
                  mailboxes.push({
                    id: user.id,
                    type: isShared ? 'shared' : 'delegated',
                    displayName: user.displayName,
                    email: user.userPrincipalName || user.mail,
                    isPrimary: false,
                  });
                  logger.info(`Found accessible mailbox: ${user.displayName}`);
                }
              } catch (error: any) {
                // Silently skip mailboxes we don't have access to or timeout
                if (error.name !== 'AbortError') {
                  logger.debug(`No access to mailbox ${user.displayName}: ${error.message}`);
                }
              }
            }));
          }
        } else {
          const errorText = await usersResponse.text();
          logger.warn(`Could not query users for shared mailboxes: ${usersResponse.status} - ${errorText}`);
        }
      } catch (error) {
        logger.warn(`Could not query for delegated/shared mailboxes: ${(error as Error).message}`);
      }

      return {
        success: true,
        mailboxes,
        note: mailboxes.length === 1 ? 'Only personal mailbox found. Shared/delegated mailboxes may require additional permissions (User.ReadBasic.All).' : undefined,
      };
    } catch (error) {
      logger.error(`Error listing mailboxes: ${(error as Error).message}`);
      return {
        success: false,
        error: (error as Error).message,
      };
    }
  }
}

export default AuthManager;
export { buildScopesFromEndpoints };

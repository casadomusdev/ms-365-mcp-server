#!/usr/bin/env node

import 'dotenv/config';
import { parseArgs } from './cli.js';
import logger from './logger.js';
import AuthManager, { buildScopesFromEndpoints } from './auth.js';
import MicrosoftGraphServer from './server.js';
import { version } from './version.js';

async function main(): Promise<void> {
  try {
    const args = parseArgs();

    // Log authentication mode at startup
    const clientSecret = process.env.MS365_MCP_CLIENT_SECRET;
    const oauthToken = process.env.MS365_MCP_OAUTH_TOKEN;
    
    if (oauthToken) {
      logger.info('═══════════════════════════════════════════════════════════');
      logger.info('Authentication Mode: OAUTH TOKEN (Manual)');
      logger.info('Using pre-configured OAuth token from environment');
      logger.info('═══════════════════════════════════════════════════════════');
    } else if (clientSecret) {
      logger.info('═══════════════════════════════════════════════════════════');
      logger.info('Authentication Mode: CLIENT CREDENTIALS (Application Permissions)');
      logger.info('No user login required - using app-level permissions');
      logger.info('Access: ALL mailboxes/calendars/files in tenant');
      logger.info('═══════════════════════════════════════════════════════════');
    } else {
      logger.info('═══════════════════════════════════════════════════════════');
      logger.info('Authentication Mode: DEVICE CODE (Delegated Permissions)');
      logger.info('User login required - using user-level permissions');
      logger.info('Access: Limited to user\'s permissions');
      logger.info('═══════════════════════════════════════════════════════════');
    }

    const includeWorkScopes = args.orgMode || false;
    if (includeWorkScopes) {
      logger.info('Organization mode enabled - including work account scopes');
    }

    const scopes = buildScopesFromEndpoints(includeWorkScopes);
    const authManager = new AuthManager(undefined, scopes);
    await authManager.loadTokenCache();

    if (args.login) {
      await authManager.acquireTokenByDeviceCode();
      logger.info('Login completed, testing connection with Graph API...');
      const result = await authManager.testLogin();
      console.log(JSON.stringify(result));
      process.exit(0);
    }

    if (args.verifyLogin) {
      logger.info('Verifying login...');
      const result = await authManager.testLogin();
      console.log(JSON.stringify(result));
      process.exit(0);
    }

    if (args.logout) {
      await authManager.logout();
      console.log(JSON.stringify({ message: 'Logged out successfully' }));
      process.exit(0);
    }

    if (args.listAccounts) {
      const accounts = await authManager.listAccounts();
      const selectedAccountId = authManager.getSelectedAccountId();
      const result = accounts.map((account) => ({
        id: account.homeAccountId,
        username: account.username,
        name: account.name,
        selected: account.homeAccountId === selectedAccountId,
      }));
      console.log(JSON.stringify({ accounts: result }));
      process.exit(0);
    }

    if (args.selectAccount) {
      const success = await authManager.selectAccount(args.selectAccount);
      if (success) {
        console.log(JSON.stringify({ message: `Selected account: ${args.selectAccount}` }));
      } else {
        console.log(JSON.stringify({ error: `Account not found: ${args.selectAccount}` }));
        process.exit(1);
      }
      process.exit(0);
    }

    if (args.removeAccount) {
      const success = await authManager.removeAccount(args.removeAccount);
      if (success) {
        console.log(JSON.stringify({ message: `Removed account: ${args.removeAccount}` }));
      } else {
        console.log(JSON.stringify({ error: `Account not found: ${args.removeAccount}` }));
        process.exit(1);
      }
      process.exit(0);
    }

    if (args.listMailboxes) {
      const result = await authManager.listMailboxes({
        bypassImpersonation: args.all,
        clearCache: args.clearCache,
      });
      console.log(JSON.stringify(result));
      process.exit(0);
    }

    const server = new MicrosoftGraphServer(authManager, args);
    await server.initialize(version);
    await server.start();
  } catch (error) {
    logger.error(`Startup error: ${error}`);
    process.exit(1);
  }
}

main();

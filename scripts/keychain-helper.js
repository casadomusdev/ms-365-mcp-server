#!/usr/bin/env node
/**
 * Keychain Helper Script
 * 
 * Provides command-line interface to read/write tokens from/to system keychain.
 * Used by auth-export-tokens.sh and auth-import-tokens.sh to bridge bash and keychain access.
 * 
 * Commands:
 *   export-from-keychain <output-dir>  - Export tokens from keychain to files
 *   import-to-keychain <input-dir>     - Import tokens from files to keychain
 */

import keytar from 'keytar';
import fs from 'fs';
import path from 'path';

const SERVICE_NAME = 'ms-365-mcp-server';
const TOKEN_CACHE_ACCOUNT = 'msal-token-cache';
const SELECTED_ACCOUNT_KEY = 'selected-account';

async function exportFromKeychain(outputDir) {
    try {
        // Create output directory if it doesn't exist
        if (!fs.existsSync(outputDir)) {
            fs.mkdirSync(outputDir, { recursive: true });
        }

        let exportedCount = 0;

        // Export token cache
        try {
            const tokenCache = await keytar.getPassword(SERVICE_NAME, TOKEN_CACHE_ACCOUNT);
            if (tokenCache) {
                const tokenCachePath = path.join(outputDir, '.token-cache.json');
                fs.writeFileSync(tokenCachePath, tokenCache);
                console.log(`✓ Exported token cache to ${tokenCachePath}`);
                exportedCount++;
            } else {
                console.log('ℹ No token cache found in keychain');
            }
        } catch (error) {
            console.error(`✗ Error reading token cache from keychain: ${error.message}`);
        }

        // Export selected account
        try {
            const selectedAccount = await keytar.getPassword(SERVICE_NAME, SELECTED_ACCOUNT_KEY);
            if (selectedAccount) {
                const selectedAccountPath = path.join(outputDir, '.selected-account.json');
                fs.writeFileSync(selectedAccountPath, selectedAccount);
                console.log(`✓ Exported selected account to ${selectedAccountPath}`);
                exportedCount++;
            } else {
                console.log('ℹ No selected account found in keychain');
            }
        } catch (error) {
            console.error(`✗ Error reading selected account from keychain: ${error.message}`);
        }

        if (exportedCount === 0) {
            console.error('✗ No tokens found in keychain to export');
            process.exit(2);
        }

        console.log(`\n✓ Successfully exported ${exportedCount} file(s) from keychain`);
        process.exit(0);
    } catch (error) {
        console.error(`✗ Export failed: ${error.message}`);
        process.exit(1);
    }
}

async function importToKeychain(inputDir) {
    try {
        const tokenCachePath = path.join(inputDir, '.token-cache.json');
        const selectedAccountPath = path.join(inputDir, '.selected-account.json');

        let importedCount = 0;

        // Import token cache
        if (fs.existsSync(tokenCachePath)) {
            try {
                const tokenCache = fs.readFileSync(tokenCachePath, 'utf8');
                await keytar.setPassword(SERVICE_NAME, TOKEN_CACHE_ACCOUNT, tokenCache);
                console.log('✓ Imported token cache to keychain');
                importedCount++;
            } catch (error) {
                console.error(`✗ Error writing token cache to keychain: ${error.message}`);
                process.exit(1);
            }
        } else {
            console.error(`✗ Token cache file not found: ${tokenCachePath}`);
            process.exit(2);
        }

        // Import selected account (optional)
        if (fs.existsSync(selectedAccountPath)) {
            try {
                const selectedAccount = fs.readFileSync(selectedAccountPath, 'utf8');
                await keytar.setPassword(SERVICE_NAME, SELECTED_ACCOUNT_KEY, selectedAccount);
                console.log('✓ Imported selected account to keychain');
                importedCount++;
            } catch (error) {
                console.error(`✗ Error writing selected account to keychain: ${error.message}`);
            }
        } else {
            console.log('ℹ No selected account file found (optional)');
        }

        console.log(`\n✓ Successfully imported ${importedCount} file(s) to keychain`);
        process.exit(0);
    } catch (error) {
        console.error(`✗ Import failed: ${error.message}`);
        process.exit(1);
    }
}

function showUsage() {
    console.log('Keychain Helper - MS-365 MCP Server');
    console.log('');
    console.log('Usage:');
    console.log('  keychain-helper.js export-from-keychain <output-dir>');
    console.log('  keychain-helper.js import-to-keychain <input-dir>');
    console.log('');
    console.log('Commands:');
    console.log('  export-from-keychain  Export tokens from system keychain to files');
    console.log('  import-to-keychain    Import tokens from files to system keychain');
    console.log('');
    process.exit(1);
}

// Main execution
const command = process.argv[2];
const directory = process.argv[3];

if (!command || !directory) {
    showUsage();
}

switch (command) {
    case 'export-from-keychain':
        exportFromKeychain(directory);
        break;
    case 'import-to-keychain':
        importToKeychain(directory);
        break;
    default:
        console.error(`Unknown command: ${command}`);
        showUsage();
}

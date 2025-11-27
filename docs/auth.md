# Authentication Management

This guide explains how to manage authentication for the MS-365 MCP server using the provided helper scripts.

## Overview

The MS-365 MCP server uses Microsoft's **Device Code Flow** for authentication. This allows you to authenticate from any device (laptop, phone, tablet) without requiring the server to be accessible from the internet.

Once authenticated, tokens are cached in a Docker volume and automatically refreshed, so you only need to authenticate once (until tokens expire after ~90 days of inactivity).

## Quick Start

```bash
# 1. Initial authentication
./auth-login.sh

# 2. Verify it worked
./auth-verify.sh

# 3. Start the server
./start.sh
```

**For the absolute fastest start**, see [QUICK_START.md](QUICK_START.md).

## Authentication Scripts

### auth-login.sh

**Purpose**: Initiates the device code authentication flow to log in to your Microsoft 365 account.

**Usage**:
```bash
./auth-login.sh
```

**What happens**:
1. Displays a URL (e.g., `https://microsoft.com/devicelogin`)
2. Shows a device code (e.g., `AB12-CD34`)
3. You visit the URL on any device and enter the code
4. Sign in with your Microsoft 365 account
5. Approve the requested permissions
6. Tokens are cached in the Docker volume

**Options**:
- `--force-file-cache` - Force tokens to be saved to files instead of system keychain (useful for token export)

**When to use**:
- First time setup
- After tokens have expired
- After running `auth-logout.sh`
- When switching to a different Microsoft 365 account
- **With `--force-file-cache`**: When you need to export tokens for transfer to another machine

**Example output**:
```
╔════════════════════════════════════════════════════════════════╗
║         MS-365 MCP Server - Device Code Login                  ║
╔════════════════════════════════════════════════════════════════╗

Starting device code authentication flow...

Instructions:
  1. A URL and code will appear below
  2. Visit the URL on ANY device (laptop, phone, etc.)
  3. Enter the code when prompted
  4. Sign in with your Microsoft 365 account
  5. Review and approve the permissions

════════════════════════════════════════════════════════════════

To sign in, use a web browser to open the page https://microsoft.com/devicelogin
and enter the code AB12-CD34 to authenticate.

✓ Login successful!

Tokens have been cached. You can now:
  - Run ./auth-verify.sh to verify authentication
  - Start the server with: docker compose up -d
```

---

### auth-verify.sh

**Purpose**: Verifies that authentication is working and displays information about the authenticated user.

**Usage**:
```bash
./auth-verify.sh
```

**What happens**:
- Checks if cached tokens are valid
- Attempts to retrieve user information from Microsoft Graph API
- Displays the authenticated user's name and email

**When to use**:
- After initial authentication to confirm it worked
- To check if tokens are still valid
- To see which account is currently authenticated
- Troubleshooting authentication issues

**Example output**:
```
MS-365 MCP Server - Verify Authentication
===========================================

✓ Authentication verified

Authenticated as:
  Name:  John Doe
  Email: john.doe@company.com
```

**Error output** (if not authenticated):
```
✗ Authentication failed

You need to run ./auth-login.sh first to authenticate.
```

---

### auth-logout.sh

**Purpose**: Clears all cached credentials and logs out.

**Usage**:
```bash
./auth-logout.sh
```

**What happens**:
- Prompts for confirmation (destructive operation)
- Clears all tokens from the cache
- You'll need to run `auth-login.sh` again to re-authenticate

**When to use**:
- Switching to a different Microsoft 365 account
- Security concerns (removing credentials)
- Troubleshooting authentication issues
- Before decommissioning the server

**Example output**:
```
MS-365 MCP Server - Logout
============================

WARNING: This will remove all cached credentials.
You will need to re-authenticate with ./auth-login.sh

Are you sure you want to logout? (y/N): y

Logging out...

✓ Logged out successfully

All credentials have been cleared.
Run ./auth-login.sh to authenticate again.
```

---


Found 1 account(s):

  → john.doe@company.com (John Doe) [SELECTED]

Legend:
  → = Currently selected account

To select a different account:
  docker compose run --rm ms365-mcp node dist/index.js --select-account <account-id>
```

**Example output** (no accounts):
```
MS-365 MCP Server - List Accounts

No accounts found.

Run ./auth-login.sh to authenticate.
```
### auth-list-accounts.sh

**Purpose**: Lists all authenticated accounts and allows interactive account selection.

**Usage**:
```bash
./auth-list-accounts.sh
```

**What happens**:
- Displays all cached Microsoft 365 accounts with numbers
- Shows which account is currently selected
- **Interactively prompts** to select a different account (if multiple exist)
- Automatically switches to the selected account

**When to use**:
- Multi-account setups
- Checking which account is active
- Switching between accounts quickly

**Example output** (multiple accounts):
```
MS-365 MCP Server - List & Select Accounts

Found 3 account(s):

1. → John Doe (john@company.com) [SELECTED]
2.   Jane Smith (jane@company.com)
3.   Bob Johnson (bob@company.com)

Select an account to make it active:
Enter number (1-3) or press Enter to keep current: 2

Selecting account: Jane Smith (jane@company.com)

✓ Account selected successfully

Active account is now: Jane Smith
```

**Example output** (single account):
```
MS-365 MCP Server - List & Select Accounts

Found 1 account(s):

1. → John Doe (john@company.com) [SELECTED]

✓ Only one account available
```

**Example output** (no accounts):
```
MS-365 MCP Server - List & Select Accounts

No accounts found.

Run ./auth-login.sh to authenticate.
```
====================================

Found 1 account(s):

  → john.doe@company.com (John Doe) [SELECTED]

Legend:
  → = Currently selected account

To select a different account:
  docker compose run --rm ms365-mcp node dist/index.js --select-account <account-id>
```

**Example output** (no accounts):
```
MS-365 MCP Server - List Accounts
====================================

No accounts found.

Run ./auth-login.sh to authenticate.
```

---

## Multi-Account Support

The server supports multiple authenticated accounts. This is useful if you need to switch between different Microsoft 365 accounts.

### Adding Additional Accounts

```bash
# Authenticate with first account
./auth-login.sh

# Authenticate with second account (adds to cache)
./auth-login.sh

# View all accounts
./auth-list-accounts.sh
```

### Switching Between Accounts

```bash
# List accounts to get the account ID
./auth-list-accounts.sh

# Select a specific account
docker compose run --rm ms365-mcp node dist/index.js --select-account <account-id>

# Verify the switch
./auth-verify.sh
```

### Removing Specific Accounts

```bash
# List accounts to get the account ID
./auth-list-accounts.sh

# Remove a specific account
docker compose run --rm ms365-mcp node dist/index.js --remove-account <account-id>
```

---

## Troubleshooting

### "Authentication failed" when running health-check.sh

**Problem**: The health check is reporting authentication failure.

**Solution**:
```bash
# Verify authentication
./auth-verify.sh

# If it fails, re-authenticate
./auth-login.sh
```

### Tokens expired after inactivity

**Problem**: After ~90 days of no usage, refresh tokens expire.

**Solution**:
```bash
# Clear old tokens
./auth-logout.sh

# Re-authenticate
./auth-login.sh
```

### Wrong account authenticated

**Problem**: The server is using the wrong Microsoft 365 account.

**Solution**:
```bash
# Check which account is active
./auth-list-accounts.sh

# Option 1: Select a different account
docker compose run --rm ms365-mcp node dist/index.js --select-account <account-id>

# Option 2: Clear all and start fresh
./auth-logout.sh
./auth-login.sh
```

### Device code not working

**Problem**: The device code authentication flow isn't completing.

**Possible causes**:
- Firewall blocking access to `login.microsoftonline.com`
- Corporate proxy issues
- Expired device code (codes expire after 15 minutes)
- Wrong tenant/account

**Solution**:
```bash
# Try again (codes expire quickly)
./auth-login.sh

# Check Docker logs for errors
docker compose logs
```

---

## Security Best Practices

### Protecting Token Cache

The token cache contains sensitive authentication data:

```bash
# Backup tokens (encrypt the backup!)
docker run --rm -v ms365-mcp-token-cache:/data \
  -v $(pwd):/backup alpine \
  tar czf /backup/token-backup-$(date +%F).tar.gz -C /data .

# Encrypt the backup
gpg -c token-backup-*.tar.gz

# Store encrypted backup securely
rm token-backup-*.tar.gz  # Remove unencrypted version
```

### Regular Verification

Schedule regular authentication checks:

```bash
# Add to crontab for weekly verification
0 9 * * 1 cd /path/to/ms-365-mcp-server && ./auth-verify.sh || echo "Auth check failed" | mail -s "MS365 MCP Alert" admin@company.com
```

### Logout When Decommissioning

Always clear tokens before removing the server:

```bash
# Before removing server
./auth-logout.sh

# Then safe to remove
docker compose down -v
```

---

## Integration with Health Check

The `health-check.sh` script uses the same `--verify-login` functionality as `auth-verify.sh`. You can use them interchangeably:

```bash
# These are equivalent
./auth-verify.sh
./health-check.sh
docker compose exec ms365-mcp /app/health-check.sh
```

---

## Common Workflows

### Initial Setup

```bash
# 1. Configure environment
cp .env.example .env
vim .env  # Set MS365_MCP_CLIENT_ID, MS365_MCP_TENANT_ID

# 2. Build Docker image
docker compose build

# 3. Authenticate
./auth-login.sh

# 4. Verify
./auth-verify.sh

# 5. Start server
docker compose up -d

# 6. Check it's working
docker compose exec ms365-mcp /app/health-check.sh
```

### Daily Operations

```bash
# Check server status
docker compose ps

# Verify authentication
./auth-verify.sh

# View logs
docker compose logs -f
```

### Switching Accounts

```bash
# Method 1: Select from existing accounts
./auth-list-accounts.sh
docker compose run --rm ms365-mcp node dist/index.js --select-account <id>

# Method 2: Clear and re-authenticate
./auth-logout.sh
./auth-login.sh
```

### Troubleshooting

```bash
# 1. Check authentication
./auth-verify.sh

# 2. Check server logs
docker compose logs --tail 50

# 3. Re-authenticate if needed
./auth-logout.sh && ./auth-login.sh

# 4. Restart server
docker compose restart

# 5. Verify
docker compose exec ms365-mcp /app/health-check.sh
```

---

---

## Forcing File-Based Token Cache

By default, the system tries to use your operating system's secure keychain (macOS Keychain, Windows Credential Manager, etc.) to store tokens. However, when you need to export tokens for transfer to another machine, you need them in files instead.

### Using --force-file-cache Flag

```bash
# Force tokens to be saved to files (not system keychain)
./auth-login.sh --force-file-cache
```

**When to use this**:
- When you're authenticating on your local machine and need to export tokens
- Before transferring authentication to a production server
- When system keychain access is problematic

**What it does**:
- Sets `MS365_MCP_FORCE_FILE_CACHE=true` environment variable
- Skips system keychain (keytar) entirely
- Saves tokens directly to `.token-cache.json` and `.selected-account.json` files
- Makes tokens immediately available for export with `auth-export-tokens.sh`

**Example workflow**:
```bash
# Authenticate with file-based cache
./auth-login.sh --force-file-cache

# Verify it worked
./auth-verify.sh

# Export for transfer
./auth-export-tokens.sh

# Now you can transfer tokens-backup/ to another machine
```

---

## Token Transfer Between Machines

You can export authentication tokens from one machine and import them to another. This is useful for:
- Transferring authentication to production servers
- Backing up tokens
- Setting up multiple environments with the same credentials
- Avoiding repeated device code authentication
- Moving tokens between keychain and file-based storage

### Token Storage Modes

The server supports two token storage methods:

1. **System Keychain** (default for local/native installations)
   - Uses macOS Keychain, Windows Credential Manager, or Linux Secret Service
   - More secure - tokens encrypted by OS
   - Requires keychain access permissions
   - Cannot be easily transferred

2. **File-Based Storage** (default for Docker)
   - Tokens stored in `.token-cache.json` and `.selected-account.json` files
   - Less secure but portable
   - Required for Docker containers
   - Easy to backup/transfer

### Quick Token Transfer

```bash
# On source machine (Machine A) - auto-detects token location
./auth-export-tokens.sh
# Creates: tokens-20231211-143022.tar.gz

# Transfer securely to Machine B
scp tokens-*.tar.gz user@machine-b:/path/

# On destination machine (Machine B) 
# Import to keychain (default for local mode)
./auth-import-tokens.sh tokens-20231211-143022.tar.gz --to-keychain

# OR import to file-based storage
./auth-import-tokens.sh tokens-20231211-143022.tar.gz --to-file
```

### Detailed Token Export Process

**Export from ANY storage type** (keychain or file-based):

```bash
# Export tokens - automatically detects source (keychain or files)
./auth-export-tokens.sh
# Creates: tokens-YYYYMMDD-HHMMSS.tar.gz

# Or specify custom filename
./auth-export-tokens.sh my-backup.tar.gz

# The script will:
# 1. Check for tokens in system keychain (local mode)
# 2. Fall back to file-based tokens if keychain is empty
# 3. In Docker mode, always exports from container volume

# The archive includes:
# - .token-cache.json (MSAL token cache with access/refresh tokens)
# - .selected-account.json (currently selected account, if exists)
# - export-info.txt (metadata: date, hostname, export mode, source type)
```

**Export examples:**

```bash
# Example 1: Export from keychain (local machine)
./auth-export-tokens.sh
# Output: "Source: keychain"

# Example 2: Export from file-based storage
./auth-export-tokens.sh
# Output: "Source: file-based"

# Example 3: Export from Docker volume
./auth-export-tokens.sh
# Output: "Source: docker-volume"
```

### Securing Exported Tokens

**IMPORTANT**: Token files contain sensitive authentication data!

```bash
# Encrypt the tokens
cd tokens-backup/..
tar czf - tokens-backup | gpg -c > tokens-$(date +%Y%m%d).tar.gz.gpg

# Delete unencrypted files
rm -rf tokens-backup

# Transfer encrypted file securely
scp tokens-*.tar.gz.gpg user@server:/path/to/
```

### Detailed Token Import Process

**Import to ANY storage type** (keychain or file-based):

```bash
# Import to keychain (recommended for local installations)
./auth-import-tokens.sh tokens-20231211-143022.tar.gz --to-keychain

# Import to file-based storage
./auth-import-tokens.sh tokens-20231211-143022.tar.gz --to-file

# Auto-detect destination (keychain for local, file for Docker)
./auth-import-tokens.sh tokens-20231211-143022.tar.gz

# Verify authentication works
./auth-verify.sh
```

**Import behavior:**

- **Local mode (no Docker)**:
  - Default: Imports to keychain (`--to-keychain`)
  - Optional: Import to files with `--to-file`
  
- **Docker mode**:
  - Always imports to Docker volume (file-based)
  - Keychain not available in containers

**Import examples:**

```bash
# Example 1: Import to keychain on local machine
./auth-import-tokens.sh tokens-backup.tar.gz --to-keychain
# Output: "Location: keychain"

# Example 2: Import to files on local machine
./auth-import-tokens.sh tokens-backup.tar.gz --to-file
# Output: "Location: file-based"

# Example 3: Import to Docker container
./auth-import-tokens.sh tokens-backup.tar.gz
# Output: "Location: docker-volume"
```

### Converting Between Storage Types

You can convert tokens between keychain and file-based storage:

```bash
# Convert keychain tokens to files (for backup/transfer)
./auth-export-tokens.sh backup.tar.gz
./auth-import-tokens.sh backup.tar.gz --to-file

# Convert file-based tokens to keychain (more secure)
./auth-export-tokens.sh backup.tar.gz
./auth-import-tokens.sh backup.tar.gz --to-keychain

# This is useful when:
# - Moving from development (keychain) to Docker (files)
# - Moving from Docker (files) to production Mac (keychain)
# - Improving security by moving files to keychain
```

### Manual Token Transfer (Without Scripts)

If you prefer manual control or the scripts don't work:

**Export manually from Docker:**

```bash
# From source machine (Docker)
docker compose cp ms365-mcp:/app/data/.token-cache.json ./token-cache.json
docker compose cp ms365-mcp:/app/data/.selected-account.json ./selected-account.json

# Transfer files to destination machine
scp token-cache.json selected-account.json user@dest-server:/path/to/
```

**Export manually from keychain:**

```bash
# From source machine (local with keychain)
node scripts/keychain-helper.js export-from-keychain ./export-dir

# Transfer files to destination machine
scp export-dir/.token-cache.json export-dir/.selected-account.json user@dest-server:/path/to/
```

**Import manually to Docker:**

```bash
# On destination machine (Docker)
# Copy tokens to container
docker compose cp /path/to/token-cache.json ms365-mcp:/app/data/.token-cache.json
docker compose cp /path/to/selected-account.json ms365-mcp:/app/data/.selected-account.json

# Fix permissions
docker compose exec ms365-mcp chown node:node /app/data/.token-cache.json
docker compose exec ms365-mcp chown node:node /app/data/.selected-account.json

# Verify
./auth-verify.sh
```

**Import manually to keychain:**

```bash
# On destination machine (local with keychain)
# Create temp directory with token files
mkdir -p /tmp/token-import
cp /path/to/token-cache.json /tmp/token-import/.token-cache.json
cp /path/to/selected-account.json /tmp/token-import/.selected-account.json

# Import to keychain
node scripts/keychain-helper.js import-to-keychain /tmp/token-import

# Clean up
rm -rf /tmp/token-import

# Verify
./auth-verify.sh
```

### Token Storage Location

**System Keychain (Local mode, default):**
- macOS: Keychain Access app → "ms-365-mcp-server" entries
- Windows: Credential Manager → "ms-365-mcp-server"
- Linux: Secret Service (GNOME Keyring, KWallet, etc.)
- Files: No visible files, managed by OS
- Access: Via `keytar` Node.js library
- Service: `ms-365-mcp-server`
- Accounts: `msal-token-cache`, `selected-account`

**File-Based Storage (Docker or explicit):**
- Location in Docker: `/app/data/`
- Location locally: Project root directory
- Files: `.token-cache.json`, `.selected-account.json`
- Volume: Named Docker volume (`ms365-mcp-token-cache`)
- Controlled by: `MS365_MCP_TOKEN_CACHE_DIR` environment variable

**Environment Variables:**

```bash
# MS365_MCP_TOKEN_CACHE_DIR - where to store file-based tokens
# Default (project root for backward compatibility)
# Unset or empty = ./

# Docker (set in docker-compose.yaml)
MS365_MCP_TOKEN_CACHE_DIR=/app/data

# MS365_MCP_FORCE_FILE_CACHE - skip keychain, use files
# Default: not set (uses keychain if available)
# Set to 'true' or '1' to force file-based storage
MS365_MCP_FORCE_FILE_CACHE=true  # Use with ./auth-login.sh --force-file-cache
```

**Storage decision logic:**

1. **Docker mode**: Always uses file-based storage (`/app/data/`)
2. **Local mode with `FORCE_FILE_CACHE=true`**: Uses file-based storage
3. **Local mode (default)**: Attempts keychain, falls back to files if keychain unavailable

### Token Transfer Security Considerations

1. **Encrypt in transit**: Always encrypt tokens before transferring
2. **Secure channels**: Use scp, VPN, or encrypted USB drives
3. **Delete after transfer**: Remove unencrypted token files after successful import
4. **Verify permissions**: Ensure token files are only readable by the application user
5. **Audit trail**: Keep track of which machines have access to which accounts

### Troubleshooting Token Transfer

**"Container not found" error:**
```bash
# Make sure container exists
docker compose up -d
# Then retry import
./auth-import-tokens.sh
```

**"Permission denied" after import:**
```bash
# Fix permissions manually
docker compose exec ms365-mcp chown node:node /app/data/.token-cache.json
docker compose exec ms365-mcp chown node:node /app/data/.selected-account.json
```

**Tokens don't work after import:**
```bash
# Check storage location
# For Docker:
docker compose exec ms365-mcp ls -la /app/data/

# For keychain:
# macOS: Open Keychain Access, search for "ms-365-mcp-server"
# Or use the helper:
node scripts/keychain-helper.js export-from-keychain /tmp/check

# Verify token validity
./auth-verify.sh

# If expired, re-authenticate
./auth-login.sh
```

**Import to keychain fails:**
```bash
# Common cause: Keychain access denied
# Solution: Grant access when macOS prompts

# Alternative: Use file-based storage
./auth-import-tokens.sh tokens-backup.tar.gz --to-file
```

**Cannot export from keychain:**
```bash
# Check if tokens are actually in keychain
# macOS: Keychain Access → search "ms-365-mcp-server"

# If tokens are in files instead, export works automatically:
./auth-export-tokens.sh
# Will find and export file-based tokens

# Force file-based login for easier export:
./auth-login.sh --force-file-cache
```

---

## See Also

- [HEALTH_CHECK.md](HEALTH_CHECK.md) - Health check script documentation
- [SERVER_SETUP.md](SERVER_SETUP.md) - Complete server setup guide
- [docker-compose.yaml](docker-compose.yaml) - Docker configuration

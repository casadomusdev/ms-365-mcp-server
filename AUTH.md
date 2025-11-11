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
docker compose up -d
```

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

**When to use**:
- First time setup
- After tokens have expired
- After running `auth-logout.sh`
- When switching to a different Microsoft 365 account

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

### auth-list-accounts.sh

**Purpose**: Lists all authenticated accounts cached in the token store.

**Usage**:
```bash
./auth-list-accounts.sh
```

**What happens**:
- Displays all cached Microsoft 365 accounts
- Shows which account is currently selected
- Provides instructions for switching accounts

**When to use**:
- Multi-account setups
- Checking which account is active
- Before selecting a different account

**Example output** (single account):
```
MS-365 MCP Server - List Accounts
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

## See Also

- [HEALTH_CHECK.md](HEALTH_CHECK.md) - Health check script documentation
- [SERVER_SETUP.md](SERVER_SETUP.md) - Complete server setup guide
- [docker-compose.yaml](docker-compose.yaml) - Docker configuration

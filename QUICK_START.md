# MS-365 MCP Server - Quick Start Guide

Get up and running with the MS-365 MCP Server in minutes!

## Prerequisites

- **Azure AD Application** registered with appropriate permissions
- **Docker** (optional - can run locally without it)
- **Node.js 18+** (if running locally)

## 🚀 Quick Start (3 Steps)

### 1. Configure Environment

```bash
# Copy example configuration
cp .env.example .env

# Edit .env and add your Azure AD credentials
# Required:
#   MS365_MCP_CLIENT_ID      - Your Azure AD App Client ID
#   MS365_MCP_TENANT_ID      - Your tenant ID or 'common'
```

### 2. Start the Server

```bash
./start.sh
```

The script will:
- Validate your configuration
- Ask if you want Docker or Local mode (if Docker available)
- Set up and start the server automatically

### 3. Authenticate

```bash
./auth-login.sh
```

Follow the prompts to:
1. Choose token storage (system keychain or file-based)
2. Visit the provided URL
3. Enter the device code
4. Sign in with your Microsoft 365 account

**Done!** Your server is running and authenticated.

---

## 📋 All Available Scripts

### Server Management

#### `./start.sh [options]`
Start the MCP server with automatic setup.

**Options:**
```bash
./start.sh              # Interactive mode selection
./start.sh --docker     # Force Docker mode
./start.sh --local      # Force local Node.js mode
./start.sh --build      # Rebuild Docker image
```

**What it does:**
- Validates `.env` configuration
- Detects Docker availability
- Builds/starts server automatically
- Smart rebuilding (only when needed)
- **Handles restarts** (detects if already running)

#### `./stop.sh [options]`
Stop the MCP server (Docker or local).

**Options:**
```bash
./stop.sh               # Graceful stop
./stop.sh --force       # Force stop (docker down / kill -9)
```

**What it does:**
- Tries Docker first (docker compose stop)
- Falls back to local Node.js process if no Docker
- Graceful shutdown by default
- Force mode: removes containers (Docker) or kills process (local)

### Authentication

#### `./auth-login.sh [options]`
Authenticate with Microsoft 365.

**Options:**
```bash
./auth-login.sh                    # Interactive (choose storage type)
./auth-login.sh --force-file-cache # Use file-based cache
```

**Storage Options:**
1. **System Keychain** (default) - Secure, OS-managed
2. **File-based Cache** - Enables token export/import

#### `./auth-verify.sh`
Verify your authentication is working.

```bash
./auth-verify.sh

# Output:
# ✓ Authentication verified
# 
# Authenticated as:
#   Name:  John Doe
#   Email: john.doe@company.com
```

#### `./auth-logout.sh`
Log out and clear all cached credentials.

```bash
./auth-logout.sh
# Prompts for confirmation before clearing tokens
```

#### `./auth-list-accounts.sh`
List and switch between multiple authenticated accounts.

```bash
./auth-list-accounts.sh

# Interactive account selection:
# Found 3 account(s):
# 
# 1. → John Doe (john@company.com) [SELECTED]
# 2.   Jane Smith (jane@company.com)
# 3.   Bob Johnson (bob@company.com)
# 
# Select an account: _
```

### Token Management

#### `./auth-export-tokens.sh [filename]`
Export authentication tokens to a compressed archive.

```bash
./auth-export-tokens.sh
# Creates: tokens-20231211-143022.tar.gz

./auth-export-tokens.sh my-backup.tar.gz
# Creates: my-backup.tar.gz
```

**Includes:**
- Token cache file
- Export metadata (date, hostname, mode)
- Security warnings

**Note:** Only works with file-based token cache.

#### `./auth-import-tokens.sh <archive>`
Import authentication tokens from an archive.

```bash
./auth-import-tokens.sh tokens-20231211-143022.tar.gz

# Shows metadata and confirms import
```

### Health & Monitoring

#### `./health-check.sh`
Verify server health and Microsoft Graph API connectivity.

```bash
./health-check.sh

# Output:
# ✓ Authentication verified
# ✓ Microsoft Graph API connection successful
#
# Authenticated as:
#   Name:  John Doe
#   Email: john.doe@company.com
```

---

## 🔄 Common Workflows

### First Time Setup

```bash
# 1. Configure
cp .env.example .env
vim .env  # Add your credentials

# 2. Start (choose Docker or Local)
./start.sh

# 3. Authenticate
./auth-login.sh  # Choose file-based cache for portability

# 4. Verify
./auth-verify.sh
./health-check.sh
```

### Transfer Tokens to Another Machine

**On Machine A:**
```bash
./auth-export-tokens.sh
# Creates: tokens-YYYYMMDD-HHMMSS.tar.gz

# Transfer securely (scp, encrypted email, etc.)
scp tokens-*.tar.gz user@machine-b:/path/
```

**On Machine B:**
```bash
./auth-import-tokens.sh tokens-YYYYMMDD-HHMMSS.tar.gz
./auth-verify.sh
```

### Switch Between Multiple Accounts

```bash
# Login with first account
./auth-login.sh --force-file-cache
# (Complete authentication)

# Login with second account  
./auth-login.sh --force-file-cache
# (Complete authentication)

# Switch between them
./auth-list-accounts.sh
# Select account by number
```

### Daily Development

```bash
# Start server
./start.sh --local  # Or just ./start.sh

# Check auth status
./auth-verify.sh

# View logs (Docker mode)
docker compose logs -f

# Stop server
./stop.sh           # Graceful stop
./stop.sh --force   # Force stop & remove containers
```

---

## 🐳 Docker vs Local Mode

### Docker Mode (Recommended)

**Pros:**
- ✅ Isolated environment
- ✅ Consistent across machines
- ✅ Easy management (start/stop/restart)
- ✅ No Node.js version conflicts
- ✅ No port exposure needed

**Cons:**
- ⚠️ Requires Docker installed
- ⚠️ Slightly slower startup

**Usage:**

**For Claude Desktop Integration (or any local MCP client):**

This is the recommended setup for using the Dockerized MCP server with Claude Desktop or other MCP clients running on your local machine.

```bash
# 1. Configure environment
cp .env.example .env
# Edit .env with your Azure AD credentials:
#   MS365_MCP_CLIENT_ID - Your Azure AD App Client ID
#   MS365_MCP_TENANT_ID - Your tenant ID or 'common'

# 2. Start the Docker container
./start.sh --docker

# 3. Authenticate (one-time setup)
./auth-login.sh
# Follow the prompts:
# - Choose token storage (file-based recommended for Docker)
# - Visit the URL shown
# - Enter the device code
# - Sign in with your Microsoft 365 account

# 4. Verify authentication
./auth-verify.sh

# 5. Configure Claude Desktop (Settings > Developer)
# Use absolute path to docker-mcp-wrapper.sh
{
  "mcpServers": {
    "ms365": {
      "command": "/absolute/path/to/ms-365-mcp-server/docker-mcp-wrapper.sh",
      "args": ["--org-mode"]
    }
  }
}

# 6. Restart Claude Desktop to load the MCP server
```

**What the wrapper does:**
- Automatically starts the Docker container if not running
- Bridges STDIO communication between Claude and the container
- No manual `docker compose up` needed - fully automated
- Passes arguments like `--org-mode` to the MCP server

**Note:** Replace `/absolute/path/to/ms-365-mcp-server/` with the actual path, e.g., `/Users/rob/dev/projects/ms-365-mcp-server/`

**For Manual/Development Use:**
```bash
./start.sh --docker

# Management:
docker compose logs -f    # View logs
docker compose down       # Stop
docker compose restart    # Restart
```

**How Docker Integration Works:**
- The `docker-mcp-wrapper.sh` script bridges STDIO communication
- Automatically starts the container if not running
- No network ports need to be exposed
- Complete container isolation with secure communication

### Local Mode

**Pros:**
- ✅ Faster startup
- ✅ No Docker required
- ✅ Direct Node.js debugging

**Cons:**
- ⚠️ Requires Node.js 18+
- ⚠️ May have dependency conflicts

**Usage:**
```bash
./start.sh --local
# Server runs in foreground
# Press Ctrl+C to stop
```

---

## 🔍 Troubleshooting

### "Docker not available - using local mode"

Docker isn't running or installed. Either:
- Start Docker Desktop, or
- Continue with local mode (requires Node.js)

### "Authentication failed"

1. Check you've authenticated:
   ```bash
   ./auth-login.sh
   ```

2. Verify it worked:
   ```bash
   ./auth-verify.sh
   ```

3. Check token location:
   - System keychain: Managed by OS
   - File-based: `.token-cache.json` in project root

### "Token export failed - no tokens found"

Token export only works with file-based cache:
```bash
./auth-logout.sh             # Clear existing
./auth-login.sh --force-file-cache  # Re-auth with file cache
./auth-export-tokens.sh      # Now works
```

### Scripts showing "command not found"

Make scripts executable:
```bash
chmod +x *.sh
```

---

## 💡 Tips & Best Practices

### Security

- ✅ **Never commit** `.env` or token files to git (already in .gitignore)
- ✅ **Use file-based cache** only when you need token portability
- ✅ **System keychain** is more secure for daily use
- ✅ **Encrypt token archives** before transferring

### Performance

- ✅ **Use `--build` only when needed** - smart detection avoids unnecessary rebuilds
- ✅ **Local mode** is faster for rapid development
- ✅ **Docker mode** for production-like environment

### Multi-Account

- ✅ **File-based cache required** for multiple accounts
- ✅ **Use `auth-list-accounts.sh`** to switch easily
- ✅ **Export separately** for each account if needed

---

## 🔗 Related Documentation

- **[AUTH.md](AUTH.md)** - Detailed authentication guide
- **[HEALTH_CHECK.md](HEALTH_CHECK.md)** - Health check documentation
- **[README.md](README.md)** - Full project documentation
- **[.env.example](.env.example)** - Configuration template

---

## ⚡ Advanced Usage

### Running Without Scripts

**Docker:**
```bash
docker compose up -d
docker compose exec ms365-mcp node dist/index.js --verify-login
docker compose down
```

**Local:**
```bash
npm install
npm run build
node dist/index.js
```

### Custom Docker Project Name

```bash
# In .env:
COMPOSE_PROJECT_NAME=my-custom-name

# Scripts automatically use this name
./start.sh
```

### Automation

All scripts support non-interactive use:
```bash
# CI/CD example
./start.sh --docker --build
./auth-import-tokens.sh ./ci-tokens.tar.gz
./auth-verify.sh || exit 1
./health-check.sh || exit 1
```

---

**Need help?** Check the full documentation in [README.md](README.md) or open an issue.

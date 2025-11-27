# Server Setup Guide

This guide covers deploying the MS-365 MCP Server in Docker on a server without internet exposure.

## Overview

The server runs in Docker using STDIO mode with automatic token refresh:
- No inbound internet access required (no public IP, domain, or open ports)
- Only outbound HTTPS to Microsoft endpoints
- Client credentials authentication (service account for multi-user support)
- Optional device code flow for single-user/personal accounts
- Persistent token storage in Docker volume
- Automatic token refresh using MSAL

## Prerequisites

**Server:**
- Docker and Docker Compose installed
- Outbound HTTPS access to `login.microsoftonline.com` and `graph.microsoft.com`
- SSH access for initial setup

**Optional - PowerShell for Shared Mailbox Discovery:**
- PowerShell Core 7.x installed (for impersonation mode shared mailbox discovery)
- Exchange Online PowerShell Module v3.x
- **Auto-detected** by the server - gracefully falls back if not available
- See [POWERSHELL_SETUP.md](POWERSHELL_SETUP.md) for installation details

**Microsoft 365:**
- Microsoft 365 Business account
- Azure AD admin access for app registration

## Authentication Modes

Choose the authentication mode based on your use case:

### Client Credentials Flow (Default - Multi-User Support)

**When to use:**
- Multi-user deployments (server supports multiple users)
- Automated processes and background services
- Production deployments with service accounts
- When you need access to ALL mailboxes/calendars in the tenant

**How it works:**
- No interactive login required
- Uses `ConfidentialClientApplication` from MSAL
- Requires **Application Permissions** in Azure AD
- Access token represents the application itself
- Access to ALL resources the app has permission to access

**Configuration:**
- Set `MS365_MCP_CLIENT_SECRET` in `.env`
- Set `MS365_MCP_TENANT_ID` to your specific tenant ID (not "common")
- Use **Application Permissions** in Azure AD (see below)
- Grant admin consent for all permissions

### Device Code Flow (Optional - Single-User)

**When to use:**
- Personal Microsoft 365 accounts
- Single-user deployments (same user every time)
- When you want access limited to what a specific user can access
- Development and testing scenarios

**How it works:**
- Interactive user login required (one-time setup)
- Uses `PublicClientApplication` from MSAL
- Requires **Delegated Permissions** in Azure AD
- Access token represents the signed-in user
- Access limited to resources the user has permission to access

**Configuration:**
- Do NOT set `MS365_MCP_CLIENT_SECRET` in `.env`
- Run `./auth-login.sh` for initial authentication
- Use **Delegated Permissions** in Azure AD (see below)

## Azure AD App Registration

Microsoft 365 Business requires a custom Azure AD app.

### Basic Setup

1. Navigate to [Azure Portal](https://portal.azure.com) → **Azure Active Directory** → **App registrations**
2. Click **New registration**
3. Configure:
   - **Name**: `MS365 MCP Server`
   - **Supported account types**: "Accounts in this organizational directory only" (Single tenant)
   - **Redirect URI**: Leave blank
4. Click **Register** and note:
   - **Application (client) ID**
   - **Directory (tenant) ID**

### Create Client Secret (Required for Client Credentials Flow)

1. Go to **Certificates & secrets**
2. Click **New client secret**
3. Add description: "MS365 MCP Server"
4. Set expiration (recommend 24 months)
5. Click **Add**
6. **IMPORTANT**: Copy the secret value immediately (it won't be shown again)

### Configure Permissions

Choose the permission set based on your authentication mode:

#### Application Permissions (for Client Credentials Flow - Default)

1. Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Application permissions**
2. Add the following permissions:

**Mail:**
- `Mail.Read`
- `Mail.ReadWrite`
- `Mail.Send`

**Calendars:**
- `Calendars.Read`
- `Calendars.ReadWrite`

**Files:**
- `Files.Read.All`
- `Files.ReadWrite.All`

**User:**
- `User.Read.All`

**Mailbox Settings (Required for Impersonation & Shared Mailbox Discovery):**
- `MailboxSettings.Read`

**Tasks:**
- `Tasks.Read.All`
- `Tasks.ReadWrite.All`

**Contacts:**
- `Contacts.Read`
- `Contacts.ReadWrite`

**OneNote:**
- `Notes.Read.All`
- `Notes.ReadWrite.All`

**Teams:**
- `Chat.Read.All`
- `Chat.ReadWrite.All`
- `Team.ReadBasic.All`
- `Channel.ReadBasic.All`
- `ChannelMessage.Read.All`
- `TeamMember.Read.All`

**SharePoint:**
- `Sites.Read.All`
- `Sites.ReadWrite.All`

**People & Search:**
- `People.Read.All`
- `Organization.Read.All`

3. Click **Grant admin consent for [Your Organization]** (REQUIRED for application permissions)

#### Delegated Permissions (for Device Code Flow - Optional)

If using device code flow instead, configure these delegated permissions:

1. Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**
2. Add the following permissions:

**Core:**
- `User.Read`
- `User.Read.All`
- `User.ReadBasic.All`

**Mail (Personal & Shared):**
- `Mail.Read`
- `Mail.ReadWrite`
- `Mail.Send`
- `Mail.Read.Shared`
- `Mail.ReadWrite.Shared`
- `Mail.Send.Shared`

**Calendar:**
- `Calendars.Read`
- `Calendars.ReadWrite`
- `Calendars.Read.Shared`
- `Calendars.ReadWrite.Shared`

**Files:**
- `Files.Read`
- `Files.ReadWrite`
- `Files.Read.All`

**Tasks:**
- `Tasks.Read`
- `Tasks.ReadWrite`

**Contacts:**
- `Contacts.Read`
- `Contacts.ReadWrite`

**OneNote:**
- `Notes.Read`
- `Notes.Create`

**Teams:**
- `Chat.Read`
- `ChatMessage.Read`
- `ChatMessage.Send`
- `Team.ReadBasic.All`
- `Channel.ReadBasic.All`
- `ChannelMessage.Read.All`
- `ChannelMessage.Send`
- `TeamMember.Read.All`

**SharePoint:**
- `Sites.Read.All`

**People:**
- `People.Read`

3. Click **Grant admin consent for [Your Organization]**
4. Go to **Authentication** → **Advanced settings**
5. Set **Allow public client flows** to **Yes**
6. Click **Save**

## Installation

### 1. Deploy Files to Server

```bash
ssh user@your-server
sudo mkdir -p /opt/ms-365-mcp-server
cd /opt/ms-365-mcp-server

# Clone or copy project files
git clone <your-repo> .
# or: scp -r /local/path/* user@server:/opt/ms-365-mcp-server/
```

### 2. Configure Environment

Create `.env` file:

**For Client Credentials Flow (Default - Multi-User):**

```bash
cat > .env << 'EOF'
# Docker Project Name (used for container/network/volume naming)
COMPOSE_PROJECT_NAME=ms365-mcp

# Organization Mode (enables shared mailboxes, Teams, SharePoint)
MS365_MCP_ORG_MODE=true

# Azure AD Configuration
MS365_MCP_CLIENT_ID=your-client-id-here
MS365_MCP_TENANT_ID=your-tenant-id-here
MS365_MCP_CLIENT_SECRET=your-client-secret-here

# Optional Settings
MS365_MCP_LOG_LEVEL=info
MS365_MCP_ENABLE_MAIL=true
MS365_MCP_ENABLE_CALENDAR=true
MS365_MCP_ENABLE_FILES=true
MS365_MCP_ENABLE_TEAMS=true
MS365_MCP_ENABLE_EXCEL_POWERPOINT=true
MS365_MCP_ENABLE_ONENOTE=true
MS365_MCP_ENABLE_TASKS=true
EOF

chmod 600 .env
```

Replace:
- `your-client-id-here` - Application (client) ID from Azure AD
- `your-tenant-id-here` - Directory (tenant) ID from Azure AD
- `your-client-secret-here` - Client secret value from Azure AD

**For Device Code Flow (Optional - Single-User):**

```bash
cat > .env << 'EOF'
# Docker Project Name
COMPOSE_PROJECT_NAME=ms365-mcp

# Organization Mode
MS365_MCP_ORG_MODE=true

# Azure AD Configuration (NO CLIENT_SECRET for device code flow)
MS365_MCP_CLIENT_ID=your-client-id-here
MS365_MCP_TENANT_ID=your-tenant-id-here

# Optional Settings
MS365_MCP_LOG_LEVEL=info
EOF

chmod 600 .env
```

### 3. Build Docker Image

```bash
docker compose build
docker images | grep ms365-mcp  # Verify
```

### 4. Initial Authentication

**For Client Credentials Flow (Default):**

No authentication step required! The service will automatically acquire tokens using the client secret.

```bash
# Skip to step 5 - just start the service
docker compose up -d
```

**For Device Code Flow (Optional):**

```bash
# Start container interactively for device code authentication
docker compose run --rm ms365-mcp

# Follow device code instructions:
# 1. Open https://microsoft.com/devicelogin on any device
# 2. Enter the displayed code
# 3. Sign in with your M365 Business account
# 4. Grant permissions

# Wait for "Device code login successful"
# Press Ctrl+C to exit
```

### 5. Start Service

```bash
docker compose up -d
docker compose ps  # Verify running
```

## Token Management

**How it works:**
- Tokens stored in Docker volume at `/app/data/`
- Access tokens refresh automatically when expired
- No user interaction required after initial setup
- Refresh tokens valid for ~90 days (with active use)

**Token files:**
```bash
docker compose exec ms365-mcp ls -lh /app/data/
# .token-cache.json - Access and refresh tokens
# .selected-account.json - Current account
```


## Management

### View Logs

```bash
docker compose logs -f                    # Real-time
docker compose logs --tail 100           # Last 100 lines
docker compose logs --since 1h           # Last hour
docker compose logs | grep -i token      # Search tokens
```

### Container Operations

```bash
docker compose up -d      # Start
docker compose stop       # Stop
docker compose restart    # Restart
docker compose ps         # Status
docker compose down       # Remove (keeps volume)
docker compose down -v    # Remove with volume (DELETES TOKENS)
```

### Updates

```bash
git pull
docker compose down
docker compose build
docker compose up -d
# Tokens preserved in volume
```

## Security

### Backup Token Volume

```bash
# Backup
docker run --rm -v ms365-mcp-token-cache:/data \
  -v $(pwd):/backup alpine \
  tar czf /backup/token-backup-$(date +%F).tar.gz -C /data .

# Encrypt
gpg -c token-backup-*.tar.gz
```

### Network Verification

```bash
# Verify no exposed ports
docker compose ps

# Check network isolation
docker inspect ms365-mcp-server | grep -A 20 NetworkSettings
```

## Troubleshooting

### Expired Tokens

If refresh token expires (90+ days inactive):

```bash
docker compose down
docker volume rm ms365-mcp-token-cache
docker compose run --rm ms365-mcp  # Re-authenticate
docker compose up -d
```

### Check Status

```bash
docker compose ps                                # Running?
docker compose logs --tail 50                   # Recent logs
docker compose logs | grep -i "auth\|error"    # Errors
```

### Permission Issues

```bash
docker compose exec ms365-mcp ls -la /app/data/
docker compose exec -u root ms365-mcp chown -R node:node /app/data
```

## Container Communication

For connecting other containers to the MCP server, see the detailed sections in the original document covering:
- Shared Docker networks
- STDIO mode with `docker exec`
- HTTP mode on private networks
- Security considerations

**Quick setup:** Both containers join the same Docker network, then either:
- Use `docker exec -i ms365-mcp-server node /app/dist/index.js` (STDIO)
- Use `http://ms365-mcp-server:3000/mcp` (HTTP mode)

## Maintenance Schedule

**Monthly:**
```bash
docker compose ps
docker compose logs --since 7d | grep -i refresh
```

**Every 60 Days:**
- Ensure at least one API call to keep refresh token active
- Service does this automatically if in use

**After Reboots:**
- Container auto-starts (`restart: unless-stopped` policy)
- Verify with `docker compose ps`

## Production Checklist

- [ ] Azure AD app registered with correct permissions (Application or Delegated based on auth mode)
- [ ] Client secret created (for client credentials flow)
- [ ] Admin consent granted for all permissions
- [ ] Environment variables configured in `.env`
- [ ] `.env` file permissions set to 600
- [ ] Initial authentication completed (device code if using that flow)
- [ ] Container running in detached mode
- [ ] Logs show successful startup and token acquisition
- [ ] Token cache backup configured
- [ ] Monitoring alerts configured

## Network Requirements

**Outbound (Required):**
- `login.microsoftonline.com:443` - Authentication & token refresh
- `graph.microsoft.com:443` - API calls

**Inbound (None):**
- No inbound access required
- Server makes outbound requests only

## Summary

Your MS-365 MCP Server now:
- ✅ Runs without internet exposure
- ✅ Automatically refreshes access tokens
- ✅ Requires authentication only once (until refresh expires)
- ✅ Operates in STDIO mode for maximum security
- ✅ Auto-starts via Docker Compose
- ✅ Persists tokens in Docker volume
- ✅ Works with M365 Business via custom Azure AD app

**Critical success factors:**
1. Outbound internet to Microsoft endpoints
2. Docker volume persistence (named volumes)
3. Active refresh token (regular API calls)
4. Azure AD app permissions maintained

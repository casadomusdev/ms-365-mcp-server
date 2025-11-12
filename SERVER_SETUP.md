# Server Setup Guide - Docker Deployment

This guide explains how to deploy the Microsoft 365 MCP Server in Docker with automatic token refresh, without requiring internet exposure.

## Overview

The server runs in **Docker** using **STDIO mode** which:
- Requires zero internet exposure (no public IP, domain, or open ports)
- Automatically refreshes access tokens using cached refresh tokens
- Only needs outbound internet access to Microsoft's endpoints
- Uses device code authentication (one-time setup)
- Persists tokens in a Docker volume
- **Runs in Organization Mode** for full M365 Business feature support

## Organization Mode (Org-Mode)

**CRITICAL for Microsoft 365 Business**: This setup uses **Organization Mode (--org-mode)** which is **REQUIRED** to access:

### Features That Require Org-Mode:
- ✅ **Shared Mailboxes** - Read and send from shared mailboxes
- ✅ **Microsoft Teams** - Read/send messages in teams and channels
- ✅ **SharePoint** - Access SharePoint sites, lists, and documents
- ✅ **List All Users** - Query organization directory
- ✅ **Meeting Scheduling** - Find meeting times across calendars
- ✅ **Organization-wide Search** - Search across all organization content

### What Org-Mode Does:
- Enables all organization-level scopes (*.All, *.Shared, Teams, SharePoint)
- Registers these scopes during initial authentication
- Ensures all advanced M365 Business features are available

**Without org-mode**, you would only have access to personal mailbox, calendar, and files. For Microsoft 365 Business deployments, **org-mode is mandatory**.

## Prerequisites

### Server Requirements
- Docker and Docker Compose installed
- Outbound internet access to:
  - `login.microsoftonline.com` (for authentication & token refresh)
  - `graph.microsoft.com` (for API calls)
- SSH access to the server

### Microsoft 365 Business - Azure App Registration

**IMPORTANT**: For Microsoft 365 Business accounts, you **MUST register a custom Azure AD app**. The default public client is not available for organization accounts.

#### Step-by-Step Azure AD App Setup

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** → **App registrations**
3. Click **New registration**
4. Configure the application:
   - **Name**: `MS365 MCP Server`
   - **Supported account types**: 
     - Select **"Accounts in this organizational directory only" (Single tenant)**
     - **IMPORTANT**: Use single tenant, NOT multitenant
     - This avoids publisher verification requirements (only applies to multitenant apps)
   - **Redirect URI**: Leave blank (not needed for device code flow)
5. Click **Register**

   **Note on Publisher Verification**: Azure requires publisher verification for newly registered multitenant apps. However, since we're creating a **single-tenant app** for your organization only, this requirement does not apply. Your organization's admin can grant consent to apps registered in their own tenant without verification.

6. **Note the following values** (you'll need them):
   - **Application (client) ID**
   - **Directory (tenant) ID**

7. **Configure API Permissions**:
   - Click **API permissions** in the left menu
   - Click **Add a permission**
   - Select **Microsoft Graph** → **Delegated permissions**
   
   **Add ALL of the following permissions** (organized by category):
   
   **User & Authentication:**
   - `User.Read` - Read user profile
   - `User.Read.All` - Read all users' profiles
   
   **Mail (Personal Mailbox):**
   - `Mail.Read` - Read mail
   - `Mail.ReadWrite` - Read and write mail
   - `Mail.Send` - Send mail
   
   **Mail (Shared Mailboxes):**
   - `Mail.Read.Shared` - Read shared mailboxes
   - `Mail.ReadWrite.Shared` - Read and write shared mailboxes
   - `Mail.Send.Shared` - Send from shared mailboxes
   
   **Calendar:**
   - `Calendars.Read` - Read calendars
   - `Calendars.ReadWrite` - Read and write calendars
   - `Calendars.Read.Shared` - Read shared calendars (for meeting scheduling)
   - `Calendars.ReadWrite.Shared` - Read and write shared calendars (for meeting scheduling)
   
   **Files & OneDrive:**
   - `Files.Read` - Read files
   - `Files.ReadWrite` - Read and write files
   - `Files.Read.All` - Read all files (for search)
   
   **Tasks & Planning:**
   - `Tasks.Read` - Read tasks (To Do & Planner)
   - `Tasks.ReadWrite` - Read and write tasks
   
   **Contacts:**
   - `Contacts.Read` - Read contacts
   - `Contacts.ReadWrite` - Read and write contacts
   
   **OneNote:**
   - `Notes.Read` - Read OneNote notebooks
   - `Notes.Create` - Create OneNote pages
   
   **Teams & Chat:**
   - `Chat.Read` - Read chats
   - `ChatMessage.Read` - Read chat messages
   - `ChatMessage.Send` - Send chat messages
   - `Team.ReadBasic.All` - Read basic team information
   - `Channel.ReadBasic.All` - Read basic channel information
   - `ChannelMessage.Read.All` - Read channel messages
   - `ChannelMessage.Send` - Send channel messages
   - `TeamMember.Read.All` - Read team members
   
   **SharePoint:**
   - `Sites.Read.All` - Read all SharePoint sites
   
   **Search:**
   - `People.Read` - Read people (for search)
   
   - Click **Add permissions**
   - **IMPORTANT - Admin Consent**: Click "Grant admin consent for [Your Organization]"
     - This is REQUIRED for all organization-level permissions (*.All, Shared, Teams, etc.)

8. **Enable Device Code Flow**:
   - Click **Authentication** in the left menu
   - Under **Advanced settings**
   - Set **Allow public client flows** to **Yes**
   - Click **Save**

**You do NOT need a client secret for device code flow.**

## Installation Steps

### 1. Deploy to Server

```bash
# SSH to your server
ssh user@your-server

# Create project directory
sudo mkdir -p /opt/ms-365-mcp-server
cd /opt/ms-365-mcp-server

# Clone/copy the project files
git clone <your-repo> .

# Or use scp to copy from local:
# scp -r /path/to/ms-365-mcp-server/* user@server:/opt/ms-365-mcp-server/
```

### 2. Configure Environment

Create a `.env` file with your Azure AD app details:

```bash
cat > .env << 'EOF'
# Docker Compose Project Name (used for container/network/volume naming)
COMPOSE_PROJECT_NAME=my-project-ms365-mcp

# Organization Mode (REQUIRED for M365 Business - enables shared mailboxes, Teams, SharePoint)
MS365_MCP_ORG_MODE=true

# Optional: Log level (debug, info, warn, error)
LOG_LEVEL=info

# Azure AD App Configuration (REQUIRED for M365 Business)
MS365_MCP_CLIENT_ID=your-application-client-id-here
MS365_MCP_TENANT_ID=your-directory-tenant-id-here

# Feature toggles for tool groups (optional - all enabled by default in org-mode)
MS365_MCP_ENABLE_MAIL=true
MS365_MCP_ENABLE_CALENDAR=true
MS365_MCP_ENABLE_FILES=true
MS365_MCP_ENABLE_TEAMS=true
MS365_MCP_ENABLE_EXCEL_POWERPOINT=true
MS365_MCP_ENABLE_ONENOTE=true
MS365_MCP_ENABLE_TASKS=true
EOF

# Secure the .env file
chmod 600 .env
```

**Replace** `your-application-client-id-here` and `your-directory-tenant-id-here` with the values from your Azure AD app registration.

**Configuration Notes**:
- `COMPOSE_PROJECT_NAME`: Names all Docker resources (container, network, volume). Change this to run multiple instances.
- `MS365_MCP_ORG_MODE`: Enabled by default - REQUIRED for shared mailboxes, Teams, SharePoint.
- `LOG_LEVEL`: Controls logging verbosity. Use `debug` for troubleshooting.
- Feature toggles: Enable/disable specific M365 service groups as needed.

**Docker Resources Created**:
- Container: `ms365-mcp-server` (or `${COMPOSE_PROJECT_NAME}-server` if customized)
- Network: `ms365-mcp-net` (isolated bridge network)
- Volume: `ms365-mcp-token-cache` (persistent token storage)
- Logging: `local` driver with automatic log rotation

### 3. Build the Docker Image

```bash
# Build the image
docker compose build

# Verify the image was created
docker images | grep ms365-mcp
```

### 4. Initial Authentication

This is a one-time process to authenticate and cache tokens:

```bash
# Start the container interactively for authentication
docker compose run --rm ms365-mcp

# You'll see output like:
# To sign in, use a web browser to open the page https://microsoft.com/devicelogin
# and enter the code AB12-CD34 to authenticate.
```

**Complete Authentication**:
1. On ANY device (your laptop, phone, etc.), visit the URL shown
2. Enter the code displayed
3. Sign in with your Microsoft 365 Business account
4. Review and grant the requested permissions
5. Return to the server terminal

**Success Indicators**:
- You'll see "Device code login successful"
- The Docker volume `ms365-mcp-token-cache` is created with token data
- The server is now authenticated

Press `Ctrl+C` to exit the interactive session.

### 5. Start the Service

```bash
# Start the container in detached mode
docker compose up -d

# Verify it's running
docker compose ps
```

The container will now run continuously and automatically restart if it stops.

### 6. Verify Token Cache

```bash
# Check that tokens are cached in the volume
docker compose exec ms365-mcp ls -lh /app/data/

# View volume details
docker volume inspect ms365-mcp-token-cache
```

## How Automatic Token Refresh Works

Once authenticated:

1. **Token Storage**: The Docker volume `/app/data/` contains:
   - `.token-cache.json` - Access and refresh tokens
   - `.selected-account.json` - Currently selected account
   - Both files are automatically managed by MSAL

2. **Automatic Refresh**: When the server needs to make an API call:
   - Checks if cached access token is still valid
   - If expired, automatically uses refresh token to get new access token
   - Makes outbound HTTPS request to `login.microsoftonline.com/token`
   - Updates token cache in the volume
   - Continues with API call

3. **No User Interaction**: All refresh happens automatically in the background

4. **No Internet Exposure**: The server makes outbound requests only - no inbound connections needed

## Docker Management

### Viewing Logs

```bash
# View logs in real-time
docker compose logs -f

# View logs for specific time period
docker compose logs --since 1h
docker compose logs --since "2024-01-15"

# View last 100 lines
docker compose logs --tail 100

# Search logs for token refresh activity
docker compose logs | grep -i "token\|refresh\|auth"
```

### Managing the Container

```bash
# Start the service
docker compose up -d

# Stop the service
docker compose stop

# Restart the service
docker compose restart

# View container status
docker compose ps

# Execute command in running container
docker compose exec ms365-mcp sh

# Remove container (keeps volume)
docker compose down

# Remove container AND volume (DELETES TOKENS!)
docker compose down -v
```

### Updating the Application

```bash
# Pull latest code
git pull

# Rebuild and restart
docker compose down
docker compose build
docker compose up -d

# Tokens are preserved in the volume
```

## Security Best Practices

### Volume Security

The token cache volume contains sensitive authentication data. Ensure proper protection:

```bash
# Backup the volume (adjust volume name if COMPOSE_PROJECT_NAME is different)
docker run --rm -v ms365-mcp-token-cache:/data \
  -v $(pwd):/backup alpine \
  tar czf /backup/token-cache-backup-$(date +%F).tar.gz -C /data .

# Encrypt the backup
gpg -c token-cache-backup-*.tar.gz

# Store encrypted backup securely offsite
```

### Monitoring

```bash
# Monitor container health
docker compose ps
docker stats ms365-mcp-server

# Check for errors
docker compose logs | grep -i error

# Monitor token refresh activity
docker compose logs --since 1h | grep -i "token\|refresh"
```

### Network Security

```bash
# Verify no ports are exposed
docker compose ps
# Should show no port mappings

# Inspect network configuration
docker inspect ms365-mcp-server | grep -A 20 NetworkSettings
```

## Troubleshooting

### Tokens Expired After Long Inactivity

If refresh token expires (typically 90 days of no use):

```bash
# Stop and remove container
docker compose down

# Remove the token volume (adjust volume name if COMPOSE_PROJECT_NAME is different)
docker volume rm ms365-mcp-token-cache

# Re-run authentication
docker compose run --rm ms365-mcp
# Complete device code flow again

# Start the service
docker compose up -d
```

### Checking Container Status

```bash
# Is the container running?
docker compose ps

# View recent logs
docker compose logs --tail 50

# Check for authentication errors
docker compose logs | grep -i "auth\|token\|error"
```

### Permission Errors

```bash
# Check volume permissions
docker compose exec ms365-mcp ls -la /app/data/

# If needed, fix permissions (run as root)
docker compose exec -u root ms365-mcp chown -R node:node /app/data
```

### Testing Token Refresh

```bash
# Watch logs for token refresh
docker compose logs -f | grep -i refresh

# Force a token refresh by waiting for access token to expire (1 hour)
# Or check if automatic refresh is working
```

### Accessing Token Cache for Backup

```bash
# Create a temporary container to access volume (adjust volume name if COMPOSE_PROJECT_NAME is different)
docker run --rm -v ms365-mcp-token-cache:/data alpine ls -lh /data

# Copy token cache to host
docker run --rm -v ms365-mcp-token-cache:/data \
  -v $(pwd):/backup alpine \
  cp /data/.token-cache.json /backup/

# View token cache (careful - contains sensitive data!)
cat .token-cache.json | jq .
```

## Production Deployment Checklist

- [ ] Azure AD app registered with correct permissions
- [ ] Admin consent granted for required API permissions
- [ ] Environment variables configured in `.env`
- [ ] `.env` file has restricted permissions (600)
- [ ] Initial device code authentication completed
- [ ] Token cache volume created and populated
- [ ] Container running in detached mode
- [ ] Logs show successful startup and token acquisition
- [ ] Token cache backup process configured
- [ ] Monitoring alerts set up for container health
- [ ] Documentation updated with your specific Azure AD app details

## Network Requirements

### Outbound Access Required
- `login.microsoftonline.com` (port 443) - Authentication & token refresh
- `graph.microsoft.com` (port 443) - API calls

### Inbound Access Required
- **None** - the server makes outbound requests only

### Docker Network
- The container uses bridge network mode (default)
- No port mappings required for STDIO mode
- No need for host network access

## Maintenance

### Regular Tasks

**Monthly**: Verify the container is running and tokens are refreshing
```bash
docker compose ps
docker compose logs --since 7d | grep -i refresh
```

**Every 60 Days**: Make at least one API call to prevent refresh token expiry
- The service should do this automatically if actively used
- Monitor token expiry in logs

**After Server Reboots**: 
- Container auto-starts with `restart: unless-stopped` policy
- Verify with `docker compose ps`

**Before Production Use**: Test the complete authentication and refresh cycle

## Container-to-Container Communication

If you have an MCP client application (like a custom AI assistant or automation tool) running in a separate Docker container on the same server, you can connect them via a shared Docker network.

### Scenario: MCP Client Container → MS365 MCP Server Container

**Example Use Case**: An AI assistant container needs to access Microsoft 365 tools via the MCP server.

### Setup Steps

#### 1. Create a Shared Docker Network

```bash
# Create an external network that both containers can join
docker network create mcp-shared-network
```

#### 2. Configure MS365 MCP Server to Use Shared Network

Update `docker-compose.yaml` to connect to the external network:

```yaml
services:
  ms365-mcp:
    container_name: ${COMPOSE_PROJECT_NAME:-ms365-mcp}-server
    networks:
      - ms365-mcp-net      # Internal network (existing)
      - mcp-shared-network # External shared network (new)
    # ... rest of configuration ...

volumes:
  token-cache:
    driver: local
    name: ${COMPOSE_PROJECT_NAME:-ms365-mcp}-token-cache

networks:
  ms365-mcp-net:
    driver: bridge
    name: ${COMPOSE_PROJECT_NAME:-ms365-mcp}-net
  
  # Add external network reference
  mcp-shared-network:
    external: true
    name: mcp-shared-network
```

Apply the changes:
```bash
docker compose down
docker compose up -d
```

#### 3. Configure MCP Client Container

Your MCP client container needs to:
- Join the same `mcp-shared-network`
- Use `docker exec` to communicate with the MS365 MCP server via STDIO

**Example client `docker-compose.yaml`:**

```yaml
services:
  mcp-client:
    container_name: my-mcp-client
    image: my-client-image:latest
    networks:
      - mcp-shared-network
    environment:
      # Configure how to reach the MS365 MCP server
      MS365_MCP_CONTAINER: ms365-mcp-server
    # ... rest of configuration ...

networks:
  mcp-shared-network:
    external: true
    name: mcp-shared-network
```

#### 4. Client Communication Pattern

From within the client container, communicate with the MCP server using `docker exec`:

```bash
# Example: Send MCP protocol message from client container to MS365 MCP server
echo '{"jsonrpc":"2.0","id":1,"method":"initialize",...}' | \
  docker exec -i ms365-mcp-server node /app/dist/index.js
```

**In application code (e.g., Python):**

```python
import subprocess
import json

def call_mcp_tool(method, params):
    """Call MS365 MCP server tool from client container."""
    mcp_request = {
        "jsonrpc": "2.0",
        "id": 1,
        "method": method,
        "params": params
    }
    
    # Execute docker exec to communicate with MCP server
    process = subprocess.Popen(
        ["docker", "exec", "-i", "ms365-mcp-server", 
         "node", "/app/dist/index.js"],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE
    )
    
    stdout, stderr = process.communicate(
        input=json.dumps(mcp_request).encode()
    )
    
    return json.loads(stdout.decode())

# Example usage
response = call_mcp_tool("tools/list", {})
print(response)
```

#### 5. Security Considerations

**Docker Socket Access**: The client container needs access to the Docker socket to run `docker exec`:

```yaml
services:
  mcp-client:
    # ... other config ...
    volumes:
      - /var/run/docker.sock:/var/run/docker.sock
    user: root  # Required for Docker socket access
```

⚠️ **Security Warning**: Mounting the Docker socket gives the container full control over Docker. Only do this:
- In trusted environments
- With containers you control
- When necessary for the architecture

**Alternative: Use a Sidecar Pattern**

For better security, you could create a lightweight sidecar container that handles MCP communication:

```yaml
services:
  ms365-mcp:
    container_name: ms365-mcp-server
    # ... existing config ...

  mcp-proxy:
    container_name: mcp-proxy
    image: alpine:latest
    command: sh -c "while true; do sleep 3600; done"
    networks:
      - mcp-shared-network
    volumes:
      - /var/run/docker.sock:/var/run/docker.sock
    # This sidecar has Docker socket access
    # Your client calls this proxy instead

  mcp-client:
    container_name: my-mcp-client
    networks:
      - mcp-shared-network
    # Client has NO Docker socket access
    # Sends requests to mcp-proxy over network
```

### Network Architecture Diagram

```
┌─────────────────────────────────────────────┐
│  Server (Docker Host)                       │
│                                             │
│  ┌──────────────────────────────────────┐  │
│  │  mcp-shared-network (bridge)         │  │
│  │                                       │  │
│  │  ┌──────────────────────────────┐    │  │
│  │  │ ms365-mcp-server             │    │  │
│  │  │ - Authenticates with MS365   │    │  │
│  │  │ - Stores tokens in volume    │    │  │
│  │  │ - Exposes MCP via STDIO      │    │  │
│  │  └──────────────────────────────┘    │  │
│  │           ▲                           │  │
│  │           │ docker exec -i            │  │
│  │           │ (STDIO communication)     │  │
│  │           │                           │  │
│  │  ┌────────┴──────────────────────┐   │  │
│  │  │ mcp-client                    │   │  │
│  │  │ - Your AI assistant/app       │   │  │
│  │  │ - Calls MCP tools via exec    │   │  │
│  │  │ - Processes responses         │   │  │
│  │  └───────────────────────────────┘   │  │
│  └──────────────────────────────────────┘  │
│                                             │
│  Outbound Internet Access:                 │
│  → login.microsoftonline.com               │
│  → graph.microsoft.com                     │
└─────────────────────────────────────────────┘
```

### Testing Container-to-Container Communication

```bash
# 1. Verify both containers are on the shared network
docker network inspect mcp-shared-network

# Should show both containers in the "Containers" section

# 2. Test from client container
docker exec -it my-mcp-client sh

# Inside client container:
echo '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2024-11-05","capabilities":{},"clientInfo":{"name":"test","version":"1.0"}}}' | \
  docker exec -i ms365-mcp-server node /app/dist/index.js

# Should receive MCP protocol response
```

### Benefits of Container-to-Container Setup

- ✅ **Complete Isolation**: Both services isolated in containers
- ✅ **No Port Exposure**: Pure STDIO communication, no network ports
- ✅ **Shared Network**: Containers can discover each other by name
- ✅ **Scalable**: Easy to add more MCP servers or clients
- ✅ **Maintainable**: Each service has its own lifecycle

### Limitations

- ⚠️ **Docker Socket Required**: Client needs Docker socket access for `docker exec`
- ⚠️ **Same Host Only**: Both containers must be on the same Docker host
- ⚠️ **Not for Kubernetes**: This pattern is Docker Compose specific

For distributed deployments across multiple hosts, consider using HTTP transport instead of STDIO.

## Summary

You now have a Microsoft 365 MCP Server running in Docker that:
- ✅ Runs securely without internet exposure
- ✅ Automatically refreshes access tokens
- ✅ Requires authentication only once (until refresh token expires)
- ✅ Operates in STDIO mode for maximum security
- ✅ Starts automatically via Docker Compose restart policy
- ✅ Persists tokens in a Docker volume
- ✅ Works with Microsoft 365 Business accounts via custom Azure AD app
- ✅ Can communicate with other containers via shared Docker networks

The server will continue to function as long as:
1. The container has outbound internet access to Microsoft endpoints
2. The Docker volume persists (use named volumes, not anonymous)
3. The refresh token hasn't expired (keep it active with regular API calls)
4. Your Azure AD app permissions remain granted

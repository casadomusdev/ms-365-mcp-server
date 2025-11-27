# PowerShell Integration for Shared Mailbox Discovery

This document explains how PowerShell integration works for discovering shared mailboxes and their delegation permissions in the MS-365 MCP Server.

## Overview

The MCP server uses **Exchange Online PowerShell** to query mailbox delegation permissions (Full Access and SendAs) for shared mailboxes. This information is **not available** through Microsoft Graph API, making PowerShell the only programmatic method for accurate shared mailbox discovery.

### Key Features

✅ **Enabled by Default** - Works out of the box when PowerShell is available  
✅ **Auto-Detection** - Automatically detects if `pwsh` is installed on the system  
✅ **Graceful Fallback** - Falls back to personal mailbox only if PowerShell unavailable (with warning)  
✅ **No Additional Auth** - Reuses existing Microsoft Graph access token  
✅ **Cached Results** - 1-hour TTL cache minimizes performance impact

## How It Works

### Detection Strategy

When a user's mailboxes are discovered, the server:

1. **Personal Mailbox** - Always detected via Microsoft Graph API (fast, ~100ms)
2. **Shared Mailboxes** - Detected via Exchange Online PowerShell when available (slower, ~2-5s, but cached)

### Auto-Detection Behavior

On startup, the PowerShell service:
1. Checks if `MS365_POWERSHELL_ENABLED` environment variable is set
   - If not set → **defaults to enabled**
   - If set to `false` or `0` → explicitly disabled
2. Tests if `pwsh` command is available on the system
3. Logs the result:
   - ✅ **Both enabled + available** → Shared mailbox discovery will work
   - ⚠️ **Enabled but not available** → Warning logged, falls back to personal mailbox only
   - ℹ️ **Explicitly disabled** → Info logged, no PowerShell attempted

### Authentication Flow

#### Certificate-Based Authentication (App-Only Mode)

For client credentials (app-only) mode, PowerShell requires **certificate-based authentication**:

```
┌─────────────────┐
│  MCP Server     │
│ (app-only mode) │
└────────┬────────┘
         │
         │ 1. Read cert config from env
         ▼
┌─────────────────────┐
│ PowerShellService   │
│  (spawn pwsh)       │
└────────┬────────────┘
         │
         │ 2. Connect-ExchangeOnline -CertificateFilePath
         │    -CertificatePassword -AppId -Organization
         ▼
┌─────────────────────────────┐
│ Exchange Online PowerShell  │
│   - Get-Mailbox             │
│   - Get-MailboxPermission   │
│   - Get-RecipientPermission │
└─────────────────────────────┘
```

**Why Certificate Auth?**

Microsoft's `-AccessToken` parameter for `Connect-ExchangeOnline` **only works with delegated permissions** (user authentication). For app-only (client credentials) mode, you must use certificate-based authentication.

See: [Connect to Exchange Online PowerShell](https://learn.microsoft.com/en-us/powershell/exchange/connect-to-exchange-online-powershell)

#### Setting Up Certificate Authentication

The project includes an automated certificate generation script:

```bash
cd projects/ms-365-mcp-server

# Generate certificate (interactive - asks for validity period)
./auth-generate-cert.sh
```

**What the script does:**
1. Prompts for certificate validity (1-3 years, default 2)
2. Generates a self-signed certificate using PowerShell
3. Exports private key (.pfx) and public key (.cer)
4. Generates a secure random password
5. Updates `.env` with `MS365_CERT_PATH` and `MS365_CERT_PASSWORD`
6. Provides instructions for uploading to Azure AD

**Upload Certificate to Azure AD:**

After generating the certificate, you must upload the public key to Azure AD:

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to: Azure Active Directory → App registrations
3. Select your app (using Client ID from `.env`)
4. Go to: Certificates & secrets → Certificates tab
5. Click: [Upload certificate]
6. Select file: `certs/ms365-powershell.cer`
7. Add description: "PowerShell Exchange Online (expires YYYY-MM-DD)"
8. Click: [Add]

**Certificate is automatically integrated into auth-login.sh:**

When you run `./auth-login.sh`:
- If certificate doesn't exist → prompts to generate it
- If certificate exists → confirms it's ready
- Provides upload instructions if needed

**Certificate Files (gitignored):**
- `certs/ms365-powershell.pfx` - Private key (NEVER commit this!)
- `certs/ms365-powershell.cer` - Public key (upload to Azure AD)
- `.env` contains `MS365_CERT_PASSWORD` (NEVER commit this!)

## Installation

### Prerequisites

You need two components installed:

1. **PowerShell Core 7.x** (not Windows PowerShell 5.1)
2. **Exchange Online PowerShell Module v3.x**

### On macOS (Homebrew)

```bash
# Install PowerShell Core
brew install --cask powershell

# Verify installation
pwsh -Version

# Install Exchange Online module
pwsh -Command "Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber"

# Verify module
pwsh -Command "Get-Module -ListAvailable ExchangeOnlineManagement"
```

### On Linux (Debian/Ubuntu)

```bash
# Install PowerShell Core
# Download the Microsoft repository GPG keys
wget -q https://packages.microsoft.com/config/ubuntu/$(lsb_release -rs)/packages-microsoft-prod.deb

# Register the Microsoft repository GPG keys
sudo dpkg -i packages-microsoft-prod.deb

# Update the list of packages
sudo apt-get update

# Install PowerShell
sudo apt-get install -y powershell

# Verify installation
pwsh -Version

# Install Exchange Online module
pwsh -Command "Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber"
```

### In Docker (Debian-based)

The provided `Dockerfile` already includes PowerShell Core and Exchange Online module installation:

```dockerfile
# Install PowerShell Core 7.x
RUN apt-get update && \
    apt-get install -y wget apt-transport-https software-properties-common && \
    wget -q https://packages.microsoft.com/config/debian/12/packages-microsoft-prod.deb && \
    dpkg -i packages-microsoft-prod.deb && \
    apt-get update && \
    apt-get install -y powershell && \
    rm packages-microsoft-prod.deb

# Install Exchange Online PowerShell module
RUN pwsh -Command "Install-Module -Name ExchangeOnlineManagement -Repository PSGallery -Force -AllowClobber"
```

## Configuration

### Environment Variables

```bash
# Enable/disable PowerShell integration (default: enabled)
# Set to 'false' or '0' to explicitly disable
MS365_POWERSHELL_ENABLED=true

# PowerShell command timeout in milliseconds (default: 30000 = 30 seconds)
MS365_POWERSHELL_TIMEOUT=30000
```

### Entra ID Permissions (Required for PowerShell)

**IMPORTANT:** For PowerShell integration to work, you need BOTH API permissions AND an Exchange administrator role.

#### Step 1: Add API Permissions

1. Go to [Entra Admin Center](https://entra.microsoft.com)
2. Navigate to: **Applications** → **App registrations**
3. Select your application
4. Go to: **API permissions**
5. Click: **+ Add a permission**
6. Select: **APIs my organization uses** → Search for **Office 365 Exchange Online**
7. Choose: **Application permissions**
8. Select: **Exchange.ManageAsApp**
9. Click: **Add permissions**
10. Click: **Grant admin consent for [Your Organization]** ⚠️ CRITICAL!
11. Confirm the consent

#### Step 2: Assign Exchange Administrator Role

**This step is IN ADDITION to API permissions - both are required!**

1. In [Entra Admin Center](https://entra.microsoft.com)
2. Navigate to: **Roles & admins**
3. Search for: **Exchange Administrator**
4. Click on the role
5. Click: **+ Add assignments**
6. Search for your application by name or Client ID
7. Select your application
8. Click: **Add**

**Why both are needed:**
- API permission (`Exchange.ManageAsApp`) grants access to Exchange Online API
- Exchange Administrator role grants permission to query mailbox delegation
- Without both, you'll get: "The role assigned to application isn't supported"

#### Additional Graph API Permissions

These are used for mailbox discovery via Graph API (personal mailboxes):

**Client Credentials Flow (Application Permissions):**
- `Mail.Read` or `Mail.ReadWrite` (for full mailbox access)
- `User.Read.All` (for user validation)

**Device Code Flow (Delegated Permissions):**
- `Mail.Read` or `Mail.ReadWrite`
- `User.Read`

## Usage

Once installed and configured, PowerShell integration works automatically:

```typescript
// Mailbox discovery happens transparently
const mailboxes = await mailboxDiscoveryService.discoverMailboxes('user@company.com');

// Results include:
// 1. Personal mailbox (via Graph API)
// 2. Shared mailboxes with Full Access or SendAs (via PowerShell)
```

### Log Output Examples

**✅ Success (PowerShell available):**
```
[INFO] PowerShell integration enabled and available
[INFO] PowerShell timeout: 30000ms
[INFO] Starting discovery for user@company.com
[INFO] ✓ Found personal mailbox: John Doe
[INFO] Querying shared mailboxes via Exchange Online PowerShell...
[INFO] ✓ Found 2 shared mailbox(es) via PowerShell
[INFO] Discovery complete: 3 total (1 personal + 2 shared)
```

**⚠️ PowerShell Not Available:**
```
[WARN] PowerShell integration enabled but pwsh is not available on this system
[WARN] Shared mailbox discovery will be disabled - only personal mailboxes will be accessible
[INFO] To enable shared mailbox discovery, install PowerShell Core 7.x and Exchange Online PowerShell module
[INFO] Starting discovery for user@company.com
[INFO] ✓ Found personal mailbox: John Doe
[INFO] PowerShell integration disabled - skipping shared mailbox discovery
[INFO] Discovery complete: 1 total (1 personal + 0 shared)
```

**ℹ️ Explicitly Disabled:**
```
[INFO] PowerShell integration explicitly disabled via MS365_POWERSHELL_ENABLED=false
[INFO] Starting discovery for user@company.com
[INFO] ✓ Found personal mailbox: John Doe
[INFO] PowerShell integration disabled - skipping shared mailbox discovery
[INFO] Set MS365_POWERSHELL_ENABLED=true to enable shared mailbox detection
[INFO] Discovery complete: 1 total (1 personal + 0 shared)
```

## Performance

### Execution Times

| Operation | Method | Time | Cached |
|-----------|--------|------|--------|
| Personal mailbox | Graph API | ~100ms | No |
| Shared mailboxes | PowerShell | ~2-5s | Yes (1 hour) |

### Caching Strategy

- **Cache Key:** User email address
- **Cache TTL:** 1 hour (3600 seconds, configurable via `MS365_MCP_IMPERSONATE_CACHE_TTL`)
- **Cache Invalidation:** Automatic expiration only

The cache is critical for acceptable performance. Without caching, every mailbox discovery would take 2-5 seconds.

## Troubleshooting

### PowerShell Not Found

**Symptom:** Warning about `pwsh` not being available

**Solution:**
```bash
# Verify PowerShell is installed
which pwsh
pwsh -Version

# If not installed, install PowerShell Core 7.x (see Installation section above)
```

### Exchange Module Not Found

**Symptom:** PowerShell script fails with "Module not found"

**Solution:**
```bash
# Check if module is installed
pwsh -Command "Get-Module -ListAvailable ExchangeOnlineManagement"

# Install if missing
pwsh -Command "Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber"

# Verify installation
pwsh -Command "Get-InstalledModule ExchangeOnlineManagement"
```

### Access Token Issues

**Symptom:** "Failed to connect to Exchange Online" or "Invalid access token"

**Possible Causes:**
1. Insufficient Azure AD app permissions
2. Tenant ID misconfigured (using `common` instead of specific tenant)

**Solution:**
```bash
# Verify tenant ID is set correctly
echo $MS365_MCP_TENANT_ID  # Should be your actual tenant ID, not "common"

# Check Azure AD app permissions in Azure Portal
# Ensure Mail.Read/ReadWrite and User.Read.All are granted and admin-consented
```

### Timeout Issues

**Symptom:** "PowerShell script execution timed out"

**Solution:**
```bash
# Increase timeout (in milliseconds)
export MS365_POWERSHELL_TIMEOUT=60000  # 60 seconds

# Check network connectivity to Exchange Online
ping outlook.office365.com
```

### Permission Errors

**Symptom:** User has access to shared mailbox but PowerShell doesn't find it

**Possible Causes:**
1. Exchange permissions take time to propagate
2. User email format mismatch (display name vs UPN)

**Solution:**
```bash
# Wait 15-30 minutes for permission changes to propagate in Exchange Online

# Ensure you're using the UserPrincipalName (UPN), not display name
# Correct: user@company.com
# Wrong: "John Doe"

# Test manually with PowerShell to verify permissions:
pwsh -Command "Connect-ExchangeOnline -UserPrincipalName admin@company.com"
pwsh -Command "Get-MailboxPermission -Identity shared@company.com | Where-Object {$_.User -like '*user@company.com*'}"
```

## Disabling PowerShell Integration

To explicitly disable PowerShell integration:

```bash
# In .env file
MS365_POWERSHELL_ENABLED=false

# Or in docker-compose.yaml
environment:
  MS365_MCP_POWERSHELL_ENABLED: "false"

# Or as environment variable
export MS365_POWERSHELL_ENABLED=false
```

When disabled:
- No PowerShell detection or execution attempted
- Only personal mailboxes discovered via Graph API
- Faster startup (no pwsh availability check)
- Lower memory footprint

## Security Considerations

### Access Token Handling

- Access tokens are obtained from the existing AuthManager
- Tokens are passed to PowerShell via command-line arguments (secure on trusted systems)
- Tokens are short-lived and automatically refreshed by AuthManager
- No tokens are persisted to disk by PowerShell service

### Script Execution

- PowerShell is invoked with `-NoProfile` (skips user profile scripts)
- PowerShell is invoked with `-NonInteractive` (prevents interactive prompts)
- Script path is validated and constructed programmatically
- User input is only passed as typed parameters (no script injection possible)

### Recommendations

1. **Run in Docker** - Provides isolation and consistent environment
2. **Limit Script Access** - Ensure only authorized users can modify PowerShell scripts
3. **Monitor Logs** - Watch for unusual PowerShell execution patterns
4. **Use Specific Tenant ID** - Avoid `common` tenant in production

## Advanced Configuration

### Custom PowerShell Script Location

The PowerShell script is located at `scripts/check-mailbox-permissions.ps1`. To customize:

1. Modify the script (add logging, change query logic, etc.)
2. Rebuild the project: `npm run build`
3. Restart the server

### Connection Optimization

For environments with many users, consider:

```bash
# Increase cache TTL to reduce PowerShell executions
export MS365_MCP_IMPERSONATE_CACHE_TTL=7200  # 2 hours

# Implement cache warming (future enhancement)
# Pre-populate cache for active users during off-peak hours
```

## Comparison: Graph API vs PowerShell

| Feature | Graph API | PowerShell |
|---------|-----------|------------|
| Personal Mailbox | ✅ Works perfectly | ✅ Works but slower |
| Shared Mailbox Detection | ❌ No endpoint exists | ✅ Get-Mailbox |
| Full Access Permission | ❌ Not available | ✅ Get-MailboxPermission |
| SendAs Permission | ❌ Wrong permission type | ✅ Get-RecipientPermission |
| Performance | ⚡ Fast (~100ms) | 🐌 Slow (~2-5s) |
| Authentication | 🔑 Access token | 🔑 Same access token |
| Availability | 🌍 Always available | ⚙️ Requires pwsh installed |

## Manual Testing & Advanced Troubleshooting

### Quick Debug Script (Recommended)

The easiest way to debug PowerShell issues and extract tokens is using the included debug script:

```bash
cd projects/ms-365-mcp-server
./debug-pwsh.sh
```

**What it does:**
1. ✓ Checks PowerShell Core installation
2. ✓ Verifies Exchange Online module
3. ✓ Extracts access token automatically
4. ✓ Reads parameters from .env
5. ✓ Runs PowerShell script with verbose output
6. ✓ Shows detailed error messages and stack traces

**Output example:**
```
PowerShell Debugging Tool

[1/5] Checking PowerShell Core availability...
✓ PowerShell Core found: PowerShell 7.5.0

[2/5] Checking Exchange Online module...
✓ Exchange Online module found

[3/5] Extracting access token...
✓ Access token extracted successfully
Token length: 1842 characters
Token preview: eyJ0eXAiOiJKV1QiLCJub25jZSI...

[4/5] Gathering parameters...
Tenant ID: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
User email: user@company.com

[5/5] Testing PowerShell script...
Running: pwsh scripts/check-mailbox-permissions.ps1

--- PowerShell Script Execution ---
[Verbose output shows exactly what's happening...]
--- Script completed successfully ---

✓ PowerShell script executed successfully
```

**Common errors you'll see:**
- "No token available" → Run `./auth-login.sh` first
- "PowerShell Core not found" → Install with `brew install --cask powershell`
- "Module not found" → Install with `pwsh -Command "Install-Module -Name ExchangeOnlineManagement -Force"`
- "Access token invalid" → Token expired, re-run `./auth-login.sh`

### Extracting Access Token for Manual Testing

To test the PowerShell script directly, you need an access token. Here's how to extract it from the token cache:

#### Method 1: Using Node.js (Recommended)

Create a temporary script to extract the token:

```bash
cd projects/ms-365-mcp-server

# Create token extraction script
cat > extract-token.js << 'EOF'
import 'dotenv/config';
import AuthManager, { buildScopesFromEndpoints } from './dist/auth.js';

async function extractToken() {
  try {
    const scopes = buildScopesFromEndpoints(true);
    const authManager = new AuthManager(undefined, scopes);
    await authManager.loadTokenCache();
    
    const token = await authManager.getToken();
    
    if (token) {
      console.log('Access Token:');
      console.log(token);
      console.log('');
      console.log('Token Length:', token.length);
      console.log('Expires: Check token at https://jwt.ms to see expiration');
    } else {
      console.error('No token available. Run ./auth-login.sh first.');
    }
  } catch (error) {
    console.error('Error:', error.message);
  }
}

extractToken();
EOF

# Run the extraction
node extract-token.js

# Clean up
rm extract-token.js
```

#### Method 2: From Token Cache File (macOS/Linux)

Tokens are stored in different locations based on your OS:

**macOS (Keychain):**
```bash
# List stored tokens
security find-generic-password -s "msal.token.cache" -g 2>&1 | grep "password:"

# Or use the helper script
node scripts/keychain-helper.js list
```

**Linux/macOS (File-based):**
```bash
# Token cache is in the project directory
cd projects/ms-365-mcp-server

# View the entire cache (this is MSAL format, complex structure)
cat .token-cache.json | jq '.'

# Extract access token (MSAL cache structure - this may be empty if using keychain)
cat .token-cache.json | jq -r '.AccessToken | to_entries[] | .value.secret' 2>/dev/null

# Note: If you forced file cache with MS365_MCP_FORCE_FILE_CACHE=true, 
# the token will be here. Otherwise it may be in macOS Keychain.
```

**Docker (Inside container):**
```bash
# Enter the container
docker exec -it ms365-mcp-server sh

# View token cache
cat /app/data/.token-cache.json | jq -r '.AccessToken[] | select(.cached_at) | .secret'
```

### Running PowerShell Script Manually

Once you have the access token, test the PowerShell script directly:

```bash
cd projects/ms-365-mcp-server

# Set variables
USER_EMAIL="user@company.com"
ACCESS_TOKEN="eyJ0eXAiOiJKV1QiLCJub25jZSI..."  # Your extracted token
TENANT_ID="your-tenant-id"  # From .env: MS365_MCP_TENANT_ID

# Run the script
pwsh scripts/check-mailbox-permissions.ps1 \
  -UserEmail "$USER_EMAIL" \
  -AccessToken "$ACCESS_TOKEN" \
  -Organization "$TENANT_ID"
```

**Expected Output:**
```json
[
  {
    "id": "shared-mailbox-id",
    "email": "support@company.com",
    "displayName": "Support Mailbox",
    "permissions": ["FullAccess", "SendAs"]
  }
]
```

### Testing End-to-End Integration

#### 1. Test PowerShell Availability

```bash
# Verify PowerShell Core
pwsh -Version
# Should show: PowerShell 7.x.x

# Verify Exchange Online module
pwsh -Command "Get-Module -ListAvailable ExchangeOnlineManagement"
# Should list the module with version 3.x
```

#### 2. Test with Debug Logging

```bash
cd projects/ms-365-mcp-server

# Set debug environment
export MS365_MCP_LOG_LEVEL=debug
export MS365_MCP_DEBUG=true
export MS365_MCP_IMPERSONATE_DEBUG=true

# Run mailbox discovery
./auth-list-mailboxes.sh --clear-cache
```

**What to Look For:**
```
[DEBUG] PowerShell service initialized
[DEBUG] Checking for pwsh availability...
[INFO] PowerShell integration enabled and available
[DEBUG] Executing PowerShell script: /path/to/check-mailbox-permissions.ps1
[DEBUG] Script arguments: UserEmail, AccessToken, Organization
[DEBUG] PowerShell stdout length: 1234 characters
[DEBUG] PowerShell script completed successfully
```

#### 3. Test Without PowerShell

Temporarily disable PowerShell to verify fallback behavior:

```bash
# Disable PowerShell
export MS365_POWERSHELL_ENABLED=false

# Run discovery
./auth-list-mailboxes.sh

# Should see:
# [INFO] PowerShell integration explicitly disabled via MS365_POWERSHELL_ENABLED=false
# [INFO] Discovery complete: 1 total (1 personal + 0 shared)

# Re-enable
unset MS365_POWERSHELL_ENABLED
```

#### 4. Test Cache Behavior

```bash
# First run (cache miss)
time ./auth-list-mailboxes.sh --clear-cache
# Note the execution time (2-5 seconds with PowerShell)

# Second run (cache hit)
time ./auth-list-mailboxes.sh
# Should be much faster (< 1 second)
```

### Debugging Common Issues

#### Issue: "pwsh: command not found"

**Solution:**
```bash
# Verify installation
which pwsh

# If not found, install:
brew install --cask powershell  # macOS
# or follow Linux installation in main section
```

#### Issue: "Module ExchangeOnlineManagement not found"

**Solution:**
```bash
# Install the module
pwsh -Command "Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber"

# Verify
pwsh -Command "Get-InstalledModule ExchangeOnlineManagement"
```

#### Issue: "Connect-ExchangeOnline: Invalid access token"

**Possible Causes:**
1. Token expired (tokens expire after 1 hour)
2. Wrong tenant ID
3. Insufficient permissions

**Solutions:**
```bash
# Get fresh token
./auth-login.sh

# Verify tenant ID matches your Azure AD
echo $MS365_MCP_TENANT_ID

# Check token expiration at https://jwt.ms
# Paste your access token there to decode and check 'exp' claim
```

#### Issue: "No shared mailboxes found" (but you expect some)

**Diagnosis Steps:**

1. **Verify User Has Permissions:**
```bash
# Using Exchange Admin Center or PowerShell, check if user actually has permissions
# Admin can run:
pwsh
Connect-ExchangeOnline -UserPrincipalName admin@company.com

Get-MailboxPermission -Identity "shared@company.com" | 
  Where-Object {$_.User -like "*user@company.com*"}

Get-RecipientPermission -Identity "shared@company.com" | 
  Where-Object {$_.Trustee -like "*user@company.com*"}
```

2. **Check Azure AD Permissions:**
```bash
# Ensure these Application permissions are granted:
# - Mail.Read (or Mail.ReadWrite)
# - User.Read.All
# - MailboxSettings.Read

# In Azure Portal, check Admin Consent is granted
```

3. **Test with Different User:**
```bash
# Try with a user you know has shared mailbox access
export MS365_MCP_IMPERSONATE_USER=known-user@company.com
./auth-list-mailboxes.sh --clear-cache
```

### Performance Profiling

#### Measure PowerShell Execution Time

```bash
cd projects/ms-365-mcp-server

# Add timing wrapper
cat > test-performance.sh << 'EOF'
#!/bin/bash
echo "Testing PowerShell performance..."

for i in {1..5}; do
  echo "Run $i:"
  /usr/bin/time -p ./auth-list-mailboxes.sh --clear-cache 2>&1 | grep real
  sleep 2
done
EOF

chmod +x test-performance.sh
./test-performance.sh

# Clean up
rm test-performance.sh
```

#### Expected Timings

| Operation | First Run | Cached |
|-----------|-----------|---------|
| Personal mailbox only | ~200ms | ~200ms |
| + PowerShell (small tenant <50 users) | ~2-3s | ~200ms |
| + PowerShell (medium tenant 50-200 users) | ~3-5s | ~200ms |
| + PowerShell (large tenant >200 users) | ~5-10s | ~200ms |

### Security Testing

#### Verify Token Handling

```bash
# Enable debug to see (truncated) token handling
export MS365_MCP_DEBUG=true

# Run and check logs
./auth-list-mailboxes.sh 2>&1 | grep -i token

# Tokens should NEVER be fully logged
# Only truncated versions like: "token: eyJ0...xyz (truncated)"
```

#### Test Script Injection Protection

PowerShell service prevents script injection by using typed parameters:

```bash
# These should be safely handled (not cause injection):
export MS365_MCP_IMPERSONATE_USER="user@company.com; rm -rf /"
./auth-list-mailboxes.sh
# Should fail safely with email validation error
```

## Future Enhancements

Potential improvements tracked in `TODO.md`:

- Background cache warming for active users
- Parallel PowerShell execution for multiple users
- Real-time permission change detection
- Connection pooling for PowerShell sessions
- Azure Managed Identity support
- Incremental permission checking
- Automated permission validation tests
- Permission change webhook notifications

---

For more information:
- [Exchange Online PowerShell Documentation](https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell)
- [PowerShell Core Installation](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell)
- [MCP Server User Impersonation Guide](./USER_IMPERSONATE.md)
- [MCP Server Setup Guide](./SERVER_SETUP.md)
- [JWT Token Decoder](https://jwt.ms) - Decode and inspect access tokens

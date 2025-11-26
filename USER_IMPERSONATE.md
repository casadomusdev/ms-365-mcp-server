# User Impersonation & Mailbox Discovery

## Overview

When running with **application permissions** (client credentials mode), the MS-365 MCP Server has tenant-wide access. User impersonation allows you to:

- **Scope operations** to a specific user's mailbox context
- **Automatically discover** accessible mailboxes (personal, shared, delegated)
- **Enforce access control** to prevent unauthorized mailbox access  
- **Support multi-tenant** scenarios with per-request user context
- **Cache discovery results** for optimal performance

## How It Works

### Impersonation Resolution

The server resolves which user to impersonate using a 3-tier priority system:

| Priority | Source | Configuration | Use Case |
|----------|--------|---------------|----------|
| **1** | HTTP Header | `MS365_MCP_IMPERSONATE_HEADER` | Per-request dynamic impersonation via reverse proxy |
| **2** | AsyncLocalStorage | Internal context | Set by middleware/internal logic |
| **3** | Environment Variable | `MS365_MCP_IMPERSONATE_USER` | Static server-wide impersonation |

**Resolution Flow:**
```
Request → Check HTTP Header → Check Context → Check Env Var → Use First Found
```

### Mailbox Discovery

Once a user is identified for impersonation, the server automatically discovers which mailboxes they can access using **calendar delegation detection**:

**Detection Strategy:**
- Queries all shared mailboxes in the tenant (unlicensed, enabled accounts)
- For each shared mailbox, checks `/calendar/calendarPermissions` endpoint
- Includes mailboxes where the user has calendar delegation permissions

**Important Limitations:**
- ⚠️ Only detects mailboxes where the user has **calendar delegation**
- ❌ Does NOT detect: SendAs-only permissions, Full Access without calendar access
- ℹ️ Exchange mailbox delegation permissions (SendAs, Full Access) cannot be reliably queried via Microsoft Graph API
- ℹ️ For full delegation detection, Exchange PowerShell or EWS is required (not available in this implementation)

**Why Calendar Delegation?**
- Calendar delegation is the ONLY reliable indicator available via Graph API
- Users with full mailbox access typically also have calendar delegation
- More efficient than checking all tenant users

### Caching Architecture

**Two-Layer Caching System:**

1. **User Validation Cache** (Optional)
   - Validates user exists before impersonation
   - TTL: `MS365_MCP_USER_VALIDATION_TTL` (default: 1 hour)
   - Prevents typos and invalid email addresses

2. **Mailbox Discovery Cache**  
   - Caches discovered mailboxes per user
   - TTL: `MS365_MCP_IMPERSONATE_CACHE_TTL` (default: 1 hour)
   - Dramatically improves performance (10ms vs 2-10 seconds)

## Configuration

### Required for Impersonation

```bash
# Enable client credentials mode (app permissions)
MS365_MCP_CLIENT_SECRET=your-client-secret
MS365_MCP_TENANT_ID=your-tenant-id

# Specify the user to impersonate
MS365_MCP_IMPERSONATE_USER=user@company.com
```

### Optional Configuration

```bash
# Mailbox discovery cache TTL (default: 3600 seconds)
MS365_MCP_IMPERSONATE_CACHE_TTL=3600

# Enable user validation before impersonation (default: false)
MS365_MCP_VALIDATE_IMPERSONATION_USER=true

# User validation cache TTL (default: 3600 seconds)  
MS365_MCP_USER_VALIDATION_TTL=3600

# Custom header name for dynamic impersonation (default: X-Impersonate-User)
MS365_MCP_IMPERSONATE_HEADER=X-Impersonate-User

# Enable `/me` rewriting to `/users/{email}` (default: true)
MS365_MCP_IMPERSONATE_REWRITE_ME=true

# Enable debug logging (default: false)
MS365_MCP_IMPERSONATE_DEBUG=true
```

## Usage Modes

### Mode 1: Static Impersonation

**Best for:** Single-user deployments, development, testing

```bash
# .env
MS365_MCP_CLIENT_SECRET=your-secret
MS365_MCP_IMPERSONATE_USER=support@company.com
```

**Behavior:**
- Server discovers mailboxes at startup
- All requests use this user's context
- Requires server restart to change user

### Mode 2: Dynamic Impersonation

**Best for:** Multi-tenant applications, per-request user context

```bash
# .env  
MS365_MCP_CLIENT_SECRET=your-secret
MS365_MCP_IMPERSONATE_HEADER=X-User-Email
```

```bash
# HTTP Request
curl -X POST http://localhost:3000/mcp \
  -H "Content-Type: application/json" \
  -H "X-User-Email: jane@company.com" \
  -d '{"jsonrpc":"2.0","id":1,"method":"tools/call",...}'
```

**Behavior:**
- Each request can specify a different user
- Mailbox discovery cached per user
- No server restart needed

### Mode 3: Hybrid

**Best for:** Default user with override capability

```bash
# .env
MS365_MCP_CLIENT_SECRET=your-secret
MS365_MCP_IMPERSONATE_USER=default@company.com  
MS365_MCP_IMPERSONATE_HEADER=X-User-Email
```

**Behavior:**
- HTTP header takes precedence if present
- Falls back to env var if header missing
- Best of both worlds

## Architecture

### Service Components

```
ImpersonationResolver
├─> UserValidationCache (optional)
│   └─> GraphClient (validates user existence)
└─> Validates email format

MailboxDiscoveryCache
└─> MailboxDiscoveryService
    ├─> Query shared mailboxes (unlicensed, enabled)
    └─> Check calendar delegation permissions
```

### Data Flow

```
1. Request arrives
2. ImpersonationResolver determines user (header > context > env)
3. Optional: UserValidationCache validates user exists
4. MailboxDiscoveryCache.getMailboxes(user)
   ├─> Cache hit? Return cached mailboxes (< 10ms)
   └─> Cache miss? Run MailboxDiscoveryService (2-10 seconds)
       ├─> Personal mailbox (always included)
       ├─> Query all shared mailboxes in tenant
       ├─> Check calendar delegation for each
       └─> Cache results with TTL
5. Tool executes with allowed mailboxes enforced
```

### Discovered Mailbox Types

**Personal Mailbox**
- User's own mailbox
- Always included
- Type: `'personal'`

**Shared Mailbox**
- Multi-user shared mailbox  
- No license assigned
- Type: `'shared'`

**Delegated Mailbox** 
- Another user's mailbox
- Access via delegation
- Type: `'delegated'`

## User Validation (Optional)

Enable to catch configuration errors early:

```bash
MS365_MCP_VALIDATE_IMPERSONATION_USER=true
```

### Validation Process

1. **Format check**: Email regex validation
2. **Existence check**: Query `/users/{email}` endpoint
3. **Cache result**: Store for `MS365_MCP_USER_VALIDATION_TTL`
4. **Throw on failure**: Clear error with source information

### Error Examples

```
# Invalid format
Impersonation failed (source: environment variable 'MS365_MCP_IMPERSONATE_USER'): 
Invalid email format: 'not-an-email'

# User not found
Impersonation failed (source: HTTP header 'x-impersonate-user'): 
User not found: nonexistent@company.com
```

## Security Considerations

### HTTP Header Trust

⚠️ **Critical**: Only trust impersonation headers from a reverse proxy you control

**Recommended setup:**

```nginx
# NGINX - Only accept header from authenticated source
location /mcp {
    # Authenticate user first
    auth_request /auth;
    
    # Set impersonation header from auth result
    auth_request_set $user_email $upstream_http_x_user_email;
    proxy_set_header X-Impersonate-User $user_email;
    
    # Strip any client-provided header
    proxy_set_header X-Impersonate-User-Client "";
    
    proxy_pass http://ms365-mcp-server:3000;
}
```

### Audit Logging

All impersonation context is logged:

```typescript
// Automatically logged at debug level
{
  "impersonatedUser": "user@company.com",
  "source": "http-header", // or "environment" or "context"
  "mailboxesDiscovered": 3,
  "cacheHit": false
}
```

## Performance

### Initial Discovery

- **First request**: 2-10 seconds (depends on tenant size)
- **Concurrent limit**: 5 simultaneous API calls
- **Timeout**: 5 seconds per mailbox check
- **Tenant size impact**: ~50-100 users scanned

### Cached Lookups

- **Cache hit**: < 10ms
- **TTL**: Configurable (default 1 hour)
- **Memory**: ~1KB per cached user

### Optimization Tips

```bash
# Increase cache TTL to reduce discovery frequency
MS365_MCP_IMPERSONATE_CACHE_TTL=7200  # 2 hours

# Reduce if permissions change frequently
MS365_MCP_IMPERSONATE_CACHE_TTL=1800  # 30 minutes
```

## Troubleshooting

### No Shared Mailboxes Discovered

**Symptom:** Only personal mailbox found

**Causes:**
1. User has no calendar delegation to shared mailboxes
2. User has SendAs-only permissions (not detectable via Graph API)
3. User has Full Access without calendar delegation (rare)
4. Missing Graph API permissions
5. Discovery timing out

**Solutions:**

**1. Verify the user actually has calendar delegation:**
- In Exchange Admin Center or PowerShell, check if the user has calendar permissions on shared mailboxes
- Calendar delegation is typically granted along with full mailbox access, but not always

**2. Verify Graph API permissions are granted in Azure AD:**

**Required Application Permissions (Client Credentials Mode):**
- `Calendars.Read` - Read calendars in all mailboxes (for calendar delegation detection)
- `User.Read.All` - Read all users (for shared mailbox enumeration and user validation)

**Optional but Recommended:**
- `Mail.Read` - Allows testing actual mailbox access
- `MailboxSettings.Read` - Allows reading mailbox metadata

These permissions are marked as **required** in SERVER_SETUP.md. The core discovery functionality requires `Calendars.Read` and `User.Read.All` at minimum.

Then enable debug logging to see which strategies are working:

```bash
MS365_MCP_LOG_LEVEL=debug
MS365_MCP_IMPERSONATE_DEBUG=true
```

### Discovery Too Slow

**Symptom:** Initial requests take > 10 seconds

**Causes:**
1. Large tenant (many users)
2. Network latency
3. Graph API throttling

**Solutions:**
```bash
# Increase cache TTL  
MS365_MCP_IMPERSONATE_CACHE_TTL=7200

# Note: First request will always be slow
# Subsequent requests use cache (< 10ms)
```

### Cache Not Working

**Symptom:** Every request triggers discovery

**Diagnosis:**
```bash
# Enable debug to see cache hits/misses
MS365_MCP_DEBUG=true
MS365_MCP_IMPERSONATE_DEBUG=true

# Look for in logs:
# "Cache hit for user@company.com (age: 120s)"
# vs
# "Cache miss for user@company.com, discovering..."
```

**Solutions:**
```bash
# Ensure consistent email format (lowercase)
MS365_MCP_IMPERSONATE_USER=user@company.com  # not User@Company.com

# Increase TTL
MS365_MCP_IMPERSONATE_CACHE_TTL=3600

# Check container isn't restarting (clears memory cache)
docker logs ms365-mcp-server
```

### Validation Errors

**Symptom:** "User not found" errors

**Solutions:**
```bash
# Verify user exists
./auth-list-mailboxes.sh

# Disable validation temporarily
MS365_MCP_VALIDATE_IMPERSONATION_USER=false

# Check for typos
MS365_MCP_IMPERSONATE_USER=correct.email@company.com
```

## CLI Tools

### List Mailboxes for Impersonated User

```bash
# Shows discovered mailboxes
./auth-list-mailboxes.sh
```

**Example output:**
```json
{
  "success": true,
  "userEmail": "user@company.com",
  "mailboxes": [
    {
      "id": "user-id-123",
      "type": "personal",
      "displayName": "John Doe",
      "email": "user@company.com",
      "isPrimary": true
    },
    {
      "id": "shared-id-456",
      "type": "shared",
      "displayName": "Support Mailbox",
      "email": "support@company.com",
      "isPrimary": false
    }
  ]
}
```

### Verify Authentication

```bash
# Tests authentication and impersonation setup
./auth-verify.sh
```

## Examples

### Basic Setup

```bash
# .env
MS365_MCP_CLIENT_ID=your-app-id
MS365_MCP_CLIENT_SECRET=your-secret
MS365_MCP_TENANT_ID=your-tenant-id
MS365_MCP_IMPERSONATE_USER=support@company.com
```

### With Validation

```bash
# Validate user before allowing impersonation
MS365_MCP_VALIDATE_IMPERSONATION_USER=true
MS365_MCP_USER_VALIDATION_TTL=3600
```

### Multi-Tenant SaaS

```bash
# Reverse proxy sets user header per request
MS365_MCP_IMPERSONATE_HEADER=X-Authenticated-User

# Each request operates in different user context
# Headers: {"X-Authenticated-User": "tenant1@company.com"}
# Headers: {"X-Authenticated-User": "tenant2@company.com"}
```

### Debug Discovery

```bash
# See full discovery process
MS365_MCP_LOG_LEVEL=debug
MS365_MCP_DEBUG=true
MS365_MCP_IMPERSONATE_DEBUG=true

# Run and check logs
docker-compose up
```

## Migration from Phase 0

If you were using the old `MS365_MCP_IMPERSONATE_ALLOWED_MAILBOXES` variable:

**Old configuration:**
```bash
MS365_MCP_IMPERSONATE_USER=user@company.com
MS365_MCP_IMPERSONATE_ALLOWED_MAILBOXES=shared1@company.com,shared2@company.com
```

**New (automatic discovery):**
```bash
MS365_MCP_IMPERSONATE_USER=user@company.com
# No allowlist needed - mailboxes are automatically discovered!
```

**Benefits of new approach:**
- Automatic discovery - no manual configuration
- Always up-to-date as permissions change
- Discovers all accessible mailboxes, not just configured ones
- Validates actual access, not just configuration

## API Reference

### Environment Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `MS365_MCP_IMPERSONATE_USER` | - | Email of user to impersonate |
| `MS365_MCP_IMPERSONATE_HEADER` | `X-Impersonate-User` | Header name for dynamic impersonation |
| `MS365_MCP_IMPERSONATE_CACHE_TTL` | `3600` | Mailbox discovery cache TTL (seconds) |
| `MS365_MCP_VALIDATE_IMPERSONATION_USER` | `false` | Enable user validation |
| `MS365_MCP_USER_VALIDATION_TTL` | `3600` | User validation cache TTL (seconds) |
| `MS365_MCP_IMPERSONATE_REWRITE_ME` | `true` | Rewrite `/me` to `/users/{email}` |
| `MS365_MCP_IMPERSONATE_DEBUG` | `false` | Enable impersonation debug logging |

### Cache Statistics

Access via internal API (development mode):

```typescript
// Get cache stats
const stats = mailboxDiscoveryCache.getCacheStats();

// Returns:
{
  size: 5,  
  entries: [
    {
      email: "user@company.com",
      mailboxCount: 3,
      age: 120,        // seconds since cached
      expiresIn: 3480  // seconds until expiry
    }
  ]
}

// Clear cache
mailboxDiscoveryCache.clearCache();
```

## See Also

- [AUTH.md](AUTH.md) - Authentication setup and configuration
- [SERVER_SETUP.md](SERVER_SETUP.md) - Complete server setup guide
- [QUICK_START.md](QUICK_START.md) - Quick start guide

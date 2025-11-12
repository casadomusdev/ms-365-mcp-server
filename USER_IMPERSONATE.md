# User Impersonation Implementation Plan

## Overview

This document outlines the implementation plan for adding user impersonation capabilities to the MS-365 MCP Server when running in **Client Credentials Mode**. User impersonation allows the server to restrict access to only those mailboxes/calendars/files that a specific user has permission to access, even though the server has app-level permissions to the entire tenant.

## Use Cases

1. **Multi-tenant SaaS Application**: Different end-users accessing the MCP server should only see their own data
2. **Delegated Access Control**: Enforce user-level permissions even with app-level authentication
3. **Compliance & Auditing**: Track which user's context is being used for each operation
4. **Development & Testing**: Test different permission scenarios without changing app configuration

## Feature Modes

### Mode 1: Static User Impersonation (Environment Variable)

**Configuration**: Set `MS365_MCP_IMPERSONATE_USER=user@domain.com` in `.env`

**Behavior**:
- Server starts and discovers mailboxes accessible to the specified user
- All operations are restricted to this user's access scope
- Impersonated user remains the same for the lifetime of the server
- Requires server restart to change impersonated user

**Best for**:
- Single-user deployments
- Development and testing
- Simple setups where one service account's permissions are sufficient

### Mode 2: Dynamic User Impersonation (HTTP Header)

**Configuration**: Set `MS365_MCP_IMPERSONATE_HEADER=X-Impersonate-User` in `.env`

**Behavior**:
- Server reads the specified HTTP header from each incoming request
- Mailbox discovery is performed on-demand and cached per user
- Each request can impersonate a different user
- No server restart needed to change users

**Best for**:
- Multi-tenant applications
- Per-request user context
- Integration with reverse proxies/API gateways that set user headers
- Dynamic permission enforcement

### Mode 3: Hybrid (Both Modes)

**Configuration**: Set both environment variables

**Behavior**:
- HTTP header takes precedence if present
- Falls back to env var user if header is missing
- Allows default user with per-request override capability

## Technical Architecture

### 1. Mailbox Discovery & Caching

**Discovery Process**:

When an impersonated user is specified, the server needs to discover which mailboxes they can access:

```typescript
interface UserMailboxCache {
  userEmail: string;
  discoveredAt: number;          // Timestamp
  expiresAt: number;             // TTL for cache
  allowedMailboxes: MailboxInfo[];
}

interface MailboxInfo {
  id: string;
  email: string;
  displayName: string;
  type: 'personal' | 'shared' | 'delegated';
  permissions: string[];         // e.g., ['read', 'write', 'send']
}
```

**Discovery Methods**:

1. **Personal Mailbox**: `/users/{email}` - Always accessible
2. **Delegated Mailboxes**: `/users/{email}/mailFolders` - Check for FullAccess/SendAs permissions
3. **Shared Mailboxes**: Query `/users/{email}/mailSettings` - Check for delegate access
4. **Group Memberships**: `/users/{email}/memberOf` - Check for group-based shared mailboxes
5. **Application-level Query**: Use app permissions to query `/users` and test access

**Caching Strategy**:

```typescript
class MailboxDiscoveryCache {
  private cache: Map<string, UserMailboxCache>;
  private cacheTTL: number; // Default: 1 hour
  
  async getMailboxes(userEmail: string): Promise<MailboxInfo[]> {
    const cached = this.cache.get(userEmail);
    
    // Return cached if valid
    if (cached && cached.expiresAt > Date.now()) {
      return cached.allowedMailboxes;
    }
    
    // Discover mailboxes
    const mailboxes = await this.discoverMailboxes(userEmail);
    
    // Cache results
    this.cache.set(userEmail, {
      userEmail,
      discoveredAt: Date.now(),
      expiresAt: Date.now() + this.cacheTTL,
      allowedMailboxes: mailboxes
    });
    
    return mailboxes;
  }
  
  private async discoverMailboxes(userEmail: string): Promise<MailboxInfo[]> {
    // Implementation of discovery logic
  }
}
```

### 2. Access Validation Layer

**Validation Point**: Before every Graph API call

**Validation Logic**:

```typescript
class AccessValidator {
  private cacheManager: MailboxDiscoveryCache;
  
  async validateMailboxAccess(
    impersonatedUser: string,
    targetMailbox: string,
    operation: 'read' | 'write' | 'send'
  ): Promise<boolean> {
    // Get allowed mailboxes for this user
    const allowedMailboxes = await this.cacheManager.getMailboxes(impersonatedUser);
    
    // Find the target mailbox
    const mailbox = allowedMailboxes.find(
      mb => mb.email.toLowerCase() === targetMailbox.toLowerCase() ||
            mb.id === targetMailbox
    );
    
    if (!mailbox) {
      logger.warn(`User ${impersonatedUser} denied access to ${targetMailbox}`);
      return false;
    }
    
    // Check if user has required permission
    if (!mailbox.permissions.includes(operation)) {
      logger.warn(
        `User ${impersonatedUser} lacks ${operation} permission for ${targetMailbox}`
      );
      return false;
    }
    
    return true;
  }
}
```

### 3. Request Context Management

**HTTP Mode Request Processing**:

```typescript
class ImpersonationContext {
  private static asyncLocalStorage = new AsyncLocalStorage<string>();
  
  static setImpersonatedUser(email: string): void {
    this.asyncLocalStorage.enterWith(email);
  }
  
  static getImpersonatedUser(): string | undefined {
    return this.asyncLocalStorage.getStore();
  }
  
  static withUser<T>(email: string, fn: () => Promise<T>): Promise<T> {
    return this.asyncLocalStorage.run(email, fn);
  }
}
```

**HTTP Server Middleware** (for HTTP mode):

```typescript
// In HTTP server handler
app.use((req, res, next) => {
  const impersonateHeader = process.env.MS365_MCP_IMPERSONATE_HEADER;
  const staticUser = process.env.MS365_MCP_IMPERSONATE_USER;
  
  let impersonatedUser: string | undefined;
  
  // Priority: HTTP header > static env var
  if (impersonateHeader && req.headers[impersonateHeader.toLowerCase()]) {
    impersonatedUser = req.headers[impersonateHeader.toLowerCase()] as string;
  } else if (staticUser) {
    impersonatedUser = staticUser;
  }
  
  if (impersonatedUser) {
    ImpersonationContext.setImpersonatedUser(impersonatedUser);
    logger.info(`Request impersonating user: ${impersonatedUser}`);
  }
  
  next();
});
```

### 4. Integration with Tool Calls

**Modify Tool Execution**:

Every tool that accesses mailboxes/calendars/files needs validation:

```typescript
// Example: list-mail-messages tool
async function listMailMessages(params: {
  mailbox?: string;
  folder: string;
  max_results: number;
}) {
  const impersonatedUser = ImpersonationContext.getImpersonatedUser();
  
  // If impersonation is active, validate access
  if (impersonatedUser) {
    const targetMailbox = params.mailbox || impersonatedUser;
    
    const hasAccess = await accessValidator.validateMailboxAccess(
      impersonatedUser,
      targetMailbox,
      'read'
    );
    
    if (!hasAccess) {
      throw new Error(
        `Access denied: User ${impersonatedUser} cannot read mailbox ${targetMailbox}`
      );
    }
  }
  
  // Proceed with Graph API call
  // ...
}
```

## Implementation Steps

### Phase 1: Core Infrastructure

**Files to Create/Modify**:

1. **src/impersonation/MailboxDiscoveryCache.ts** (NEW)
   - Mailbox discovery logic
   - Cache management
   - TTL handling

2. **src/impersonation/AccessValidator.ts** (NEW)
   - Access validation logic
   - Permission checking
   - Logging

3. **src/impersonation/ImpersonationContext.ts** (NEW)
   - AsyncLocalStorage for request context
   - Context management utilities

4. **src/impersonation/index.ts** (NEW)
   - Export all impersonation classes

### Phase 2: AuthManager Integration

**Files to Modify**:

1. **src/auth.ts**
   - Add impersonation configuration properties
   - Initialize MailboxDiscoveryCache
   - Add method: `async discoverUserMailboxes(email: string)`
   - Add method: `async validateAccess(user, mailbox, operation)`

### Phase 3: HTTP Server Integration

**Files to Modify**:

1. **src/server.ts**
   - Add HTTP middleware for header extraction
   - Set up ImpersonationContext
   - Log impersonation info

2. **src/index.ts**
   - Log impersonation mode at startup
   - Initialize AccessValidator

### Phase 4: Tool Integration

**Files to Modify**:

All tool implementation files need validation checks:

1. **src/tools/mail.ts**
   - Add validation before Graph API calls
   - Handle access denied errors

2. **src/tools/calendar.ts**
   - Add validation for calendar access

3. **src/tools/files.ts**
   - Add validation for OneDrive/SharePoint access

4. **Other tool files**
   - Apply same pattern

### Phase 5: Configuration & Documentation

**Files to Create/Modify**:

1. **.env.example**
   - Add `MS365_MCP_IMPERSONATE_USER`
   - Add `MS365_MCP_IMPERSONATE_HEADER`
   - Add `MS365_MCP_IMPERSONATE_CACHE_TTL`

2. **SERVER_SETUP.md**
   - Add "User Impersonation" section
   - Document both modes
   - Provide configuration examples
   - Security considerations

3. **README.md**
   - Update feature list
   - Add impersonation examples

### Phase 6: Testing

**Test Scenarios**:

1. **Static Impersonation Tests**
   - Verify mailbox discovery works
   - Test access validation (allow/deny)
   - Test cache expiration and refresh

2. **Dynamic Impersonation Tests**
   - Test header parsing
   - Test per-request user switching
   - Test fallback to static user

3. **Edge Cases**
   - Invalid user email
   - User with no mailbox access
   - Missing user in Azure AD
   - Expired cache handling

4. **Performance Tests**
   - Cache hit rate
   - Discovery performance
   - Validation overhead

## Configuration Reference

### Environment Variables

```bash
# Static user impersonation (optional)
MS365_MCP_IMPERSONATE_USER=user@domain.com

# Dynamic user impersonation via HTTP header (optional)
# Header name is case-insensitive
MS365_MCP_IMPERSONATE_HEADER=X-Impersonate-User

# Cache TTL in seconds (default: 3600 = 1 hour)
MS365_MCP_IMPERSONATE_CACHE_TTL=3600

# Enable verbose impersonation logging (default: false)
MS365_MCP_IMPERSONATE_DEBUG=true
```

### Usage Examples

**Example 1: Static Impersonation**

```bash
# .env
MS365_MCP_CLIENT_SECRET=your-secret
MS365_MCP_IMPERSONATE_USER=john.doe@company.com

# Server restricts all operations to john.doe@company.com's access scope
```

**Example 2: Dynamic Impersonation**

```bash
# .env
MS365_MCP_CLIENT_SECRET=your-secret
MS365_MCP_IMPERSONATE_HEADER=X-User-Email
```

```bash
# HTTP Request
curl -X POST http://localhost:3000/mcp \
  -H "Content-Type: application/json" \
  -H "X-User-Email: jane.smith@company.com" \
  -d '{"jsonrpc":"2.0","id":1,"method":"tools/call","params":{"name":"list-mail-messages","arguments":{"folder":"inbox"}}}'

# Server impersonates jane.smith@company.com for this request
```

**Example 3: Hybrid Mode**

```bash
# .env
MS365_MCP_CLIENT_SECRET=your-secret
MS365_MCP_IMPERSONATE_USER=default.user@company.com
MS365_MCP_IMPERSONATE_HEADER=X-User-Email

# Default: uses default.user@company.com
# With header: uses header value
# Without header: falls back to default.user@company.com
```

### Reverse Proxy Integration

**NGINX Example**:

```nginx
location /mcp {
    # Extract authenticated user from session/JWT/etc
    # and pass as header to MCP server
    
    proxy_set_header X-Impersonate-User $authenticated_user_email;
    proxy_pass http://ms365-mcp-server:3000;
}
```

**Traefik Example**:

```yaml
http:
  middlewares:
    add-user-header:
      headers:
        customRequestHeaders:
          X-Impersonate-User: "user@domain.com"
```

## Security Considerations

### 1. Header Spoofing Prevention

**Problem**: Malicious clients could set impersonation headers

**Solutions**:
- **Option A**: Only allow impersonation headers from trusted reverse proxy
  - Configure allowed proxy IPs
  - Strip headers from direct client connections
  
- **Option B**: Require authentication at reverse proxy level
  - Reverse proxy validates user identity
  - Only proxy sets trusted headers
  
- **Option C**: Use signed headers
  - Proxy signs header with shared secret
  - MCP server verifies signature

**Recommended Implementation**:

```typescript
// Only trust headers from specific source IPs
const TRUSTED_PROXIES = process.env.TRUSTED_PROXY_IPS?.split(',') || [];

function validateImpersonationHeader(req: Request): string | undefined {
  const impersonateHeader = process.env.MS365_MCP_IMPERSONATE_HEADER;
  if (!impersonateHeader) return undefined;
  
  const clientIP = req.socket.remoteAddress;
  
  // If trusted proxies configured, only accept headers from them
  if (TRUSTED_PROXIES.length > 0) {
    if (!TRUSTED_PROXIES.includes(clientIP)) {
      logger.warn(
        `Rejecting impersonation header from untrusted IP: ${clientIP}`
      );
      return undefined;
    }
  }
  
  return req.headers[impersonateHeader.toLowerCase()] as string;
}
```

### 2. Audit Logging

Log all impersonation activities:

```typescript
logger.info({
  event: 'impersonation_access',
  impersonatedUser: 'user@domain.com',
  targetResource: 'mailbox:other@domain.com',
  operation: 'read',
  result: 'allowed',
  timestamp: new Date().toISOString()
});
```

### 3. Permission Caching

**Risk**: User permissions change but cache is stale

**Mitigations**:
- Configurable TTL (default: 1 hour)
- Manual cache invalidation endpoint
- Monitor for permission changes via webhooks
- Shorter TTL for sensitive operations

### 4. Rate Limiting

**Risk**: Discovery process can be expensive

**Mitigations**:
- Rate limit mailbox discovery per user
- Implement request throttling
- Monitor for abuse patterns

## Performance Optimization

### 1. Lazy Discovery

Don't discover all mailboxes upfront - discover on-demand:

```typescript
async validateAccess(user, mailbox, operation) {
  // Check if we've already discovered this specific mailbox
  if (this.hasMailboxInCache(user, mailbox)) {
    return this.validateFromCache(user, mailbox, operation);
  }
  
  // Discover just this mailbox
  const hasAccess = await this.discoverSingleMailbox(user, mailbox);
  this.cacheMailbox(user, mailbox, hasAccess);
  
  return hasAccess;
}
```

### 2. Batch Discovery

When discovery is needed, batch Graph API calls:

```typescript
// Use $batch endpoint to reduce round trips
const batch = {
  requests: [
    { id: '1', method: 'GET', url: `/users/${email}` },
    { id: '2', method: 'GET', url: `/users/${email}/mailFolders?$top=1` },
    { id: '3', method: 'GET', url: `/users/${email}/memberOf` }
  ]
};
```

### 3. Background Refresh

Refresh cache in background before expiration:

```typescript
setInterval(async () => {
  const expiringCache = this.getExpiringCache(threshold: 5 * 60 * 1000); // 5 min
  
  for (const userEmail of expiringCache) {
    await this.refreshMailboxCache(userEmail);
  }
}, 60000); // Check every minute
```

## Error Handling

### Access Denied Scenarios

**Error Response Format**:

```json
{
  "jsonrpc": "2.0",
  "id": 1,
  "error": {
    "code": -32000,
    "message": "Access denied: User john@company.com cannot access mailbox jane@company.com",
    "data": {
      "impersonatedUser": "john@company.com",
      "requestedMailbox": "jane@company.com",
      "reason": "mailbox_not_accessible"
    }
  }
}
```

### Invalid User

```json
{
  "error": {
    "code": -32001,
    "message": "Invalid impersonated user: user@domain.com not found in Azure AD",
    "data": {
      "impersonatedUser": "user@domain.com",
      "reason": "user_not_found"
    }
  }
}
```

## Monitoring & Metrics

### Key Metrics to Track

1. **Cache Performance**
   - Hit rate
   - Miss rate
   - Eviction rate

2. **Discovery Performance**
   - Discovery latency (p50, p95, p99)
   - Number of discoveries per minute
   - Failed discoveries

3. **Access Validation**
   - Validation latency
   - Access denied rate by user
   - Access granted rate by user

4. **User Activity**
   - Unique impersonated users per hour/day
   - Most active users
   - Mailbox access patterns

## Migration Path

### For Existing Deployments

**Step 1**: Deploy impersonation code (disabled by default)
```bash
# No env vars set - impersonation disabled
# All existing functionality works as before
```

**Step 2**: Test with static impersonation
```bash
MS365_MCP_IMPERSONATE_USER=test.user@domain.com
# Test mailbox discovery and validation
```

**Step 3**: Enable dynamic impersonation
```bash
MS365_MCP_IMPERSONATE_HEADER=X-User-Email
# Test with different users via headers
```

**Step 4**: Deploy to production with monitoring
```bash
# Enable with conservative cache TTL
MS365_MCP_IMPERSONATE_CACHE_TTL=300  # 5 minutes initially
MS365_MCP_IMPERSONATE_DEBUG=true      # Verbose logging initially
```

## Future Enhancements

### 1. Permission-Based Filtering

Instead of all-or-nothing, filter results based on permissions:

```typescript
// Return only mailboxes user can read
const mailboxes = await getAllMailboxes();
return mailboxes.filter(mb => userHasPermission(user, mb, 'read'));
```

### 2. Delegation Chains

Support impersonation chains (admin -> service account -> end user):

```typescript
MS365_MCP_IMPERSONATION_CHAIN=admin@domain.com,service@domain.com
X-Impersonate-User: end.user@domain.com
// Validate: admin can impersonate service, service can access end user's mailbox
```

### 3. Temporary Access Grants

Time-limited access to specific mailboxes:

```typescript
// Grant jane temporary access to john's inbox for 1 hour
grantTemporaryAccess('jane@domain.com', 'john@domain.com:inbox', duration: 3600);
```

### 4. Impersonation Audit Dashboard

Web UI showing:
- Who is impersonating whom
- Access patterns
- Denied access attempts
- Cache statistics

## Summary

The user impersonation feature provides:

✅ **Flexibility**: Static env var or dynamic HTTP header
✅ **Security**: Enforces user-level permissions with app-level auth
✅ **Performance**: Smart caching with configurable TTL
✅ **Compatibility**: Works in both STDIO and HTTP modes
✅ **Scalability**: Designed for multi-tenant scenarios
✅ **Auditability**: Comprehensive logging and metrics

**When to Use**:
- Multi-tenant SaaS applications
- Compliance requirements for user-level access control
- Testing different permission scenarios
- Integrating with existing authentication systems

**Trade-offs**:
- Adds complexity to authentication flow
- Requires additional Graph API calls for discovery
- Cache management overhead
- Need to handle permission changes

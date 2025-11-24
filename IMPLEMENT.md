# Mailbox Discovery & Impersonation Enhancement

## GOAL

Fix the broken MailboxDiscoveryCache that doesn't actually discover shared mailboxes, and improve the impersonation system with proper validation and documentation.

## ANALYSIS

### Current Problems

1. **MailboxDiscoveryCache doesn't discover**: The `discoverMailboxes()` method only reads the env var `MS365_MCP_IMPERSONATE_ALLOWED_MAILBOXES` and never queries the Microsoft Graph API to find shared mailboxes
2. **Code duplication**: Full discovery logic exists in `auth.ts:listImpersonatedMailboxes()` with sophisticated 4-strategy discovery, but isn't reused
3. **No validation**: Impersonation accepts any email address without checking if the user exists in the tenant
4. **Unclear precedence**: The order of header vs context vs env var for impersonation is confusing and not documented
5. **Performance concerns**: Multiple Graph API calls are slow, need caching strategy
6. **Poor documentation**: Impersonation behavior and configuration not documented

### Existing Assets

**Full discovery logic** in `auth.ts:listImpersonatedMailboxes()` with 4 strategies:
- **Strategy 1**: Check calendar delegate permissions
- **Strategy 2**: Detect shared mailboxes (mailboxes with no licenses)  
- **Strategy 3**: Check SendAs permissions
- **Strategy 4**: Verify actual mailbox access

**GraphClient** provides authenticated Graph API requests via `makeRequest(endpoint, options)`

**Caching infrastructure** already exists in MailboxDiscoveryCache (TTL-based Map)

### Architecture Discovery

```
Current State:
- GraphClient (graph-client.ts) - Handles authenticated Graph API requests
- AuthManager (auth.ts) - Contains listImpersonatedMailboxes() with full discovery logic
- MailboxDiscoveryCache (impersonation/) - Currently just a stub reading env vars

Needed State:
- Extract discovery logic into MailboxDiscoveryService
- Make MailboxDiscoveryCache use the service (caching layer)
- Make AuthManager use the service (avoid duplication)
- Add ImpersonationResolver for source precedence
- Optionally add UserValidationCache for performance
```

## IMPLEMENTATION

### Phase 1: Core Discovery Service

Create centralized mailbox discovery service to eliminate code duplication.

**File: `src/impersonation/MailboxDiscoveryService.ts`** (NEW)

Responsibilities:
- Extract all 4 discovery strategies from `auth.ts:listImpersonatedMailboxes()`
- Implement as standalone service with GraphClient dependency
- Return `MailboxInfo[]` containing personal mailbox + discovered shared/delegated mailboxes
- Include performance optimizations: timeouts (5s), concurrent request limiting (5 requests)

Key methods:
```typescript
export class MailboxDiscoveryService {
  constructor(private graphClient: GraphClient) {}
  
  async discoverMailboxes(userEmail: string): Promise<MailboxInfo[]>
  
  private async checkCalendarPermissions(userId: string, impersonatedUserEmail: string): Promise<boolean>
  private async detectSharedMailbox(user: any): Promise<boolean>
  private async checkSendAsPermissions(userId: string, impersonatedUserEmail: string): Promise<boolean>
  private async verifyMailboxAccess(userPrincipalName: string): Promise<boolean>
}
```

Documentation requirements:
- JSDoc for each strategy explaining how it works
- Clear comments on performance trade-offs
- Error handling documentation

---

### Phase 2: User Validation Service (Optional, Performance-Aware)

Add optional user existence validation with caching.

**File: `src/impersonation/UserValidationCache.ts`** (NEW)

Purpose: When enabled, validates that impersonated user exists before attempting discovery

Features:
- Format validation (email regex) - always enabled, instant
- User existence check via Graph API - optional, slow but cached
- TTL-based caching to minimize API calls

**New Environment Variables:**
- `MS365_MCP_VALIDATE_IMPERSONATION_USER=true|false` (default: `false`)
- `MS365_MCP_USER_VALIDATION_TTL=3600` (seconds, default: 1 hour)

Performance characteristics:
- **Disabled (default)**: Zero overhead, no additional API calls
- **Enabled**: Single API call per unique user, cached for TTL duration
- **Cache invalidation**: TTL-based, prevents stale data

Key methods:
```typescript
export class UserValidationCache {
  private cache = new Map<string, { exists: boolean; validatedAt: number; expiresAt: number }>();
  
  constructor(
    private graphClient: GraphClient,
    private ttlMs: number = 3600000 // 1 hour default
  ) {}
  
  async validateUser(email: string): Promise<boolean>
  private isValidEmailFormat(email: string): boolean
  private async checkUserExists(email: string): Promise<boolean>
}
```

Trade-offs:
- **Disabled**: Maximum performance, validation happens naturally when Graph API rejects requests
- **Enabled**: Fail fast with clear errors, better UX, slightly slower on first request per user

---

### Phase 3: Refactor MailboxDiscoveryCache

Transform from stub into proper caching layer using the service.

**File: `src/impersonation/MailboxDiscoveryCache.ts`** (REFACTOR)

Changes:
- **Remove**: All references to env var `MS365_MCP_IMPERSONATE_ALLOWED_MAILBOXES` (obsolete)
- **Add**: MailboxDiscoveryService instantiation and delegation
- **Keep**: Existing cache logic with TTL (Map-based, configurable via `MS365_MCP_IMPERSONATE_CACHE_TTL`)

New structure:
```typescript
export class MailboxDiscoveryCache {
  private cache = new Map<string, UserMailboxCache>();
  private service: MailboxDiscoveryService;
  private cacheTTLms: number;
  
  constructor(graphClient: GraphClient) {
    this.service = new MailboxDiscoveryService(graphClient);
    const ttlSec = Number(process.env.MS365_MCP_IMPERSONATE_CACHE_TTL || '3600');
    this.cacheTTLms = Math.max(60, ttlSec) * 1000;
  }
  
  async getMailboxes(userEmail: string): Promise<MailboxInfo[]> {
    // Check cache validity
    // If miss/expired: call service.discoverMailboxes(userEmail)
    // Cache results
    // Return mailboxes
  }
}
```

Responsibilities:
- Caching layer only - no discovery logic
- Delegates to MailboxDiscoveryService for actual discovery
- Manages cache invalidation based on TTL

---

### Phase 4: Impersonation Source Selection & Validation

Implement clear precedence with optional validation.

**File: `src/impersonation/ImpersonationResolver.ts`** (NEW)

Purpose: Resolve impersonated user from multiple sources with clear, documented precedence

**Precedence (highest to lowest):**
1. **HTTP Header** (`X-Impersonate-User` from `_meta.headers`)
2. **AsyncLocalStorage Context** (set by HTTP middleware)
3. **Environment Variable** (`MS365_MCP_IMPERSONATE_USER`)

**Validation rules:**
- **Format validation** (email regex): Always enabled, instant
- **User existence check**: Optional via UserValidationCache, enabled with `MS365_MCP_VALIDATE_IMPERSONATION_USER=true`

**Fallback behavior:**
- Empty/whitespace-only values are skipped, tries next source
- Invalid format → throws descriptive error immediately
- User not found → throws descriptive error (if validation enabled)
- No valid source found → throws descriptive error

Key methods:
```typescript
export class ImpersonationResolver {
  constructor(private validationCache?: UserValidationCache) {}
  
  async resolveImpersonatedUser(
    metaHeaders?: Record<string, string>,
    storageContext?: string,
    envVar?: string
  ): Promise<{ email: string; source: 'meta-header' | 'http-context' | 'env-var' }>
  
  private isValidEmail(email: string): boolean
  private async validateIfEnabled(email: string): Promise<void>
}
```

**Error Messages (clear and actionable):**
- `"Impersonation failed: Invalid email format 'xyz' from X-Impersonate-User header"`
- `"Impersonation failed: User 'user@domain.com' not found in tenant (source: environment variable)"`
- `"Impersonation not configured: No valid user found in header, context, or MS365_MCP_IMPERSONATE_USER"`

---

### Phase 5: Integration Updates

Update integration points to use new architecture.

#### File: `src/graph-tools.ts` (UPDATE)

In `registerGraphTools()` function:

```typescript
// Initialize services
const cache = new MailboxDiscoveryCache(graphClient);

const validationCache = process.env.MS365_MCP_VALIDATE_IMPERSONATION_USER === 'true'
  ? new UserValidationCache(graphClient)
  : undefined;

const resolver = new ImpersonationResolver(validationCache);

// In each tool handler:
// Replace existing impersonation logic with:
const { email: impersonated, source } = await resolver.resolveImpersonatedUser(
  storedMeta?.headers,
  ImpersonationContext.getImpersonatedUser(),
  process.env.MS365_MCP_IMPERSONATE_USER
);

logger.info(`Impersonation resolved: ${impersonated} (source: ${source})`);

const allowed = await cache.getMailboxes(impersonated);
const allowedEmails = allowed.map(m => m.email.toLowerCase());
```

Changes:
- Remove verbose impersonation source detection logic
- Remove inline discovery cache instantiation per request
- Add clear logging of resolved user and source
- Delegate all complexity to resolver and cache services

#### File: `src/auth.ts` (UPDATE)

In `listImpersonatedMailboxes()` method:

```typescript
async listImpersonatedMailboxes(): Promise<any> {
  try {
    const impersonateUser = process.env.MS365_MCP_IMPERSONATE_USER?.trim();
    
    if (!impersonateUser) {
      return {
        success: false,
        error: 'MS365_MCP_IMPERSONATE_USER environment variable is not configured',
      };
    }

    // Create temporary GraphClient for this CLI operation
    const tempGraphClient = new GraphClient(this);
    
    // Use centralized discovery service
    const service = new MailboxDiscoveryService(tempGraphClient);
    const mailboxes = await service.discoverMailboxes(impersonateUser);
    
    // Format response for CLI/auth tool
    return {
      success: true,
      userEmail: impersonateUser,
      mailboxes: mailboxes.map(m => ({
        id: m.id,
        type: m.type,
        displayName: m.displayName,
        email: m.email,
        isPrimary: m.type === 'personal',
      })),
      note: mailboxes.length === 1 
        ? 'Only personal mailbox found. Delegate discovery requires MailboxSettings.Read and Calendars.Read permissions.' 
        : undefined,
    };
  } catch (error) {
    logger.error(`Error listing impersonated mailboxes: ${(error as Error).message}`);
    return {
      success: false,
      error: (error as Error).message,
    };
  }
}
```

Changes:
- **Remove**: All 4 inline discovery strategies (250+ lines)
- **Add**: Service instantiation and delegation (5 lines)
- **Keep**: Error handling and response formatting
- **Result**: Single source of truth for discovery logic

---

### Phase 6: Documentation

#### File: `IMPERSONATION.md` (NEW)

Comprehensive documentation covering:

**1. Overview**
- What impersonation is and when to use it
- Security considerations
- Relationship to Microsoft Graph permissions

**2. Configuration Guide**
- All environment variables explained
- Examples for common scenarios
- Performance tuning recommendations

**3. Source Precedence**
Clear table:
```markdown
| Source | Priority | Empty Handling | Invalid Email | User Not Found |
|--------|----------|----------------|---------------|----------------|
| HTTP Header (X-Impersonate-User) | 1 (highest) | Skip, try next | ERROR | ERROR* |
| AsyncLocalStorage Context | 2 | Skip, try next | ERROR | ERROR* |
| Environment Variable (MS365_MCP_IMPERSONATE_USER) | 3 (lowest) | ERROR | ERROR | ERROR* |

* Only if MS365_MCP_VALIDATE_IMPERSONATION_USER=true
```

**4. Validation Modes**
- **Disabled (default)**: Fast, validation happens at Graph API call time
- **Enabled**: Fail-fast with clear errors, better UX, slight performance cost

**5. Mailbox Discovery Process**
- How each strategy works
- What types of mailboxes are discovered
- Performance characteristics

**6. Caching Strategy**
- What gets cached and for how long
- Cache invalidation behavior
- Tuning TTL values

**7. Troubleshooting**
Common errors with solutions:
- "Impersonation failed: User 'x' not found"
- "No valid token found"
- "Graph API scope error"
- Empty mailbox results

**8. Examples**
Real scenarios:
- Basic impersonation setup
- Header-based dynamic impersonation
- Shared mailbox access
- Multi-tenant scenarios

#### File: `.env.example` (UPDATE)

Add section:
```bash
# ============================================
# IMPERSONATION CONFIGURATION
# ============================================

# Execute API calls as a specific user (required for impersonation)
# Format: user@domain.com
MS365_MCP_IMPERSONATE_USER=

# HTTP header name for dynamic per-request impersonation
# Default: X-Impersonate-User
MS365_MCP_IMPERSONATE_HEADER=X-Impersonate-User

# Validate that impersonated user exists before processing requests
# Disabled by default for performance (validation happens at Graph API call time)
# Enable for fail-fast behavior with clearer error messages
# Default: false
MS365_MCP_VALIDATE_IMPERSONATION_USER=false

# Cache duration for user validation results (seconds)
# Only applies when MS365_MCP_VALIDATE_IMPERSONATION_USER=true
# Default: 3600 (1 hour)
MS365_MCP_USER_VALIDATION_TTL=3600

# Cache duration for mailbox discovery results (seconds)
# Higher values = better performance, but slower to reflect permission changes
# Default: 3600 (1 hour)
MS365_MCP_IMPERSONATE_CACHE_TTL=3600
```

---

### Phase 7: Cleanup

**Delete obsolete code:**
1. Remove all references to `MS365_MCP_IMPERSONATE_ALLOWED_MAILBOXES` env var
   - Search codebase for this variable name
   - Remove from env examples, documentation, code
2. Remove duplicate discovery logic from `auth.ts`
   - Delete ~250 lines of inline discovery strategies
   - Keep only service integration code
3. Clean up verbose debug logging in `graph-tools.ts`
   - Consolidate into ImpersonationResolver
   - Keep only high-level resolution logging

**Update exports:**

File: `src/impersonation/index.ts`
```typescript
export { default as ImpersonationContext } from './ImpersonationContext.js';
export { default as MailboxDiscoveryCache } from './MailboxDiscoveryCache.js';
export { MailboxDiscoveryService } from './MailboxDiscoveryService.js';
export { UserValidationCache } from './UserValidationCache.js';
export { ImpersonationResolver } from './ImpersonationResolver.js';
export type { MailboxInfo } from './MailboxDiscoveryCache.js';
```

---

## TESTING STRATEGY

### Unit Tests (Future Enhancement)

Each service should have focused unit tests:

**MailboxDiscoveryService:**
- Each discovery strategy independently
- Error handling for Graph API failures
- Concurrent request limiting
- Timeout behavior

**UserValidationCache:**
- Email format validation (valid/invalid cases)
- Cache hit/miss behavior
- TTL expiration
- Graph API call counting

**ImpersonationResolver:**
- Precedence logic (all combinations)
- Empty value handling
- Format validation
- Error message clarity

### Integration Tests (Future Enhancement)

**Full discovery flow:**
- Real Graph API calls (requires test tenant)
- Mock Graph API responses
- Cache behavior across multiple requests

**Error scenarios:**
- Invalid user email
- User exists but no mailbox access
- Graph API errors (403, 500, timeout)
- Cache expiration during request

### Manual Test Scenarios

Document these in IMPERSONATION.md:

**Scenario 1: Basic impersonation**
```bash
MS365_MCP_IMPERSONATE_USER=alice@company.com
# Expected: Discovers alice's personal mailbox + any shared/delegated mailboxes
```

**Scenario 2: Header override**
```bash
MS365_MCP_IMPERSONATE_USER=alice@company.com
Header: X-Impersonate-User: bob@company.com
# Expected: Uses bob@ (header has highest priority)
# Log: "Impersonation resolved: bob@company.com (source: meta-header)"
```

**Scenario 3: Empty header fallback**
```bash
MS365_MCP_IMPERSONATE_USER=alice@company.com
Header: X-Impersonate-User: 
# Expected: Uses alice@ (empty header ignored, falls back to env)
# Log: "Impersonation resolved: alice@company.com (source: env-var)"
```

**Scenario 4: Validation enabled - invalid user**
```bash
MS365_MCP_VALIDATE_IMPERSONATION_USER=true
MS365_MCP_IMPERSONATE_USER=nonexistent@company.com
# Expected: ERROR with message:
# "Impersonation failed: User 'nonexistent@company.com' not found in tenant (source: environment variable)"
```

**Scenario 5: No shared mailboxes**
```bash
MS365_MCP_IMPERSONATE_USER=newuser@company.com
# Expected: Returns only personal mailbox in discovery
# Shared mailbox tools will find nothing (correct behavior)
```

**Scenario 6: Shared mailbox access**
```bash
MS365_MCP_IMPERSONATE_USER=admin@company.com
# Expected: Discovers personal mailbox + all shared mailboxes admin has access to
# Each discovered with type: 'shared' or 'delegated'
```

**Scenario 7: Cache performance**
```bash
# First request: Full discovery, logs show Graph API calls
# Second request within TTL: Instant, logs show cache hit
# Request after TTL: Full discovery again
```

---

## FILE CHANGES SUMMARY

### New Files:
1. `src/impersonation/MailboxDiscoveryService.ts` - Core discovery logic (~200 lines)
2. `src/impersonation/UserValidationCache.ts` - Optional validation (~100 lines)
3. `src/impersonation/ImpersonationResolver.ts` - Source precedence (~150 lines)
4. `IMPERSONATION.md` - Comprehensive documentation (~500 lines)

### Modified Files:
1. `src/impersonation/MailboxDiscoveryCache.ts` - Refactor to use service (~100 lines, -50 from current)
2. `src/impersonation/index.ts` - Add new exports (~3 lines added)
3. `src/graph-tools.ts` - Integration updates (~30 lines changed)
4. `src/auth.ts` - Delegate to service (~250 lines removed, ~20 lines added)
5. `.env.example` - New variables documented (~15 lines added)

### Removed:
1. All references to `MS365_MCP_IMPERSONATE_ALLOWED_MAILBOXES` env var
2. Duplicate discovery logic from `auth.ts` (~250 lines)
3. Verbose debug logging in `graph-tools.ts` (~20 lines)

### Net Change:
- **Added**: ~950 lines (mostly documentation and new services)
- **Removed**: ~320 lines (duplicated/obsolete code)
- **Net**: +630 lines
- **Code quality**: Significantly improved (single source of truth, proper separation of concerns)

---

## PERFORMANCE IMPACT

### Before:
- No actual discovery (stub returns only personal mailbox)
- Fast but incomplete results

### After (validation disabled - default):
- **First request**: 5-10 Graph API calls for full discovery (~2-5 seconds)
- **Cached requests**: Instant (cache hit)
- **Cache TTL**: 1 hour default (configurable)
- **Overall**: Initial cost, then excellent performance

### After (validation enabled):
- **First request per user**: +1 Graph API call for validation (~200ms)
- **Cached validation**: Instant (separate cache)
- **Trade-off**: Slightly slower first request, better error messages

### Recommendations:
- **Default config** (validation disabled, 1h cache): Best for most users
- **Strict validation needed**: Enable validation, keep 1h cache
- **High permission churn**: Reduce cache TTL to 15-30 minutes
- **Performance critical**: Increase cache TTL to 4-8 hours

---

## SECURITY CONSIDERATIONS

1. **Impersonation validation**: Optional user existence check prevents typos/attacks
2. **Permission enforcement**: Discovery only returns mailboxes user actually has access to
3. **No permission escalation**: Service account permissions determine what's accessible
4. **Cache security**: TTL prevents stale permission data from lingering too long
5. **Clear audit trail**: All impersonation logged with source for security review

---

## FUTURE ENHANCEMENTS

Ideas for later (not in this implementation):

1. **Proactive cache warming**: Pre-populate cache for known users
2. **WebSocket cache updates**: Real-time permission change notifications
3. **Metrics/monitoring**: Track discovery performance, cache hit rates
4. **Admin UI**: Visual mailbox permission management
5. **Batch operations**: Discover multiple users in parallel
6. **Permission diff**: Show what changed when cache expires

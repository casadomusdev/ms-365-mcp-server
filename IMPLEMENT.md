# PowerShell-Based Shared Mailbox Discovery - Implementation Plan

## GOAL

Implement accurate shared mailbox discovery for impersonated users by integrating Exchange Online PowerShell to query mailbox delegation permissions (Full Access and SendAs), replacing the current Graph API calendar delegation approach which does not work for shared mailboxes.

## ANALYSIS

### Problem Statement

The current implementation attempts to discover shared mailboxes accessible to an impersonated user by checking calendar delegation permissions via Microsoft Graph API. This approach has fundamental flaws:

**Why Graph API Calendar Delegation Fails:**
- Calendar delegation is for **user-to-user mailbox sharing**, not shared mailbox access
- Shared mailboxes (unlicensed special accounts like `info@company.com`) don't have meaningful calendar permissions
- Graph API has **NO endpoint** to query Exchange mailbox delegation permissions (Full Access, SendAs)
- The `/users/{id}/sendAs` endpoint checks OAuth2 API permissions, not Exchange mailbox permissions
- These are completely different permission systems

**What We Actually Need:**
- Detect which shared mailboxes the impersonated user has **Full Access** to
- Detect which shared mailboxes the impersonated user has **SendAs** permissions for
- Both of these are Exchange-specific permissions NOT exposed via Graph API

### Why PowerShell is the Solution

Exchange Online PowerShell provides the ONLY programmatic way to query mailbox delegation permissions:

1. **Get-MailboxPermission** - Detects Full Access permissions
2. **Get-RecipientPermission** - Detects SendAs permissions

These cmdlets directly query Exchange Online permission data that is not available through any REST API.

### Architecture Overview

**Current Flow:**
```
User Request → MailboxDiscoveryCache → MailboxDiscoveryService 
  → Graph API (calendar permissions) ❌ DOESN'T WORK
```

**New Flow:**
```
User Request → MailboxDiscoveryCache → MailboxDiscoveryService 
  → PowerShellService → Exchange Online PowerShell ✓ WORKS
```

**Key Components:**
1. **PowerShellService** - New TypeScript service to execute PowerShell commands
2. **ExchangePermissionsChecker** - PowerShell script that queries permissions
3. **MailboxDiscoveryService** - Refactored to use PowerShellService instead of Graph API
4. **MailboxDiscoveryCache** - No changes needed (implementation-agnostic)

## IMPLEMENTATION

### Phase 1: Environment & Dependencies

**PowerShell Core Requirements:**
- PowerShell 7.x installed in environment (Docker container)
- Exchange Online PowerShell V3 module
- Certificate-based authentication support

**Authentication Approach:**
- Reuse existing app credentials (Client ID, Tenant ID)
- Use certificate-based auth (more secure than client secret for PowerShell)
- Certificate can be generated and stored securely in environment

**Environment Variables:**
```bash
# Existing
MS365_CLIENT_ID=<app-id>
MS365_TENANT_ID=<tenant-id>

# New for PowerShell
MS365_CERTIFICATE_PATH=/path/to/cert.pfx  # or cert thumbprint
MS365_CERTIFICATE_PASSWORD=<cert-password>  # if using PFX
MS365_POWERSHELL_ENABLED=true  # Feature flag
MS365_POWERSHELL_TIMEOUT=30000  # Timeout in ms
```

### Phase 2: PowerShell Integration Layer

**New File: `src/lib/PowerShellService.ts`**

TypeScript service to execute PowerShell scripts:
- Spawn PowerShell Core processes
- Handle stdin/stdout communication
- Parse JSON output from PowerShell
- Timeout and error handling
- Logging

**Key Methods:**
```typescript
class PowerShellService {
  async execute(script: string, args: Record<string, any>): Promise<any>
  async checkPermissions(userEmail: string): Promise<MailboxPermission[]>
}
```

### Phase 3: Exchange PowerShell Scripts

**New File: `scripts/check-mailbox-permissions.ps1`**

PowerShell script that:
1. Connects to Exchange Online (certificate auth)
2. Gets all shared mailboxes (unlicensed users with mailboxes)
3. For each shared mailbox:
   - Checks if the user has Full Access (`Get-MailboxPermission`)
   - Checks if the user has SendAs (`Get-RecipientPermission`)
4. Returns JSON array of accessible mailboxes

**Script Output Format:**
```json
[
  {
    "id": "mailbox-guid",
    "email": "info@casadomus.tech",
    "displayName": "Info Mailbox",
    "permissions": ["fullAccess", "sendAs"]
  }
]
```

**PowerShell Commands Used:**
```powershell
# Connect
Connect-ExchangeOnline -CertificateFilePath $certPath -AppId $appId -Organization $tenantId

# Get shared mailboxes (unlicensed users)
Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited

# Check Full Access
Get-MailboxPermission -Identity $mailbox | Where-Object {
  $_.User -eq $userEmail -and $_.AccessRights -contains 'FullAccess'
}

# Check SendAs
Get-RecipientPermission -Identity $mailbox | Where-Object {
  $_.Trustee -eq $userEmail -and $_.AccessRights -contains 'SendAs'
}
```

### Phase 4: Refactor MailboxDiscoveryService

**Changes to `src/impersonation/MailboxDiscoveryService.ts`:**

1. **Add PowerShellService dependency**
2. **Remove calendar delegation methods:**
   - `checkCalendarDelegation()`
   - `checkCalendarPermissions()`
3. **Add PowerShell-based discovery:**
   - `discoverViaExchangePowerShell()`
4. **Update `discoverMailboxes()` method:**
   - Keep personal mailbox detection (Graph API)
   - Replace calendar checking with PowerShell permission checking
5. **Error handling:**
   - Fallback if PowerShell not available
   - Clear error messages for configuration issues

**Pseudo-code:**
```typescript
async discoverMailboxes(userEmail: string): Promise<MailboxInfo[]> {
  // 1. Get personal mailbox (Graph API - works fine)
  const personalMailbox = await this.getPersonalMailbox(userEmail);
  
  // 2. Discover shared mailboxes (PowerShell)
  const sharedMailboxes = await this.powerShellService.checkPermissions(userEmail);
  
  // 3. Combine and return
  return [personalMailbox, ...sharedMailboxes];
}
```

### Phase 5: Documentation Updates

**Files to Update:**
1. **README.md** - Add PowerShell requirements section
2. **SERVER_SETUP.md** - PowerShell Core installation instructions
3. **USER_IMPERSONATE.md** - Update shared mailbox detection explanation
4. **Dockerfile** - Add PowerShell Core and Exchange module installation
5. **.env.example** - Add new PowerShell-related environment variables

**New Documentation:**
- **POWERSHELL_SETUP.md** - Detailed PowerShell configuration guide
  - Installing PowerShell Core
  - Installing Exchange Online module
  - Certificate generation and configuration
  - Troubleshooting

### Phase 6: Testing Strategy

**Unit Tests:**
- PowerShellService command execution
- JSON parsing from PowerShell output
- Error handling and timeouts

**Integration Tests:**
- End-to-end permission checking
- Cache integration
- Fallback scenarios

**Manual Testing Scenarios:**
1. User with no shared mailbox access → Only personal mailbox
2. User with Full Access to shared mailbox → Personal + shared
3. User with SendAs only → Personal + shared
4. User with both permissions → Personal + shared
5. PowerShell unavailable → Graceful degradation with clear error
6. PowerShell timeout → Error handling

### Phase 7: Deployment Considerations

**Docker Image:**
- Base image must support PowerShell Core
- Exchange Online PowerShell module pre-installed
- Certificate handling (volume mount or env variable)

**Performance:**
- First discovery will be slow (~2-5 seconds for PowerShell)
- Cache is critical (TTL-based, already implemented)
- Consider background refresh for active users

**Security:**
- Certificate must be securely stored
- Certificate password (if PFX) in secrets management
- Minimal permissions on app registration

## TECHNICAL NOTES

### Graph API vs PowerShell Comparison

| Feature | Graph API | PowerShell |
|---------|-----------|------------|
| Personal mailbox | ✅ Works | ✅ Works |
| Calendar delegation | ✅ Works (user-to-user) | ✅ Works |
| Shared mailbox Full Access | ❌ No endpoint | ✅ Get-MailboxPermission |
| Shared mailbox SendAs | ❌ No endpoint | ✅ Get-RecipientPermission |
| Performance | Fast (~100ms) | Slow (~2-5s) |
| Auth | App token | Certificate required |

### Why We Keep Some Graph API

**Graph API is still used for:**
- Personal mailbox detection (fast, works well)
- User validation (checking user exists)
- Future: User-to-user calendar delegation (if needed)

**PowerShell is only used for:**
- Shared mailbox Full Access detection
- Shared mailbox SendAs detection

### Alternative Approaches Considered

1. **Configuration File** - Too manual, requires constant updates
2. **Attempt Access Pattern** - Doesn't work with client credentials + impersonation
3. **EWS API** - Deprecated, poor modern auth support, no permission endpoints
4. **Admin SDK** - Doesn't exist for Exchange permissions

PowerShell is the **only viable solution**.

## FUTURE IMPROVEMENTS

1. **Performance Optimization:**
   - Background cache warming
   - Parallel PowerShell queries
   - Incremental permission checks

2. **Enhanced Features:**
   - Real-time permission change detection
   - Audit logging for permission queries
   - Permission type filtering (Full Access vs SendAs)

3. **Alternative Auth:**
   - Managed Identity support (Azure)
   - Azure Key Vault integration
   - Certificate rotation automation

4. **Monitoring:**
   - PowerShell execution time tracking
   - Permission query success rate
   - Cache hit/miss ratios
